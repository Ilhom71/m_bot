"""Integration tests use synthetic data and never send real messages."""
import asyncio
import json
import logging
import os
import sqlite3
import subprocess
import sys
import tempfile
import time
import unittest
from dataclasses import replace
from contextlib import closing
from zipfile import ZipFile
from xml.etree import ElementTree
from datetime import datetime
from io import BytesIO
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, patch
from zoneinfo import ZoneInfo

import httpx
import sms
from openpyxl import Workbook, load_workbook
from telegram.error import BadRequest, Forbidden, RetryAfter, TimedOut

from config import Settings, load_demo_settings, load_settings
from db import Database
from excel_utils import ValidationError, appeals_workbook, build_workbook, import_notes, in_shift, parse_excel, phone_number
from handlers.bot import (ADMIN_MENU, USER_MENU, BUTTON_ACTIONS, automatic_schedule, send_now, command_actions, document, export_employees,
                          join_callback, join_contact, my_appeals, new_appeal, status, text_message, templates)
from main import InstanceLock, RedactingFormatter, build_application, main as run_main
from scheduler import DeliveryService
from sms import SMSNotConfigured, SMSRetryableError

TZ = ZoneInfo("Asia/Tashkent")


def workbook_bytes(headers, rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()


def employee(name="Test Worker", ref="123456789", phone="+998900000001", row=2):
    return dict(name=name, telegram_ref=ref, phone=phone, shift="", row_number=row)


def appeal(appeal_id="TEST-1", due=None, ref="Test Worker"):
    return dict(id=appeal_id, client_name="Test Client", phone="+998900000002", body="Synthetic test only",
                due_at=due, employee_ref=ref, row_number=2)


class InstanceLockTests(unittest.TestCase):
    def test_empty_file_and_release(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "bot.lock"
            first = InstanceLock(path)
            with first:
                self.assertFalse(first.handle.closed)
            self.assertTrue(first.handle.closed)
            with InstanceLock(path):
                pass

    def test_legacy_file_is_not_read_or_modified(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "bot.lock"
            path.write_bytes(b"0")
            with InstanceLock(path):
                pass
            self.assertEqual(path.read_bytes(), b"0")

    def test_second_process_gets_readable_error(self):
        code = (
            "import sys; from pathlib import Path; from main import InstanceLock\n"
            "try:\n"
            "    with InstanceLock(Path(sys.argv[1])): pass\n"
            "except RuntimeError as exc:\n"
            "    print(str(exc)); sys.exit(7)\n"
        )
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "bot.lock"
            path.write_bytes(b"0")
            with InstanceLock(path):
                result = subprocess.run([sys.executable, "-c", code, str(path)],
                                        capture_output=True, text=True, timeout=30)
            self.assertEqual(result.returncode, 7, result.stderr)
            self.assertIn("boshqa nusxasi", result.stdout)
            self.assertNotIn("PermissionError", result.stderr)


class ExcelTests(unittest.TestCase):
    def test_uzbek_phone_normalization(self):
        for value in ('90 123 45 67', '901234567', '998901234567', '+998 (90) 123-45-67'):
            self.assertEqual(phone_number(value), '+998901234567')
        with self.assertRaises(ValueError):
            phone_number('+15551234567')

    def test_optional_extra_phone_excel(self):
        for columns, values, expected in [
            (['ism','telefon'], ['Worker','901234567'], None),
            (['ism','telefon','qoshimcha_telefon'], ['Worker','901234567',None], None),
            (['ism','telefon','qoshimcha_telefon'], ['Worker','901234567','91 123 45 67'], '+998911234567'),
            (['ism','telefon','qoshimcha_telefon'], ['Worker','901234567','998901234567'], None),
        ]:
            row = parse_excel(workbook_bytes(columns, [values]), 'employees', TZ)[0]
            self.assertEqual(row['phone_extra'], expected)

    def test_excel_allows_shared_primary_and_extra_phones(self):
        for first, second in [(['A','901234567','911234567'], ['B','911234567',None]),
                              (['A','901234567','911234567'], ['B','931234567','911234567']),
                              (['A','901234567',None], ['B','931234567','901234567']),
                              (['A','901234567',None], ['B','901234567',None])]:
            self.assertEqual(len(parse_excel(workbook_bytes(['ism','telefon','qoshimcha_telefon'], [first, second]),
                                             'employees', TZ)), 2)
        with self.assertRaisesRegex(ValidationError, '3-qator.*Telegram'):
            payload = workbook_bytes(['ism', 'telefon', 'telegram_id'], [['A', '901234567', '123'], ['B', '901234567', '123']])
            parse_excel(payload, 'employees', TZ)

    def test_workbook_without_optional_dimension(self):
        original = workbook_bytes(["ism", "telefon", "telegram_id"], [["Test Worker", "+998900000001", None]])
        stream = BytesIO()
        with ZipFile(BytesIO(original)) as source, ZipFile(stream, "w") as target:
            for entry in source.infolist():
                data = source.read(entry.filename)
                if entry.filename == "xl/worksheets/sheet1.xml":
                    root = ElementTree.fromstring(data)
                    dimension = root.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}dimension")
                    if dimension is not None:
                        root.remove(dimension)
                    data = ElementTree.tostring(root)
                target.writestr(entry, data)
        self.assertEqual(len(parse_excel(stream.getvalue(), "employees", TZ)), 1)

    def test_valid_and_blank_schedule(self):
        payload = workbook_bytes(["id", "mijoz_ismi", "telefon", "murojaat_matni", "yuborish_vaqti"],
                                 [["A1", "Test", "900000002", "Test", None]])
        row = parse_excel(payload, "appeals", TZ)[0]
        self.assertIsNone(row["due_at"])
        self.assertEqual(row["phone"], "+998900000002")

    def test_excel_datetime_and_username(self):
        payload = workbook_bytes(["id", "mijoz_ismi", "telefon", "murojaat_matni", "yuborish_vaqti"],
                                 [["A1", "Test", "900000002", "Test", datetime(2027, 1, 1, 10)]])
        self.assertEqual(datetime.fromtimestamp(parse_excel(payload, "appeals", TZ)[0]["due_at"], TZ).hour, 10)
        payload = workbook_bytes(["ism", "telefon", "telegram_id", "smena_vaqti"],
                                 [["Test", "900000001", "@testworker", "22:00-06:00"]])
        self.assertEqual(parse_excel(payload, "employees", TZ)[0]["telegram_ref"], "@testworker")

    def test_phone_export_filename_changes_the_phone_column_meaning(self):
        headers = ["id", "mijoz_ismi", "telefon", "murojaat_matni", "yuborish_vaqti", "xodim"]
        filename = "murojaatlar_bot_murojatlar asil_20261124_120154.xlsx"
        payload = workbook_bytes(headers, [["TEST/1", "Test Client", "900000002", "Test body", None, "Test Worker"]])
        row = parse_excel(payload, "appeals", TZ, filename=filename)[0]
        self.assertEqual(row["phone"], "Ko'rsatilmagan")
        self.assertEqual(row["employee_ref"], "phone:+998900000002|Test Worker")
        standard = parse_excel(payload, "appeals", TZ, filename="murojaatlar.xlsx")[0]
        self.assertEqual(standard["phone"], "+998900000002")
        self.assertEqual(standard["employee_ref"], "Test Worker")
        blank = workbook_bytes(headers, [["TEST/1", "Test Client", None, "Test body", None, "Test Worker"]])
        with self.assertRaisesRegex(ValidationError, "2-qator: telefon .*xodim"):
            parse_excel(blank, "appeals", TZ, filename=filename)

    def test_missing_columns_and_row_error(self):
        with self.assertRaisesRegex(ValidationError, "1-qator"):
            parse_excel(workbook_bytes(["ism"], [["Test"]]), "employees", TZ)
        payload = workbook_bytes(["ism", "telefon", "telegram_id"], [["Test", "bad", 123]])
        with self.assertRaisesRegex(ValidationError, "2-qator.*telefon"):
            parse_excel(payload, "employees", TZ)

    def test_duplicate_formula_and_corruption(self):
        for rows in ([["Test", "900000001", 123], ["Test", "900000002", 456]],
                     [["=1+1", "900000001", 123]]):
            with self.assertRaises(ValidationError):
                parse_excel(workbook_bytes(["ism", "telefon", "telegram_id"], rows), "employees", TZ)
        with self.assertRaises(ValidationError):
            parse_excel(b"not a zip", "employees", TZ)

    def test_overnight_shift(self):
        self.assertTrue(in_shift("22:00-06:00", datetime(2027, 1, 1, 23)))
        self.assertTrue(in_shift("22:00-06:00", datetime(2027, 1, 1, 5)))
        self.assertFalse(in_shift("22:00-06:00", datetime(2027, 1, 1, 6)))

    def test_legacy_header_and_deadline_not_schedule(self):
        rows = [["Test title"]] * 7 + [["Кириш рақами", "Аризачи", "Ҳужжат мазмуни", "Кимга юборилган", "Муддати"],
                ["12/TEST", "Test Client", "Test body", "Test Worker", datetime(2027, 1, 1)]]
        payload = workbook_bytes([None], rows)
        parsed = parse_excel(payload, "appeals", TZ)[0]
        self.assertTrue(parsed["id"].startswith("L-"))
        self.assertIn("12/TEST", parsed["body"])
        self.assertIsNone(parsed["due_at"])
        self.assertEqual(parsed["phone"], "Ko'rsatilmagan")
        self.assertEqual(parsed["row_number"], 10)

    def test_legacy_long_and_missing_applicants_preserve_source(self):
        long_name, long_body = "Test applicant " * 60, "Synthetic document " * 600
        payload = workbook_bytes([None], [["Title"]] * 7 + [
            ["Кириш рақами", "Аризачи", "Ҳужжат мазмуни", "Кимга юборилган", "Муддати"],
            ["TEST/1", long_name, long_body, "Unknown Employee", "2027-01-01"],
            ["TEST/2", None, None, None, None]])
        rows = parse_excel(payload, "appeals", TZ)
        self.assertEqual(rows[0]["client_name"], long_name.strip())
        self.assertIn(long_body.strip(), rows[0]["body"])
        self.assertEqual(rows[0]["source_data"]["муддати"], "2027-01-01")
        self.assertEqual(rows[1]["client_name"], "Ko'rsatilmagan")
        self.assertIn("ko'rsatilmagan", rows[1]["body"])

    def test_legacy_headers_on_first_row_and_blank_telegram(self):
        payload = workbook_bytes(["Кириш рақами", "Аризачи", "Ҳужжат мазмуни", "Кимга юборилган"],
                                 [["T1", "Test", "Test", "Worker"]])
        self.assertTrue(parse_excel(payload, "appeals", TZ)[0]["legacy"])
        payload = workbook_bytes(["ism", "telefon"], [["Worker", "+998900000003"]])
        self.assertEqual(parse_excel(payload, "employees", TZ)[0]["telegram_ref"], "phone:+998900000003")

    def test_iiv_imports_first_sheet_even_when_second_is_active(self):
        workbook = Workbook()
        for index, number_header in enumerate(("Ариза рақами", "Мурожаат рақами"), 1):
            sheet = workbook.active if index == 1 else workbook.create_sheet()
            sheet.append(["Synthetic report"])
            sheet.append(["№", number_header, "Мурожаат муаллифининг\nФ.И.Ш.", "Мурожат таснифи",
                          "Яшаш манзили ёки юридик манзили", "Муддат", "ижрочи", "Хозирги муддати"])
            sheet.append([index, f"IIV/{index}", "Test applicant", "Test body", "Unused address",
                          datetime(2027, 2, 1), "Test Worker", "2027-02-02"])
        workbook.active = 1
        stream = BytesIO()
        workbook.save(stream)

        rows = parse_excel(stream.getvalue(), "appeals", TZ)
        self.assertEqual(len(rows), 1)
        self.assertTrue(all(row["id"].startswith("IIV-") and row["legacy"] for row in rows))
        self.assertEqual(rows[0]["body"], "Murojaat raqami: IIV/1\nTest body")
        self.assertEqual(rows[0]["employee_ref"], "Test Worker")
        self.assertEqual(rows[0]["phone"], "Ko'rsatilmagan")
        self.assertIsNone(rows[0]["due_at"])
        self.assertEqual(rows[0]["source_data"], {"муддат": "2027-02-01 00:00:00", "хозирги муддати": "2027-02-02"})
        self.assertEqual(rows[0]["row_location"], "1-varaq, 3-qator")
        self.assertIn("Муддат", import_notes(rows))

    def test_pp100_imports_only_the_first_sheet(self):
        workbook = Workbook()
        for index, serial_header in enumerate(("№", None), 1):
            sheet = workbook.active if index == 1 else workbook.create_sheet()
            sheet.append(["Synthetic report"])
            sheet.append([serial_header, "Ариза рақами", "Фамиля", "Исми", "Мурожат таснифи", "Ижрочи"])
            sheet.append([index, f"PP/{index}", "Test surname", "Test name", "Test body", None])
        hidden = workbook.create_sheet()
        hidden.sheet_state = "hidden"
        hidden.append(["Ариза рақами", "Фамиля", "Исми", "Мурожат таснифи"])
        hidden.append(["PP/2", "Test surname", "Test name", "Test body"])
        workbook.active = 1
        stream = BytesIO()
        workbook.save(stream)

        rows = parse_excel(stream.getvalue(), "appeals", TZ)
        self.assertEqual([row["client_name"] for row in rows], ["Test surname Test name"])
        self.assertTrue(rows[0]["id"].startswith("PP100-"))
        self.assertIn("PP/1", rows[0]["body"])
        self.assertEqual(rows[0]["row_location"], "1-varaq, 3-qator")
        self.assertEqual(rows[0]["employee_ref"], "Test surname Test name")
        exported = appeals_workbook([dict(row, employee_hint=row["employee_ref"],
                                          source_json=json.dumps(row["source_data"])) for row in rows], TZ)
        reimported = parse_excel(exported.getvalue(), "appeals", TZ)
        self.assertEqual([row["id"] for row in rows], [row["id"] for row in reimported])
        self.assertEqual(reimported[0]["employee_ref"], rows[0]["employee_ref"])

        workbook.create_sheet(index=0)
        workbook.active = 1
        stream = BytesIO()
        workbook.save(stream)
        with self.assertRaisesRegex(ValidationError, "shablon sarlavhasi topilmadi"):
            parse_excel(stream.getvalue(), "appeals", TZ)

    def test_pp100_skips_empty_appeal_rows_and_reads_the_next_appeal(self):
        headers = ["№", "Ариза рақами", "Фамиля", "Исми", "Мурожат таснифи", "Муддат"]
        rows = [[1, "PP/1", "Test surname", "Test name", "Only appeal"],
                [2, "  ", None, " ", None, "2027-02-01"],
                [3, "PP/2", "Test surname", "Test name", "Below the gap"]]
        for empty_row in (rows[1], [None] * len(headers)):
            with self.subTest(empty_row=empty_row):
                payload = workbook_bytes(headers, [rows[0], empty_row, rows[2]])
                result = parse_excel(payload, "appeals", TZ)
                self.assertEqual([row["body"] for row in result],
                                 ["Murojaat raqami: PP/1\nOnly appeal", "Murojaat raqami: PP/2\nBelow the gap"])
        with self.assertRaisesRegex(ValidationError, "ma'lumot qatorlari yo'q"):
            parse_excel(workbook_bytes(headers, [rows[1]]), "appeals", TZ)

    def test_pp100_loads_hidden_appeals_from_first_sheet(self):
        workbook = Workbook()
        filtered = workbook.active
        other_sheet = workbook.create_sheet()
        workbook.active = other_sheet
        header = ["№", "Ариза рақами", "Фамиля", "Исми", "Мурожат таснифи"]
        other_sheet.append(header)
        other_sheet.append([1, "PP/other", "Test surname", "Test name", "Other sheet"])
        filtered.append(["Synthetic PP100 report"])
        filtered.append(header)
        filtered.auto_filter.ref = "A2:E5"
        filtered.append([1, "PP/1", "Test surname", "Test name", "Hidden by filter"])
        filtered.row_dimensions[3].hidden = True
        filtered.append([2])
        filtered.append([3, "PP/2", "Test surname", "Test name", "Visible appeal"])
        stream = BytesIO()
        workbook.save(stream)

        rows = parse_excel(stream.getvalue(), "appeals", TZ)
        self.assertEqual([row["body"] for row in rows],
                         ["Murojaat raqami: PP/1\nHidden by filter", "Murojaat raqami: PP/2\nVisible appeal"])

    def test_pp100_skips_duplicates_in_first_sheet(self):
        workbook = Workbook()
        first, second = workbook.active, workbook.create_sheet()
        header = ["Ариза рақами", "Фамиля", "Исми", "Мурожат таснифи"]
        first.append(header)
        first.append(["PP/1", "Test surname", "Test name", "First body"])
        first.append(["PP/2", "Test surname", "Test name", "Unique body"])
        first.append(["PP/1", "Test surname", "Test name", "First duplicate"])
        first.append(["PP/1", "Test surname", "Test name", "Second duplicate"])
        second.append(header)
        second.append(["PP/3", "Test surname", "Test name", "Second sheet"])
        workbook.active = 1
        stream = BytesIO()
        workbook.save(stream)

        rows = parse_excel(stream.getvalue(), "appeals", TZ)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["body"], "Murojaat raqami: PP/1\nFirst body")
        self.assertEqual(rows[1]["body"], "Murojaat raqami: PP/2\nUnique body")
        self.assertIn("2 ta takroriy qator", import_notes(rows))

    def test_iiv_duplicate_is_reported_with_sheet_and_row(self):
        payload = workbook_bytes(["Мурожаат рақами", "Мурожаат муаллифининг Ф.И.Ш.", "Мурожат таснифи"],
                                 [["IIV/1", "Test name", "Test body"]] * 2)
        with self.assertRaisesRegex(ValidationError, "1-varaq, 3-qator: .*takrorlangan"):
            parse_excel(payload, "appeals", TZ)

    def test_export_strings_cannot_become_formulas(self):
        stream = build_workbook(["name"], [["=1+1"], ["+SUM(A1)"], ["@test"]])
        workbook = load_workbook(stream)
        self.assertEqual(workbook.active["A2"].data_type, "s")
        self.assertEqual(workbook.active["A2"].value, "=1+1")
        workbook.close()


class BotTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        path = Path(self.temp.name)
        self.settings = Settings("123:test", frozenset(), TZ, 9, path / "bot.db", path / "jobs.db")
        self.db = Database(self.settings.database)
        self.db.import_employees([employee()])
        request_id = self.db.request_join(1, 123456789)
        self.db.decide_join(request_id, True, 987)
        self.bot = SimpleNamespace(username="synthetic_test_bot", send_message=AsyncMock(return_value=SimpleNamespace(message_id=42)),
                                   send_document=AsyncMock(return_value=SimpleNamespace(message_id=43)))
        self.sms = AsyncMock(return_value="provider-test-id")
        self.service = DeliveryService(self.bot, self.db, self.settings, self.sms)

    def tearDown(self):
        self.temp.cleanup()

    def add_due(self, appeal_id="TEST-1", ref="Test Worker"):
        self.db.import_appeals([appeal(appeal_id, time.time()-10, ref)])

    def channel(self, channel="telegram", appeal_id="TEST-1"):
        return next(row for row in self.db.details(appeal_id)[1] if row["channel"] == channel)

    async def test_legacy_sender_without_applicant_appears_in_telegram(self):
        payload = workbook_bytes(
            ['Кириш рақами', 'Ҳужжат мазмуни', 'Кимга юборилган', 'Юборган   ижрочи'],
            [['SYNTHETIC-1', 'Synthetic text', 'Test Worker', 'Synthetic Sender']],
        )
        rows = parse_excel(payload, 'appeals', TZ)
        self.assertEqual(rows[0]['employee_ref'], 'Test Worker')
        self.db.import_appeals(rows)
        self.db.set_time(rows[0]['id'], time.time()-10)
        await self.service.tick()
        message = self.bot.send_message.await_args.kwargs['text']
        self.assertIn('Юборган ижрочи: Synthetic Sender', message)
        self.assertNotIn('Mijoz:', message)

    async def test_demo_starts_empty_with_separate_database_and_no_sms_transport(self):
        with patch('config.ROOT', self.settings.database.parent), patch.dict(os.environ, {
                'DEMO_BOT_TOKEN': '123:synthetic-demo', 'DEMO_ADMIN_ID': '987'}, clear=True), \
                patch('config.load_dotenv') as dotenv_reader, patch('config.dotenv_values') as dotenv_values_reader:
            settings = load_demo_settings()
        dotenv_reader.assert_not_called()
        dotenv_values_reader.assert_not_called()
        app = build_application(settings)
        self.assertNotEqual(settings.database, self.settings.database)
        self.assertEqual(app.bot_data['db'].employee_list(), [])
        self.assertEqual(app.bot_data['db'].list_appeals(TZ), [])
        with self.assertRaises(SMSNotConfigured):
            await app.bot_data['service'].sms_sender('+998900000001', 'Synthetic')

    async def test_fetched_deliveries_respect_rescheduled_time_and_edits(self):
        self.add_due()
        telegram, sms_item = self.db.ready_deliveries(time.time())
        future = time.time() + 3600
        self.db.set_time('TEST-1', future)
        await self.service.deliver(telegram)
        await self.service.deliver(sms_item)
        self.bot.send_message.assert_not_awaited()
        self.sms.assert_not_awaited()

        edited = dict(appeal(due=time.time()-10), body='Edited synthetic text')
        self.db.import_appeals([edited])
        await self.service.deliver(telegram)
        self.bot.send_message.assert_not_awaited()
        await self.service.tick()
        self.assertIn('Edited synthetic text', self.bot.send_message.await_args.kwargs['text'])

    async def test_sms_queue_scan_keeps_unchanged_confirmation_snapshot(self):
        self.add_due()
        self.db.ready_deliveries(time.time())
        before = self.db.details('TEST-1')
        self.db.ready_deliveries(time.time())
        self.assertEqual(self.db.details('TEST-1'), before)

    async def test_send_once_across_restarts_and_reimport(self):
        item = appeal(due=time.time()-10)
        self.db.import_appeals([item])
        await self.service.tick()
        self.assertEqual(self.db.import_appeals([item]), (0, 1))
        restarted = DeliveryService(self.bot, Database(self.settings.database), self.settings, self.sms)
        await restarted.tick()
        self.assertEqual(self.bot.send_message.await_count, 1)
        self.assertEqual(self.sms.await_count, 1)
        self.assertEqual(self.channel()["status"], "sent")

    async def test_sms_placeholder_not_sent_and_telegram_independent(self):
        self.add_due()
        self.sms.side_effect = SMSNotConfigured()
        await self.service.tick()
        self.assertEqual(self.channel()["status"], "sent")
        self.assertEqual(self.channel("sms")["status"], "blocked")
        self.assertEqual(self.channel("sms")["attempts"], 0)

    async def test_sms_switch_holds_old_sms_and_keeps_future_schedule(self):
        self.db.import_employees([dict(employee(), phone_extra='+998900000003')])
        self.db.decide_join(self.db.request_join(1, 123456789), True, 987)
        self.add_due()
        self.db.set_sms_enabled(False)
        restarted_db = Database(self.settings.database)
        restarted = DeliveryService(self.bot, restarted_db, self.settings, self.sms)
        self.assertFalse(restarted_db.sms_enabled())
        await restarted.tick()
        self.bot.send_message.assert_awaited_once()
        self.sms.assert_not_awaited()
        self.assertEqual(self.channel('sms')['attempts'], 0)
        self.assertEqual(self.channel()['status'], 'sent')
        restarted_db.set_sms_enabled(True)
        restarted_db.set_sms_enabled(True)
        restarted_db = Database(self.settings.database)
        restarted = DeliveryService(self.bot, restarted_db, self.settings, self.sms)
        await restarted.tick()
        self.sms.assert_not_awaited()
        self.bot.send_message.assert_awaited_once()

        due_at = time.time() + 60
        self.db.import_appeals([appeal('FUTURE', due_at)])
        self.db.set_sms_enabled(False)
        self.db.set_sms_enabled(True)
        with patch('time.time', return_value=due_at + 1):
            await restarted.tick()
        self.assertEqual(self.channel('sms', 'FUTURE')['status'], 'sent')
        self.assertEqual(self.channel('sms')['status'], 'pending')
        self.assertEqual(self.sms.await_count, 2)

        restarted_db.retry('TEST-1', 'sms')
        await restarted.tick()
        self.assertEqual(self.sms.await_count, 4)
        self.assertEqual({call.args[0] for call in self.sms.await_args_list}, {'+998900000001', '+998900000003'})
        await restarted.tick()
        self.assertEqual(self.sms.await_count, 4)
        self.assertEqual(self.bot.send_message.await_count, 2)

    async def test_disabling_sms_blocks_previously_fetched_deliveries(self):
        self.add_due()
        self.assertTrue(self.db.sms_enabled())
        sms_item = next(item for item in self.db.ready_deliveries(time.time()) if item['channel'] == 'sms')
        self.db.set_sms_enabled(False)
        await self.service.deliver(sms_item)
        self.sms.assert_not_awaited()
        self.assertEqual(self.db.sms_details('TEST-1')[0]['attempts'], 0)
        self.assertFalse(self.db.claim('TEST-1', 'sms', time.time()))
        self.db.set_sms_enabled(True)
        await self.service.deliver(sms_item)
        self.sms.assert_not_awaited()
        await self.service.tick()
        self.sms.assert_not_awaited()

    async def test_sms_password_lockout_survives_restart_and_rotation(self):
        original = "synthetic first password"
        replacement = "synthetic replacement password"
        self.db.set_sms_password(original)
        for _ in range(4):
            self.assertEqual(self.db.change_sms_with_password(False, "synthetic wrong password", 987), "invalid")
        self.assertEqual(self.db.change_sms_with_password(False, "synthetic wrong password", 987), "locked")
        restarted = Database(self.settings.database)
        self.assertEqual(restarted.change_sms_with_password(False, original, 987), "locked")
        self.assertTrue(restarted.sms_enabled())
        with patch("db.time.time", return_value=time.time() + 301):
            self.assertEqual(restarted.change_sms_with_password(False, original, 987), "changed")
        self.assertFalse(restarted.sms_enabled())
        restarted.set_sms_password(replacement)
        self.assertEqual(restarted.change_sms_with_password(True, original, 987), "invalid")
        self.assertEqual(restarted.change_sms_with_password(True, replacement, 987), "changed")
        self.assertTrue(restarted.sms_enabled())

    async def test_upgrade_holds_backlog_when_sms_was_already_enabled(self):
        self.add_due()
        self.db.set_sms_enabled(True)
        with self.db.connect() as conn:
            conn.execute("ALTER TABLE appeals DROP COLUMN sms_held")

        upgraded_db = Database(self.settings.database)
        upgraded = DeliveryService(self.bot, upgraded_db, self.settings, self.sms)
        await upgraded.tick()
        self.sms.assert_not_awaited()
        self.assertEqual(self.channel()['status'], 'sent')

    async def test_disabled_sms_does_not_generate_pending_sms_reminders(self):
        self.add_due()
        self.db.set_sms_enabled(False)
        await self.service.tick()
        settings = replace(self.settings, admin_ids=frozenset({987}), reminder_hour=0)
        service = DeliveryService(self.bot, self.db, settings, self.sms)
        service.prepare_reminders()
        notices = self.db.ready_notices(time.time())
        self.assertFalse(any(row['notice_key'].startswith('attention:') for row in notices))
        self.db.set_sms_enabled(True)
        service.prepare_reminders()
        self.assertFalse(any(row['notice_key'].startswith('attention:') for row in self.db.ready_notices(time.time())))

    async def test_sms_is_short_and_links_to_the_assigned_appeal(self):
        item = appeal("LONG.SYNTHETIC-ID", time.time()-10)
        item['body'] = 'Synthetic long appeal content. ' * 1000
        self.db.import_appeals([item])
        await self.service.tick()
        phone, text = self.sms.await_args.args
        self.assertEqual(phone, '+998900000001')
        self.assertEqual(text, 'Sizga yangi vazifa biriktirildi. Batafsil: '
                         'https://t.me/synthetic_test_bot?start=task_1')
        self.assertTrue(text.isascii())
        self.assertLessEqual(len(text), 160)
        self.assertNotIn(item['client_name'], text)
        self.assertNotIn(item['phone'], text)
        self.assertNotIn('Synthetic long appeal content.', text)
        parts = [call.kwargs['text'] for call in self.bot.send_message.await_args_list]
        self.assertGreater(len(parts), 1)
        content = [part.split('\n', 1)[1] for part in parts]
        content[0] = content[0].split('\n\n', 1)[1]
        self.assertEqual(''.join(''.join(content).split()), ''.join(item['body'].split()))

    async def test_textup_results_are_persisted_without_repeating_sms(self):
        cases = [(httpx.Response(200, json={'smsId': 'synthetic-sms-id'}), 'sent'),
                 (httpx.Response(400, text='private-provider-response'), 'failed'),
                 (httpx.Response(500, text='private-provider-response'), 'unknown'),
                 (httpx.ReadTimeout('private-transport-error'), 'unknown'),
                 (httpx.Response(401), 'blocked')]
        for index, (response, expected) in enumerate(cases):
            with self.subTest(status=expected, index=index):
                requests = []
                login = httpx.Response(200, json={'accessToken': 'synthetic-token',
                    'user': {'id': '00000000-0000-4000-8000-000000000001'}})
                responses = [response] if expected == 'blocked' else [login, response]
                def handle(request):
                    requests.append(request)
                    result = responses.pop(0)
                    if isinstance(result, Exception):
                        raise result
                    return result
                provider = sms.TextUPClient('synthetic@example.invalid', 'synthetic-password',
                                            transport=httpx.MockTransport(handle))
                self.service.sms_sender = provider.send
                identifier = f'TEXTUP-{index}'
                self.add_due(identifier)
                await self.service.tick()
                result = self.db.sms_details(identifier)[0]
                self.assertEqual(result['status'], expected)
                self.assertEqual(result['attempts'], 0 if expected == 'blocked' else 1)
                self.assertNotIn('private-', result['last_error'] or '')
                if expected == 'sent':
                    self.assertEqual(result['message_id'], 'synthetic-sms-id')
                self.assertEqual(self.channel('telegram', identifier)['status'], 'sent')
                request_count = len(requests)
                await self.service.tick()
                self.assertEqual(len(requests), request_count)

    async def test_initial_attempt_plus_three_retries(self):
        self.add_due()
        self.bot.send_message.side_effect = Forbidden("synthetic rejection")
        for _ in range(5):
            with self.db.connect() as conn:
                conn.execute("UPDATE deliveries SET next_attempt=0")
            await self.service.tick()
        self.assertEqual(self.bot.send_message.await_count, 4)
        self.assertEqual(self.channel()["attempts"], 4)
        self.assertEqual(self.channel()["status"], "failed")
        self.assertEqual(self.sms.await_count, 1)

    async def test_assignment_summary_counts_unique_employees_and_does_not_repeat(self):
        settings = replace(self.settings, admin_ids=frozenset({987}))
        service = DeliveryService(self.bot, self.db, settings, self.sms)
        self.db.import_employees([employee('Pending Worker', '222', '+998900000003')])
        missing = appeal('MISSING', time.time()-10, 'Missing Worker')
        missing['legacy'] = True
        self.db.import_appeals([appeal('PENDING', time.time()-10, 'Pending Worker'),
                                appeal('PENDING-2', time.time()-10, 'Pending Worker'), missing])
        service.prepare_reminders()
        service.prepare_reminders()
        notices = [r for r in self.db.ready_notices(time.time()) if r['notice_key'].startswith('assignment')]
        self.assertEqual(len(notices), 1)
        text = notices[0]['text']
        self.assertIn('Botga ulanmagan xodimlar: 1 ta', text)
        self.assertIn('Ulanishni kutayotgan murojaatlar: 2 ta', text)
        self.assertIn('Mos xodim topilmagan murojaatlar: 1 ta', text)
        self.assertIn('?start=assignment_status', text)
        self.assertNotIn('PENDING', text)
        self.assertIn('Ulanmagan xodimga biriktirilgan murojaatlar: 2', self.db.status_text())
        self.assertIn('Mos xodim topilmagan murojaatlar: 1', self.db.status_text())

    async def test_restart_retires_queued_per_appeal_assignment_notices(self):
        self.db.enqueue_notice('assignment:TEST:unconnected:old:0', [987], 'Old assignment reminder')
        self.db.enqueue_notice('daily:2026-01-01:0', [987], 'Daily reminder')
        self.db.recover()
        notices = self.db.ready_notices(time.time())
        self.assertEqual([notice['text'] for notice in notices], ['Daily reminder'])

    async def test_delivery_timeouts_stop_hangs_and_never_resend_unknown(self):
        async def hanging_send(*args, **kwargs):
            await asyncio.Future()
        self.add_due()
        self.bot.send_message.side_effect = hanging_send
        with patch('scheduler.SEND_TIMEOUT', 0.01):
            await self.service.tick()
        self.assertEqual(self.channel()['status'], 'unknown')
        self.assertEqual(self.channel('sms')['status'], 'sent')
        await self.service.tick()
        self.assertEqual(self.bot.send_message.await_count, 1)
        self.bot.send_message.side_effect = None
        self.sms.side_effect = hanging_send
        self.add_due('SMS-TIMEOUT')
        with patch('scheduler.SEND_TIMEOUT', 0.01):
            await self.service.tick()
        self.assertEqual(self.channel('sms', 'SMS-TIMEOUT')['status'], 'unknown')
        await self.service.tick()
        self.assertEqual(self.sms.await_count, 2)

    async def test_stale_sending_watchdog_preserves_fresh_claims(self):
        self.add_due()
        self.add_due('FRESH')
        now = time.time()
        self.db.ready_deliveries(now)
        self.assertTrue(self.db.claim('TEST-1', 'telegram', now))
        self.assertTrue(self.db.claim_sms('TEST-1', '+998900000001', now))
        self.assertTrue(self.db.claim('FRESH', 'telegram', now))
        self.db.enqueue_notice('stale-notice', [987], 'Synthetic notice')
        self.db.claim_notice('stale-notice:987')
        with self.db.connect() as conn:
            conn.execute("UPDATE deliveries SET updated_at=? WHERE appeal_id='TEST-1'", (now-121,))
            conn.execute("UPDATE sms_deliveries SET updated_at=? WHERE appeal_id='TEST-1'", (now-121,))
            conn.execute("UPDATE notices SET updated_at=? WHERE status='sending'", (now-121,))
        self.db.expire_sending(now)
        self.assertEqual(self.channel()['status'], 'unknown')
        self.assertEqual(self.channel('sms')['status'], 'unknown')
        self.assertEqual(self.channel(appeal_id='FRESH')['status'], 'sending')
        with self.db.connect() as conn:
            self.assertEqual(conn.execute('SELECT status FROM notices').fetchone()[0], 'unknown')

    async def test_duplicate_id_error_is_explicit_and_import_is_atomic(self):
        first, second = appeal('DUPLICATE'), appeal('DUPLICATE')
        second['row_number'] = 3
        with self.assertRaises(ValidationError) as caught:
            self.db.import_appeals([first, second])
        self.assertIn("ID 'DUPLICATE'", str(caught.exception))
        self.assertIn('3-qator', str(caught.exception))
        self.assertIsNone(self.db.details('DUPLICATE'))
        self.add_due()
        await self.service.tick()
        changed = appeal()
        changed['body'] = 'Changed synthetic text'
        with self.assertRaises(ValidationError) as caught:
            self.db.import_appeals([appeal('NEW-ROLLBACK'), changed])
        self.assertIn("ID 'TEST-1' mavjud", str(caught.exception))
        self.assertIn('Yangi murojaatga yangi ID', str(caught.exception))
        self.assertIsNone(self.db.details('NEW-ROLLBACK'))

    async def test_telegram_alias_and_numeric_id_cannot_share_employees(self):
        self.db.register_user(123456789, 'syntheticworker')
        with self.assertRaises(ValidationError):
            self.db.import_employees([employee('Other', '@syntheticworker', '+998900000003')])
        self.db.import_employees([employee('Other', '222', '+998900000003')])
        with self.assertRaises(ValueError):
            self.db.request_join(2, 123456789)

    async def test_completion_is_owned_idempotent_and_notifies_admin(self):
        self.add_due()
        with self.assertRaises(ValueError):
            self.db.complete_appeal('TEST-1', 555, [987])
        self.assertTrue(self.db.complete_appeal('TEST-1', 123456789, [987]))
        completed_at = self.db.details('TEST-1')[0]['completed_at']
        self.assertFalse(self.db.complete_appeal('TEST-1', 123456789, [987]))
        self.assertEqual(self.db.details('TEST-1')[0]['completed_at'], completed_at)
        self.assertEqual(len(self.db.ready_notices(time.time())), 1)
        self.assertIn('Bajarildi: 1', self.db.status_text())
        await self.service.tick()
        self.sms.assert_not_awaited()
        self.bot.send_message.assert_awaited_once()
        self.assertEqual(self.bot.send_message.await_args.kwargs['chat_id'], 987)
        self.assertIn('Bajarildi', self.bot.send_message.await_args.kwargs['text'])
        with self.assertRaises(ValueError):
            self.db.retry('TEST-1', 'telegram')
        self.db.import_appeals([appeal('FUTURE-COMPLETION', time.time()+3600)])
        with self.assertRaises(ValueError):
            self.db.complete_appeal('FUTURE-COMPLETION', 123456789, [987])

    async def test_disconnect_revokes_access_and_allows_verified_rejoin(self):
        self.add_due()
        self.db.disconnect_employee(1)
        worker = self.db.employee_list()[0]
        self.assertEqual(worker['approved'], 0)
        self.assertIsNone(worker['telegram_id'])
        self.assertEqual(self.db.export_appeals(telegram_id=123456789), [])
        self.assertEqual(self.db.ready_deliveries(time.time()), [])
        notices = self.db.ready_notices(time.time())
        self.assertEqual(notices[0]['admin_id'], 123456789)
        request = self.db.join_by_phone('90 000 00 01', 123456789)
        self.db.decide_join(request, True, 987)
        self.assertEqual(self.db.employee_list()[0]['approved'], 1)

    async def test_daily_sqlite_backups_include_wal_and_scheduler_database(self):
        self.add_due()
        with closing(sqlite3.connect(self.settings.jobs_database)) as conn:
            conn.execute('CREATE TABLE synthetic_jobs(value TEXT)')
            conn.execute("INSERT INTO synthetic_jobs VALUES ('persisted')")
            conn.commit()
        now = datetime(2027, 1, 1, 23, 55, tzinfo=TZ).timestamp()
        with self.db.connect() as conn:
            conn.execute('CREATE TABLE synthetic_wal(value TEXT)')
            conn.execute("INSERT INTO synthetic_wal VALUES ('committed WAL record')")
            conn.commit()
            self.service.backup_databases(now)
        files = sorted((self.settings.database.parent/'backups').glob('*.sqlite3'))
        self.assertEqual(len(files), 2)
        self.assertIn('2027-01-01', files[0].name)
        with closing(sqlite3.connect(files[0])) as conn:
            self.assertEqual(conn.execute('PRAGMA quick_check').fetchone()[0], 'ok')
            self.assertEqual(conn.execute('SELECT value FROM synthetic_wal').fetchone()[0], 'committed WAL record')
        with closing(sqlite3.connect(files[1])) as conn:
            self.assertEqual(conn.execute('SELECT value FROM synthetic_jobs').fetchone()[0], 'persisted')
        original = [path.read_bytes() for path in files]
        self.service.backup_databases(now+60)
        self.assertEqual([path.read_bytes() for path in files], original)
        self.service.backup_databases(now+600)
        self.assertEqual(len(list(files[0].parent.glob('*.sqlite3'))), 4)

    async def test_backup_failure_does_not_prevent_delivery(self):
        service = DeliveryService(self.bot, self.db, replace(self.settings, admin_ids=frozenset({987})), self.sms)
        self.add_due()
        with patch.object(Database, 'create_backup', side_effect=OSError('synthetic-private-error')):
            await service.tick()
        self.assertEqual(self.channel()['status'], 'sent')
        self.assertEqual(self.channel('sms')['status'], 'sent')
        notices = self.db.ready_notices(time.time())
        self.assertTrue(any('backup' in r['text'] for r in notices))
        self.assertFalse(any('synthetic-private-error' in r['text'] for r in notices))

    async def test_network_ambiguity_not_blindly_retried(self):
        self.add_due()
        self.bot.send_message.side_effect = TimedOut()
        await self.service.tick()
        await self.service.tick()
        self.assertEqual(self.channel()["status"], "unknown")
        self.assertEqual(self.bot.send_message.await_count, 1)
        with self.assertRaises(ValueError):
            self.db.retry("TEST-1", "telegram")
        self.db.resolve_unknown("TEST-1", "telegram", True)
        self.assertEqual(self.channel()["status"], "sent")

    async def test_crash_recovery_and_atomic_claim(self):
        self.add_due()
        self.assertTrue(self.db.claim("TEST-1", "telegram", time.time()))
        self.assertFalse(self.db.claim("TEST-1", "telegram", time.time()))
        Database(self.settings.database).recover()
        self.assertEqual(self.channel()["status"], "unknown")

    async def test_username_binds_before_and_after_import(self):
        self.db.import_employees([employee(ref="@testworker")])
        self.add_due()
        await self.service.tick()
        self.assertEqual(self.bot.send_message.await_count, 0)
        self.assertEqual(self.channel()["attempts"], 0)
        self.assertEqual(self.db.register_user(999, "TestWorker"), [])
        request_id = self.db.join_by_phone("+998900000001", 999)
        self.db.decide_join(request_id, True, 987)
        with self.db.connect() as conn:
            conn.execute("UPDATE deliveries SET next_attempt=0")
        await self.service.tick()
        self.assertEqual(self.bot.send_message.call_args.kwargs["chat_id"], 999)
        self.assertEqual(self.db.register_user(111, "testworker"), [])
        self.db.import_employees([employee(ref="@testworker")])
        self.assertEqual(self.db.employee_list()[0]["telegram_id"], 999)

    async def test_atomic_import_rollback(self):
        with self.assertRaises(ValidationError):
            self.db.import_appeals([appeal("GOOD", time.time()), appeal("BAD", time.time(), "Missing")])
        self.assertIsNone(self.db.details("GOOD"))
        with self.assertRaises(ValidationError):
            self.db.import_employees([employee("Other", "222", "+998900000003"), employee("Third", "222", "+998900000004", 3)])
        self.assertEqual(len(self.db.employee_list()), 1)

    async def test_unscheduled_future_and_assignment(self):
        self.db.import_appeals([appeal("BLANK"), appeal("FUTURE", time.time()+3600, "")])
        await self.service.tick()
        self.assertEqual(self.bot.send_message.await_count, 0)
        self.assertIsNotNone(self.db.details("FUTURE")[0]["employee_id"])
        self.db.set_time("BLANK", time.time()-10)
        await self.service.tick()
        self.assertEqual(self.bot.send_message.await_count, 1)
        with self.assertRaises(ValueError):
            self.db.set_time("BLANK", time.time()+3600)

    async def test_persistent_scheduler_and_overdue_restart(self):
        self.add_due()
        self.service.start()
        await asyncio.sleep(0.15)
        self.assertEqual(self.channel()["status"], "sent")
        await self.service.stop()
        self.assertTrue(self.settings.jobs_database.exists())
        self.add_due("TEST-2")
        restarted = DeliveryService(self.bot, Database(self.settings.database), self.settings, self.sms)
        restarted.start()
        await asyncio.sleep(0.15)
        await restarted.stop()
        self.assertEqual(self.channel(appeal_id="TEST-2")["status"], "sent")
        self.assertEqual(self.bot.send_message.await_count, 2)

    async def test_admin_only_guard(self):
        message = SimpleNamespace(reply_text=AsyncMock())
        update = SimpleNamespace(effective_chat=SimpleNamespace(type="private"),
                                 effective_user=SimpleNamespace(id=555), effective_message=message)
        context = SimpleNamespace(bot_data={"settings": self.settings})
        await status(update, context)
        self.assertIn("faqat adminlar", message.reply_text.call_args.args[0])

    async def test_reminder_deduplication(self):
        settings = Settings("123:test", frozenset({987}), TZ, 0, self.settings.database, self.settings.jobs_database)
        service = DeliveryService(self.bot, self.db, settings, self.sms)
        service.prepare_reminders()
        service.prepare_reminders()
        await service.send_notices()
        await service.send_notices()
        self.assertEqual(self.bot.send_message.await_count, 1)

    async def test_bad_request_and_rate_limit_are_retryable(self):
        self.add_due()
        self.bot.send_message.side_effect = BadRequest("synthetic rejection")
        await self.service.tick()
        self.assertEqual(self.channel()["status"], "retry")
        with self.db.connect() as conn:
            conn.execute("UPDATE deliveries SET next_attempt=0")
        self.bot.send_message.side_effect = RetryAfter(300)
        await self.service.tick()
        self.assertEqual(self.channel()["status"], "retry")
        self.assertGreater(self.channel()["next_attempt"], time.time()+290)

    async def test_public_form_saved_and_admin_notice_queued(self):
        settings = Settings("123:test", frozenset({987}), TZ, 9, self.settings.database, self.settings.jobs_database)
        service = DeliveryService(self.bot, self.db, settings, self.sms)
        message = SimpleNamespace(reply_text=AsyncMock(), text="")
        update = SimpleNamespace(effective_chat=SimpleNamespace(type="private"),
                                 effective_user=SimpleNamespace(id=555), effective_message=message)
        context = SimpleNamespace(bot_data={"settings": settings, "db": self.db, "service": service}, user_data={})
        await new_appeal(update, context)
        for value in ("Test Client", "+998900000002", "Synthetic body"):
            message.text = value
            await text_message(update, context)
        rows = self.db.list_appeals(TZ)
        self.assertEqual(len(rows), 1)
        self.assertIsNone(rows[0]["due_at"])
        self.assertEqual(len(self.db.ready_notices(time.time())), 1)
        self.assertNotIn("form", context.user_data)

    async def test_upload_and_templates(self):
        settings = Settings("123:test", frozenset({987}), TZ, 9, self.settings.database, self.settings.jobs_database)
        payload = workbook_bytes(["id", "mijoz_ismi", "telefon", "murojaat_matni", "yuborish_vaqti"],
                                 [["A1", "Test", "900000002", "Test", None]])
        remote = SimpleNamespace(download_as_bytearray=AsyncMock(return_value=bytearray(payload)))
        doc = SimpleNamespace(file_name="test.xlsx", file_size=len(payload), get_file=AsyncMock(return_value=remote))
        message = SimpleNamespace(reply_text=AsyncMock(), reply_document=AsyncMock(), document=doc)
        update = SimpleNamespace(effective_chat=SimpleNamespace(type="private"),
                                 effective_user=SimpleNamespace(id=987), effective_message=message)
        context = SimpleNamespace(bot_data={"settings": settings, "db": self.db}, user_data={"upload": "appeals"})
        await document(update, context)
        self.assertIsNotNone(self.db.details("A1"))
        self.assertNotIn("upload", context.user_data)
        await templates(update, context)
        self.assertEqual(message.reply_document.await_count, 2)

    async def test_admin_upload_rejected_before_download(self):
        doc = SimpleNamespace(get_file=AsyncMock())
        message = SimpleNamespace(reply_text=AsyncMock(), document=doc)
        update = SimpleNamespace(effective_chat=SimpleNamespace(type="private"),
                                 effective_user=SimpleNamespace(id=555), effective_message=message)
        context = SimpleNamespace(bot_data={"settings": self.settings}, user_data={"upload": "appeals"})
        await document(update, context)
        doc.get_file.assert_not_awaited()

    async def test_unapproved_employee_receives_neither_channel(self):
        self.db.import_employees([employee("New Worker", "222", "+998900000003")])
        self.add_due("NEW", "New Worker")
        await self.service.tick()
        self.bot.send_message.assert_not_awaited()
        self.sms.assert_not_awaited()
        self.assertEqual(self.db.register_user(222, None), [])
        request_id = self.db.join_by_phone("+998900000003", 222)
        self.assertEqual(self.db.join_by_phone("+998900000003", 222), request_id)
        await self.service.tick()
        self.bot.send_message.assert_not_awaited()
        self.db.decide_join(request_id, True, 987)
        await self.service.tick()
        self.assertEqual(self.bot.send_message.call_args.kwargs["chat_id"], 222)
        self.assertEqual(self.sms.await_count, 1)
        self.assertIn("Murojaat #NEW", self.bot.send_message.call_args.kwargs["text"])
        self.assertIn("Synthetic test only", self.bot.send_message.call_args.kwargs["text"])
        self.assertNotIn("document", self.bot.send_message.call_args.kwargs)

    async def test_legacy_unresolved_employee_is_not_randomly_assigned(self):
        item = appeal(due=time.time()-10, ref="Missing Worker")
        item.update(legacy=True, source_data={"аризачи": "Test", "муддати": "2027-01-01"})
        self.db.import_appeals([item])
        await self.service.tick()
        self.assertIsNone(self.db.details(item["id"])[0]["employee_id"])
        self.bot.send_message.assert_not_awaited()
        self.db.import_employees([employee("Missing Worker", "333", "+998900000003")])
        self.db.assign_due(time.time(), TZ)
        self.assertIsNotNone(self.db.details(item["id"])[0]["employee_id"])
        self.assertEqual(len(self.db.ready_deliveries(time.time())), 0)

    async def test_portal_upload_queues_only_unique_employee_matches(self):
        self.db.import_employees([
            employee("TestWorker", "333", "+998900000003"),
            employee("Other Employee", "444", "+998900000004"),
            employee("Pending Employee", "555", "+998900000005"),
        ])
        request_id = self.db.request_join(3, 444)
        self.db.decide_join(request_id, True, 987)
        payload = workbook_bytes(
            ["Ариза рақами", "Фамиля", "Исми", "Мурожат таснифи", "ижрочи"],
            [
                ["PP/1", "OTHER  ", "  employee", "Combined name", None],
                ["PP/2", "OTHER", "employee", "Explicit executor", " tESt\u00a0wOrker "],
                ["PP/3", "Test", "Worker", "Ambiguous name", None],
                ["PP/4", "Other", "Employee", "Unmatched executor", "Unknown employee"],
                ["PP/5", None, None, "Missing name", None],
                ["PP/6", "Other", "Employee", "Approved employee", "123456789"],
                ["PP/7", "Pending", "Employee", "Unapproved employee", None],
            ],
        )
        rows = parse_excel(payload, "appeals", TZ)
        self.assertEqual(self.db.import_appeals(rows), (7, 0))
        now = time.time()
        self.assertEqual(self.db.set_default_schedule("darhol", TZ, now), 3)
        self.db.assign_due(now, TZ)
        stored = [self.db.details(row["id"])[0] for row in rows]
        self.assertEqual(stored[0]["employee_id"], 3)
        self.assertTrue(all(row["employee_id"] is None and row["due_at"] is None for row in stored[1:5]))
        self.assertEqual(stored[5]["employee_id"], 1)
        self.assertEqual(stored[6]["employee_id"], 4)
        ready = self.db.ready_deliveries(now)
        self.assertEqual({(row["appeal_id"], row["telegram_id"]) for row in ready},
                         {(rows[0]["id"], 444), (rows[5]["id"], 123456789)})

        iiv_payload = workbook_bytes(
            ["Ариза рақами", "Мурожаат муаллифининг Ф.И.Ш.", "Мурожат таснифи"],
            [["IIV/1", "Test Worker", "Applicant name alone"]],
        )
        iiv_row = parse_excel(iiv_payload, "appeals", TZ)[0]
        self.db.import_appeals([iiv_row])
        self.assertEqual(self.db.apply_default_schedule(TZ, now), 0)
        self.db.assign_due(now, TZ)
        self.assertIsNone(self.db.details(iiv_row["id"])[0]["employee_id"])

    async def test_excel_employee_phone_uses_primary_and_extra_and_holds_ambiguous_rows(self):
        shared = "+998900000001"
        unique = "+998900000003"
        self.db.import_employees([dict(employee("Other Worker", "444", unique), phone_extra=shared)])
        request = self.db.join_by_phone(unique, 444)
        self.db.decide_join(request, True, 987)
        headers = ["id", "mijoz_ismi", "murojaat_matni", "yuborish_vaqti", "xodim", "xodim_telefoni"]
        payload = workbook_bytes(headers, [
            ["UNIQUE", "Client", "Test body", None, None, "90 000 00 03"],
            ["SHARED", "Client", "Test body", None, "TEST   worker", shared],
            ["AMBIGUOUS", "Client", "Test body", None, None, shared],
            ["UNKNOWN", "Client", "Test body", None, None, "+998900000099"],
            ["CONFLICT", "Client", "Test body", None, "Test Worker", unique],
            ["SHORTCUT", "Client", "Test body", None, unique, None],
        ])
        rows = parse_excel(payload, "appeals", TZ)
        self.assertTrue(all(row["phone"] == "Ko'rsatilmagan" for row in rows))
        self.assertEqual(self.db.import_appeals(rows), (6, 0))
        exported = appeals_workbook(self.db.export_appeals(), TZ)
        self.assertEqual(self.db.import_appeals(parse_excel(exported.getvalue(), "appeals", TZ)), (0, 6))
        now = time.time()
        self.db.set_default_schedule("darhol", TZ, now)
        self.db.assign_due(now, TZ)
        ready = self.db.ready_deliveries(now)
        self.assertEqual({(row["appeal_id"], row["telegram_id"]) for row in ready},
                         {("UNIQUE", 444), ("SHARED", 123456789), ("SHORTCUT", 444)})
        self.assertTrue(all(self.db.details(key)[0]["employee_id"] is None
                            for key in ("AMBIGUOUS", "UNKNOWN", "CONFLICT")))

    async def test_standard_excel_keeps_slash_id_on_import_and_export(self):
        identifier = "L-TEST/2026-00001"
        headers = ["id", "mijoz_ismi", "telefon", "murojaat_matni", "yuborish_vaqti"]
        payload = workbook_bytes(headers, [[identifier, "Test Client", "900000002", "Test body", None]])
        rows = parse_excel(payload, "appeals", TZ)
        self.assertEqual(self.db.import_appeals(rows), (1, 0))
        self.assertEqual(self.db.details(identifier)[0]["id"], identifier)
        exported = appeals_workbook(self.db.export_appeals(), TZ)
        reimported = parse_excel(exported.getvalue(), "appeals", TZ)
        self.assertEqual(reimported[0]["id"], identifier)
        self.assertEqual(self.db.import_appeals(reimported), (0, 1))

    async def test_excel_edit_roundtrip_and_sent_immutability(self):
        item = appeal()
        item.update(legacy=True, phone="Ko'rsatilmagan", client_name="Test " * 100,
                    body="Body " * 1000, source_data={"муддати": "2027-01-01"})
        self.db.import_appeals([item])
        workbook = load_workbook(appeals_workbook(self.db.export_appeals(), TZ))
        workbook.active["D2"] = "Corrected text"
        stream = BytesIO()
        workbook.save(stream)
        workbook.close()
        rows = parse_excel(stream.getvalue(), "appeals", TZ)
        self.assertEqual(self.db.import_appeals(rows), (1, 0))
        row = self.db.details(item["id"])[0]
        self.assertEqual(row["body"], "Corrected text")
        self.assertEqual(json.loads(row["source_json"])["муддати"], "2027-01-01")
        self.assertEqual(self.db.import_appeals(rows), (0, 1))
        self.db.set_time(item["id"], time.time()-1)
        await self.service.tick()
        rows[0]["body"] = "Unsafe edit after send"
        with self.assertRaises(ValidationError):
            self.db.import_appeals(rows)
        self.assertEqual(self.db.details(item["id"])[0]["body"], "Corrected text")

    async def test_own_download_excludes_others_future_and_unapproved(self):
        self.add_due("OWN")
        self.db.import_appeals([appeal("FUTURE", time.time()+600)])
        self.db.import_employees([employee("Other", "222", "+998900000003")])
        self.add_due("OTHER", "Other")
        self.assertEqual([r["id"] for r in self.db.export_appeals(telegram_id=123456789)], ["OWN"])
        self.assertEqual(self.db.export_appeals(telegram_id=222), [])
        self.assertEqual(self.db.export_appeals(telegram_id=555), [])
        message = SimpleNamespace(reply_text=AsyncMock(), reply_document=AsyncMock())
        update = SimpleNamespace(effective_user=SimpleNamespace(id=555), effective_message=message)
        context = SimpleNamespace(bot_data={"db": self.db, "settings": self.settings})
        await my_appeals(update, context)
        message.reply_document.assert_not_awaited()

    async def test_changed_contacts_revoke_approval_and_invalidate_old_buttons(self):
        self.assertEqual(self.db.register_user(123456789, None), [])
        self.db.import_employees([employee(phone="+998900000004")])
        self.assertEqual(self.db.employee_list()[0]["approved"], 0)
        new_request = self.db.join_by_phone("+998900000004", 123456789)
        self.assertGreater(new_request, 1)
        with self.assertRaises(ValueError):
            self.db.decide_join(1, True, 987)
        self.db.decide_join(new_request, True, 987)
        self.assertEqual(self.db.employee_list()[0]["approved"], 1)

    async def test_reject_and_double_approval_are_not_replayed(self):
        self.db.import_employees([employee("Other", "222", "+998900000003")])
        request_id = self.db.join_by_phone("+998900000003", 222)
        self.db.decide_join(request_id, False, 987)
        with self.assertRaises(ValueError):
            self.db.decide_join(request_id, True, 987)
        self.assertEqual(self.db.export_appeals(telegram_id=222), [])

    async def test_button_command_parity_and_time_wizard(self):
        self.assertEqual(set(command_actions().values()), set(BUTTON_ACTIONS.values()))
        for keyboard in (ADMIN_MENU, USER_MENU):
            for line in keyboard.keyboard:
                for button in line:
                    self.assertIn(button.text, BUTTON_ACTIONS)
        self.db.import_appeals([appeal()])
        settings = Settings("123:test", frozenset({987}), TZ, 9, self.settings.database, self.settings.jobs_database)
        message = SimpleNamespace(reply_text=AsyncMock(), text="")
        update = SimpleNamespace(effective_chat=SimpleNamespace(type="private"),
                                 effective_user=SimpleNamespace(id=987), effective_message=message)
        context = SimpleNamespace(bot_data={"db": self.db, "settings": settings}, user_data={}, args=[])
        for value in ("Vaqt belgilash", "TEST-1", "2027-01-01 10:30"):
            message.text = value
            await text_message(update, context)
        self.assertIsNotNone(self.db.details("TEST-1")[0]["due_at"])
        self.assertNotIn("input", context.user_data)

    async def test_non_admin_cannot_approve_callback(self):
        query = SimpleNamespace(answer=AsyncMock(), edit_message_text=AsyncMock(), data="join:yes:1")
        update = SimpleNamespace(callback_query=query, effective_chat=SimpleNamespace(type="private"), effective_user=SimpleNamespace(id=555))
        context = SimpleNamespace(bot_data={"settings": self.settings})
        await join_callback(update, context)
        self.assertTrue(query.answer.call_args.kwargs["show_alert"])
        query.edit_message_text.assert_not_awaited()

    async def test_contact_join_requires_own_contact_and_admin_approval(self):
        self.db.import_employees([employee("Other", "phone:+998900000003", "+998900000003")])
        message = SimpleNamespace(reply_text=AsyncMock(), contact=SimpleNamespace(user_id=111, phone_number="998900000003"))
        update = SimpleNamespace(effective_user=SimpleNamespace(id=222, username=None), effective_message=message)
        context = SimpleNamespace(bot_data={"db": self.db, "settings": self.settings, "service": self.service}, user_data={"join": True})
        await join_contact(update, context)
        self.assertEqual(self.db.pending_joins(), [])
        message.contact.user_id = 222
        await join_contact(update, context)
        self.assertEqual(len(self.db.pending_joins()), 1)
        self.assertEqual(self.db.employee_list()[0]["approved"], 0)

    async def test_employee_download_can_be_reimported_without_approving(self):
        self.db.import_employees([employee("Other", "phone:+998900000003", "+998900000003")])
        settings = Settings("123:test", frozenset({987}), TZ, 9, self.settings.database, self.settings.jobs_database)
        message = SimpleNamespace(reply_text=AsyncMock(), reply_document=AsyncMock())
        update = SimpleNamespace(effective_user=SimpleNamespace(id=987), effective_chat=SimpleNamespace(type="private"), effective_message=message)
        context = SimpleNamespace(bot_data={"db": self.db, "settings": settings})
        await export_employees(update, context)
        stream = message.reply_document.call_args.kwargs["document"]
        rows = parse_excel(stream.getvalue(), "employees", TZ)
        self.db.import_employees(rows)
        states = {row["name"]: row["approved"] for row in self.db.employee_list()}
        self.assertEqual(states, {"Other": 0, "Test Worker": 1})

    async def test_admin_can_add_worker_from_button(self):
        settings = Settings("123:test", frozenset({987}), TZ, 9, self.settings.database, self.settings.jobs_database)
        message = SimpleNamespace(reply_text=AsyncMock(), text="")
        update = SimpleNamespace(effective_chat=SimpleNamespace(type="private"), effective_user=SimpleNamespace(id=987), effective_message=message)
        context = SimpleNamespace(bot_data={"settings": settings, "db": self.db}, user_data={}, args=[])
        for value in ("Xodim qo‘shish", "Yangi Xodim", "+998900000099", "o‘tkazish", "o‘tkazish", "-"):
            message.text = value
            await text_message(update, context)
        added = next(row for row in self.db.employee_list() if row["name"] == "Yangi Xodim")
        self.assertEqual(added["phone"], "+998900000099")
        self.assertEqual(added["approved"], 0)

    async def test_username_alone_cannot_create_employee_join_request(self):
        from handlers.bot import start
        self.db.import_employees([employee(ref="@testworker")])
        settings = Settings("123:test", frozenset(), TZ, 9, self.settings.database, self.settings.jobs_database)
        message = SimpleNamespace(reply_text=AsyncMock())
        update = SimpleNamespace(effective_chat=SimpleNamespace(type="private"), effective_user=SimpleNamespace(id=777, username="testworker"), effective_message=message)
        context = SimpleNamespace(bot_data={"db": self.db, "settings": settings}, user_data={})
        await start(update, context)
        self.assertEqual(self.db.pending_joins(), [])
        self.assertEqual(self.db.employee_list()[0]["approved"], 0)

    async def test_old_database_additive_migration_preserves_records(self):
        path = Path(self.temp.name) / "old.db"
        with sqlite3.connect(path) as conn:
            conn.executescript("""CREATE TABLE employees (id INTEGER PRIMARY KEY,name TEXT NOT NULL,name_key TEXT UNIQUE NOT NULL,
                phone TEXT UNIQUE NOT NULL,telegram_ref TEXT UNIQUE NOT NULL,telegram_id INTEGER,shift TEXT NOT NULL DEFAULT '');
                CREATE TABLE appeals (id TEXT PRIMARY KEY,client_name TEXT NOT NULL,phone TEXT NOT NULL,body TEXT NOT NULL,
                due_at REAL,employee_id INTEGER REFERENCES employees(id),created_at REAL NOT NULL,source_user INTEGER);
                INSERT INTO employees VALUES (1,'Test','test','+998900000003','222',222,'');
                INSERT INTO appeals VALUES ('OLD','Test','+998900000004','Original',NULL,1,0,NULL);""")
        conn.close()
        migrated = Database(path)
        self.assertEqual(migrated.details("OLD")[0]["body"], "Original")
        self.assertEqual(migrated.employee_list()[0]["approved"], 0)
        self.assertIsNone(migrated.employee_list()[0]["phone_extra"])
        Database(path)

    async def test_sms_both_tracks_each_phone_without_repeating_success(self):
        row = employee()
        row['phone_extra'] = '+998900000003'
        self.db.import_employees([row])
        request_id = self.db.join_by_phone(row['phone_extra'], 123456789)
        self.db.decide_join(request_id, True, 987)
        self.add_due()
        self.sms.side_effect = ['primary-id', SMSRetryableError('synthetic rejection'), 'extra-id']
        await self.service.tick()
        states = {r['phone']: r for r in self.db.sms_details('TEST-1')}
        self.assertEqual(states[row['phone']]['status'], 'sent')
        self.assertEqual(states[row['phone_extra']]['status'], 'retry')
        with self.db.connect() as conn:
            conn.execute("UPDATE sms_deliveries SET next_attempt=0")
        await self.service.tick()
        self.assertEqual([call.args[0] for call in self.sms.await_args_list], [row['phone'], row['phone_extra'], row['phone_extra']])
        self.assertEqual(self.channel('sms')['status'], 'sent')
        self.bot.send_message.assert_awaited_once()
        await DeliveryService(self.bot, Database(self.settings.database), self.settings, self.sms).tick()
        self.assertEqual(self.sms.await_count, 3)

    async def test_sms_primary_and_unconfigured_both(self):
        row = employee()
        row['phone_extra'] = '+998900000003'
        self.db.import_employees([row])
        self.db.decide_join(self.db.join_by_phone(row['phone_extra'], 123456789), True, 987)
        self.add_due('PRIMARY')
        service = DeliveryService(self.bot, self.db, replace(self.settings, sms_target='primary'), self.sms)
        self.sms.side_effect = SMSNotConfigured()
        await service.tick()
        self.assertEqual([r['phone'] for r in self.db.sms_details('PRIMARY')], [row['phone']])
        self.add_due('BOTH')
        await self.service.tick()
        states = self.db.sms_details('BOTH')
        self.assertEqual(len(states), 2)
        self.assertTrue(all(r['status']=='blocked' and r['attempts']==0 for r in states))
        self.assertEqual(self.channel('sms', 'BOTH')['status'], 'blocked')

    async def test_shared_phone_sends_both_employees_their_own_appeals(self):
        shared_phone = '+998900000003'
        first = employee('First', '333', shared_phone)
        second = dict(employee('Second', '@second_worker', '+998900000004'), phone_extra=shared_phone)
        self.db.import_employees([first, second])
        with self.assertRaisesRegex(ValueError, 'bir nechta xodimda'):
            self.db.join_by_phone(shared_phone, 555)
        self.assertEqual(self.db.pending_joins(), [])
        for telegram_id, username in ((333, None), (444, 'SECOND_WORKER')):
            request = self.db.join_by_phone(shared_phone, telegram_id, username)
            self.db.decide_join(request, True, 987)
        self.assertIsNone(self.db.join_by_phone(shared_phone, 444, 'renamed_worker'))

        now = time.time() - 5
        self.db.import_appeals([appeal('FIRST', now, 'First'), appeal('SECOND', now, 'Second')])
        await self.service.tick()
        telegram_tasks = {call.kwargs['chat_id']: call.kwargs['text'] for call in self.bot.send_message.await_args_list}
        self.assertIn('#FIRST', telegram_tasks[333])
        self.assertIn('#SECOND', telegram_tasks[444])
        self.assertEqual([call.args[0] for call in self.sms.await_args_list].count(shared_phone), 2)
        self.assertEqual(self.channel('sms', 'FIRST')['status'], 'sent')
        self.assertEqual(self.channel('sms', 'SECOND')['status'], 'sent')

    def test_shared_phone_schema_upgrade_keeps_employee_links(self):
        path = self.settings.database.with_name("old-phones.db")
        with closing(sqlite3.connect(path)) as connection, connection:
            connection.executescript("""
                CREATE TABLE employees (
                    id INTEGER PRIMARY KEY, name TEXT NOT NULL, name_key TEXT UNIQUE NOT NULL,
                    phone TEXT UNIQUE NOT NULL, telegram_ref TEXT UNIQUE NOT NULL, telegram_id INTEGER,
                    shift TEXT NOT NULL DEFAULT '', approved INTEGER NOT NULL DEFAULT 0, phone_extra TEXT
                );
                CREATE TABLE appeals (
                    id TEXT PRIMARY KEY, client_name TEXT NOT NULL, phone TEXT NOT NULL,
                    body TEXT NOT NULL, due_at REAL, employee_id INTEGER REFERENCES employees(id),
                    created_at REAL NOT NULL, source_user INTEGER
                );
                INSERT INTO employees VALUES (7,'First','first','+998900000003','phone:+998900000003',333,'',1,NULL);
                INSERT INTO appeals VALUES ('OLD','Test','+998900000009','Test',NULL,7,1,NULL);
            """)
        upgraded = Database(path)
        upgraded.import_employees([employee('Second', 'phone:+998900000003', '+998900000003')])
        self.assertEqual(upgraded.details('OLD')[0]['employee_id'], 7)
        self.assertIsNone(upgraded.join_by_phone('+998900000003', 333))
        with upgraded.connect() as connection:
            self.assertIsNone(connection.execute('PRAGMA foreign_key_check').fetchone())
            self.assertEqual(connection.execute('SELECT COUNT(*) FROM employees WHERE approved=1').fetchone()[0], 1)

    async def test_sms_manual_retry_only_retries_failed_recipient(self):
        row = employee()
        row['phone_extra'] = '+998900000003'
        self.db.import_employees([row])
        self.db.decide_join(self.db.join_by_phone(row['phone_extra'], 123456789), True, 987)
        self.add_due()
        self.sms.side_effect = ['ok', SMSRetryableError()]
        await self.service.tick()
        with self.db.connect() as conn:
            conn.execute("UPDATE sms_deliveries SET status='failed',attempts=4 WHERE status='retry'")
        self.db.retry('TEST-1', 'sms')
        self.sms.side_effect = None
        self.sms.return_value = 'extra-ok'
        await self.service.tick()
        self.assertEqual([c.args[0] for c in self.sms.await_args_list], [row['phone'], row['phone_extra'], row['phone_extra']])

    async def test_migration_does_not_resend_historical_sms(self):
        self.add_due()
        with self.db.connect() as conn:
            conn.execute("UPDATE deliveries SET status='sent',attempts=1,message_id='old-id' WHERE channel='sms'")
        await self.service.tick()
        self.sms.assert_not_awaited()
        self.assertEqual(self.db.sms_details('TEST-1')[0]['message_id'], 'old-id')

    async def test_pending_sms_contacts_refresh_only_after_approval(self):
        self.add_due()
        self.sms.side_effect = SMSNotConfigured()
        await self.service.tick()
        row = employee()
        row['phone_extra'] = '+998900000003'
        self.db.import_employees([row])
        before = self.sms.await_count
        await self.service.tick()
        self.assertEqual(self.sms.await_count, before)
        self.db.decide_join(self.db.join_by_phone(row['phone_extra'],123456789), True,987)
        await self.service.tick()
        self.assertEqual(len(self.db.sms_details('TEST-1')), 2)

    async def test_extra_change_revokes_approval_and_same_phone_ignored(self):
        row = employee()
        row['phone_extra'] = '+998900000001'
        self.db.import_employees([row])
        self.assertIsNone(self.db.employee_list()[0]['phone_extra'])
        self.assertEqual(self.db.employee_list()[0]['approved'], 1)
        row['phone_extra'] = '+998900000003'
        self.db.import_employees([row])
        self.assertEqual(self.db.employee_list()[0]['approved'], 0)
        request_id = self.db.join_by_phone('90 000 00 03', 123456789)
        self.assertIsNotNone(request_id)
        self.db.decide_join(request_id, True, 987)
        self.assertEqual(self.db.employee_list()[0]['approved'], 1)
        with self.assertRaises(ValueError):
            self.db.join_by_phone('+998900000099', 999)

    async def test_extra_phone_export_and_blank_template(self):
        row = employee()
        row['phone_extra'] = '+998900000003'
        self.db.import_employees([row])
        settings = replace(self.settings, admin_ids=frozenset({987}))
        message = SimpleNamespace(reply_text=AsyncMock(), reply_document=AsyncMock())
        update = SimpleNamespace(effective_chat=SimpleNamespace(type='private'),
                                 effective_user=SimpleNamespace(id=987), effective_message=message)
        context = SimpleNamespace(bot_data={'settings': settings, 'db': self.db})
        await export_employees(update, context)
        data = message.reply_document.call_args.kwargs['document'].getvalue()
        imported = parse_excel(data, 'employees', TZ)
        self.assertEqual(imported[0]['phone_extra'], row['phone_extra'])
        await templates(update, context)
        template = load_workbook(message.reply_document.call_args.kwargs['document'])
        self.assertIn('qoshimcha_telefon', [cell.value for cell in template.active[1]])
        self.assertEqual(template.active.max_row, 1)
        template.close()

    async def test_automatic_daily_time_and_persistence(self):
        before = datetime(2027, 1, 1, 8, 30, tzinfo=TZ).timestamp()
        expected = datetime(2027, 1, 1, 9, 0, tzinfo=TZ).timestamp()
        self.db.import_appeals([appeal("AUTO"), appeal("EXPLICIT", before+20)])
        self.assertEqual(self.db.set_default_schedule("09:00", TZ, before), 1)
        self.assertEqual(self.db.details("AUTO")[0]["due_at"], expected)
        self.assertEqual(self.db.details("EXPLICIT")[0]["due_at"], before+20)
        restarted = Database(self.settings.database)
        self.assertEqual(restarted.default_schedule(), "09:00")
        restarted.import_appeals([appeal("LATER")])
        restarted.apply_default_schedule(TZ, expected+60)
        self.assertEqual(restarted.details("LATER")[0]["due_at"], expected+86400)
        self.assertEqual(restarted.details("AUTO")[0]["due_at"], expected)

    async def test_automatic_immediate_send_without_per_appeal_id(self):
        self.db.set_default_schedule("darhol", TZ)
        item = appeal("AUTO")
        self.db.import_appeals([item])
        await self.service.tick()
        self.assertEqual(self.channel(appeal_id="AUTO")["status"], "sent")
        self.assertEqual(self.bot.send_message.await_count, 1)
        self.db.import_appeals([item])
        await self.service.tick()
        self.assertEqual(self.bot.send_message.await_count, 1)

    async def test_auto_schedule_waits_for_approval_then_delivers(self):
        self.db.import_employees([employee("Other", "222", "+998900000003")])
        self.db.set_default_schedule("darhol", TZ)
        self.db.import_appeals([appeal("AUTO", ref="Other")])
        await self.service.tick()
        self.bot.send_message.assert_not_awaited()
        request_id = self.db.join_by_phone("+998900000003", 222)
        self.db.decide_join(request_id, True, 987)
        await self.service.tick()
        self.assertEqual(self.bot.send_message.call_args.kwargs["chat_id"], 222)

    async def test_changing_auto_schedule_preserves_explicit_and_sent(self):
        now = time.time()
        self.db.import_appeals([appeal("AUTO"), appeal("EXPLICIT", now+90000)])
        self.db.set_default_schedule("09:00", TZ, now)
        self.db.set_default_schedule("darhol", TZ, now)
        self.assertEqual(self.db.details("AUTO")[0]["due_at"], now)
        self.assertEqual(self.db.details("EXPLICIT")[0]["due_at"], now+90000)
        await self.service.tick()
        self.db.set_default_schedule("18:00", TZ, now+1)
        self.assertEqual(self.db.details("AUTO")[0]["due_at"], now)

    async def test_explicit_excel_change_stops_auto_rescheduling(self):
        now = time.time()
        self.db.import_appeals([appeal("AUTO")])
        self.db.set_default_schedule("09:00", TZ, now)
        self.db.import_appeals([appeal("AUTO", now+90000)])
        self.db.set_default_schedule("darhol", TZ, now)
        self.assertEqual(self.db.details("AUTO")[0]["due_at"], now+90000)

    async def test_schedule_wizard_and_send_now_do_not_request_id(self):
        settings = Settings("123:test", frozenset({987}), TZ, 9, self.settings.database, self.settings.jobs_database)
        self.db.import_appeals([appeal("AUTO")])
        message = SimpleNamespace(reply_text=AsyncMock(), text="Avtomatik yuborish")
        update = SimpleNamespace(effective_chat=SimpleNamespace(type="private"), effective_user=SimpleNamespace(id=987), effective_message=message)
        context = SimpleNamespace(bot_data={"settings": settings, "db": self.db}, user_data={}, args=[])
        await text_message(update, context)
        self.assertTrue(context.user_data["schedule"])
        message.text = "09:00"
        await text_message(update, context)
        self.assertEqual(self.db.default_schedule(), "09:00")
        message.text = "Hozir yuborish"
        await text_message(update, context)
        self.assertEqual(self.db.default_schedule(), "darhol")
        self.assertNotIn("input", context.user_data)
        await self.service.tick()
        self.bot.send_message.assert_awaited_once()

    async def test_non_admin_cannot_change_default_schedule(self):
        message = SimpleNamespace(reply_text=AsyncMock())
        update = SimpleNamespace(effective_chat=SimpleNamespace(type="private"), effective_user=SimpleNamespace(id=555), effective_message=message)
        context = SimpleNamespace(bot_data={"settings": self.settings, "db": self.db}, user_data={}, args=["darhol"])
        await automatic_schedule(update, context)
        self.assertIsNone(self.db.default_schedule())
        with self.assertRaises(ValueError):
            self.db.set_default_schedule("25:00", TZ)
        self.assertIsNone(self.db.default_schedule())


class ConfigurationSecurityTests(unittest.TestCase):
    def test_sms_password_command_uses_local_prompt_and_skips_environment_loading(self):
        password = "synthetic local password"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "bot.db"
            with patch.object(Settings, "database", path), patch.object(sys, "argv", ["main.py", "--set-sms-password"]), \
                    patch("main.sys.stdin.isatty", return_value=True), patch("main.getpass", side_effect=[password, password]), \
                    patch("main.load_settings") as load_settings_mock, patch("builtins.print") as output:
                run_main()
            load_settings_mock.assert_not_called()
            self.assertNotIn(password, repr(output.call_args_list))
            database = Database(path)
            self.assertTrue(database.sms_password_configured())
            with database.connect() as connection:
                stored = connection.execute("SELECT value FROM settings WHERE key='sms_password_hash'").fetchone()[0]
            self.assertNotIn(password, stored)
            self.assertEqual(database.change_sms_with_password(False, password, 987), "changed")

    def test_admin_ids_come_only_from_dotenv_without_interpolation(self):
        environment = {'BOT_TOKEN': '123:synthetic', 'ADMIN_IDS': '555', 'TIMEZONE': 'UTC'}
        with patch.dict(os.environ, environment, clear=True), patch('config.load_dotenv'), \
                patch('config.dotenv_values', return_value={'ADMIN_IDS': '987, 654'}) as read_config:
            settings = load_settings()
            self.assertEqual(settings.admin_ids, frozenset({987, 654}))
            self.assertEqual(settings.timezone.key, 'Asia/Tashkent')
            self.assertFalse(read_config.call_args.kwargs['interpolate'])
        with patch.dict(os.environ, environment, clear=True), patch('config.load_dotenv'), \
                patch('config.dotenv_values', return_value={}):
            self.assertEqual(load_settings().admin_ids, frozenset())
        with patch.dict(os.environ, environment, clear=True), patch('config.load_dotenv'), \
                patch('config.dotenv_values', return_value={'ADMIN_IDS': '${ADMIN_IDS}'}):
            with self.assertRaises(ValueError):
                load_settings()

    def test_imported_and_displayed_times_always_use_tashkent(self):
        from excel_utils import parse_datetime
        utc = ZoneInfo('UTC')
        settings = Settings('123:synthetic', frozenset(), utc)
        self.assertEqual(settings.timezone.key, 'Asia/Tashkent')
        moment = parse_datetime(datetime(2027, 1, 1, 9, tzinfo=utc), utc)
        self.assertEqual(moment.hour, 14)
        self.assertEqual(moment.tzinfo.key, 'Asia/Tashkent')
        naive = parse_datetime('2027-01-01 09:00', utc)
        self.assertEqual(naive.hour, 9)
        self.assertEqual(naive.utcoffset().total_seconds(), 18000)

    def test_log_formatter_redacts_secrets_in_messages_and_tracebacks(self):
        token = '123456789:ABCDEFGHIJKLMNOPQRSTUVWXYZabcde'
        password = 'synthetic-password-secret'
        message = ('https://api.telegram.org/bot' + token + '/sendMessage '
                   'Bearer synthetic-access-secret {"password": "hidden-password", '
                   '"accessToken": "eyJhbGciOiJIUzI1NiJ9.synthetic.signature"}')
        try:
            raise RuntimeError(password)
        except RuntimeError:
            record = logging.LogRecord('synthetic', logging.ERROR, '', 0, message, (), sys.exc_info())
        record.created = datetime(2027, 1, 1, 0, tzinfo=ZoneInfo('UTC')).timestamp()
        with patch.dict(os.environ, {'BOT_TOKEN': token, 'TEXTUP_PASSWORD': password}):
            rendered = RedactingFormatter('%(asctime)s %(message)s').format(record)
        for secret in (token, password, 'synthetic-access-secret', 'hidden-password', 'eyJhbGciOiJIUzI1NiJ9.synthetic.signature'):
            self.assertNotIn(secret, rendered)
        self.assertIn('2027-01-01 05:00:00+0500', rendered)
        self.assertNotIn(token, repr(Settings(token, frozenset(), TZ)))


class TextUPTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        self.requests = []
        self.responses = []
        self.provider = sms.TextUPClient('synthetic@example.invalid', 'synthetic-password',
                                        transport=httpx.MockTransport(self.handle))

    def handle(self, request):
        self.requests.append(request)
        response = self.responses.pop(0)
        if isinstance(response, Exception):
            raise response
        return response

    def login_response(self, token='synthetic-access-token'):
        return httpx.Response(200, json={'accessToken': token, 'refreshToken': 'unused-refresh',
                                       'user': {'id': '00000000-0000-4000-8000-000000000001'}})

    async def test_login_send_and_cached_token_use_documented_contract(self):
        self.responses = [self.login_response(), httpx.Response(200, json={'smsId': 'synthetic-sms-1'}),
                          httpx.Response(200, json={'smsId': 'synthetic-sms-2'})]
        text = sms.task_notification('synthetic_test_bot', 1)
        self.assertEqual(await self.provider.send('90 000 00 01', text), 'synthetic-sms-1')
        self.assertEqual(await self.provider.send('+998900000001', text), 'synthetic-sms-2')
        self.assertEqual(len(self.requests), 3)
        self.assertEqual(str(self.requests[0].url), 'https://api-auth.textup.uz/v1/login')
        self.assertEqual(json.loads(self.requests[0].content),
                         {'email': 'synthetic@example.invalid', 'password': 'synthetic-password'})
        request = self.requests[1]
        self.assertEqual(str(request.url), 'https://sms-api.textup.uz/v1/send')
        self.assertEqual(request.method, 'POST')
        self.assertEqual(request.headers['Authorization'], 'Bearer synthetic-access-token')
        self.assertEqual(request.headers['Content-Type'], 'application/json')
        self.assertEqual(json.loads(request.content), {'message': text,
            'userId': '00000000-0000-4000-8000-000000000001', 'name': 'Task notification',
            'recipients': ['+998900000001']})

    async def test_optional_template_and_nickname_without_group(self):
        self.provider = sms.TextUPClient('synthetic@example.invalid', 'synthetic-password',
            'synthetic-template', 'synthetic-nickname', transport=httpx.MockTransport(self.handle))
        self.responses = [self.login_response(), httpx.Response(200, json={'smsId': 'synthetic-sms'})]
        await self.provider.send('+998900000001', 'Synthetic notification')
        payload = json.loads(self.requests[1].content)
        self.assertEqual(payload['templateId'], 'synthetic-template')
        self.assertEqual(payload['nicknameId'], 'synthetic-nickname')
        self.assertNotIn('groupId', payload)

    async def test_expired_token_logs_in_again_once(self):
        self.responses = [self.login_response(), httpx.Response(401), self.login_response('new-token'),
                          httpx.Response(200, json={'smsId': 'synthetic-sms'})]
        self.assertEqual(await self.provider.send('+998900000001', 'Synthetic notification'), 'synthetic-sms')
        self.assertEqual(len(self.requests), 4)
        self.assertEqual(self.requests[-1].headers['Authorization'], 'Bearer new-token')
        self.assertEqual(json.loads(self.requests[1].content), json.loads(self.requests[-1].content))

    async def test_repeated_unauthorized_does_not_loop(self):
        self.responses = [self.login_response(), httpx.Response(401), self.login_response('new-token'),
                          httpx.Response(401)]
        with self.assertRaises(SMSNotConfigured):
            await self.provider.send('+998900000001', 'Synthetic notification')
        self.assertEqual(len(self.requests), 4)

    async def test_login_failures_are_safe_and_do_not_send(self):
        cases = [(httpx.Response(status, text='private-provider-response'), expected)
                 for status, expected in ((400, SMSNotConfigured), (401, SMSNotConfigured),
                     (403, SMSNotConfigured), (429, SMSRetryableError), (500, SMSRetryableError))]
        cases += [(httpx.Response(200, json={'accessToken': 'secret-token', 'user': {}}), SMSRetryableError),
                  (httpx.Response(200, text='invalid-json'), SMSRetryableError),
                  (httpx.ReadTimeout('private-transport-error'), SMSRetryableError)]
        for response, expected in cases:
            with self.subTest(response=type(response).__name__, expected=expected.__name__):
                self.requests = []
                self.responses = [response]
                with self.assertRaises(expected) as caught:
                    await self.provider.send('+998900000001', 'Synthetic notification')
                self.assertEqual(len(self.requests), 1)
                self.assertNotIn('private-', str(caught.exception))
                self.assertNotIn('secret-token', str(caught.exception))
                self.assertNotIn('synthetic-password', str(caught.exception))

    async def test_send_statuses_distinguish_rejection_retry_and_uncertainty(self):
        cases = [(httpx.Response(status, text='private-provider-response'), expected)
                 for status, expected in ((400, sms.SMSRejectedError), (403, sms.SMSRejectedError),
                     (422, sms.SMSRejectedError), (429, SMSRetryableError),
                     (408, sms.SMSDeliveryUnknown), (500, sms.SMSDeliveryUnknown),
                     (502, sms.SMSDeliveryUnknown), (302, sms.SMSDeliveryUnknown))]
        cases += [(httpx.Response(200, json=payload), sms.SMSDeliveryUnknown)
                  for payload in ({}, {'smsId': ''}, {'smsId': None}, {'smsId': 123}, [])]
        cases += [(httpx.Response(200, text='invalid-json'), sms.SMSDeliveryUnknown)]
        for response, expected in cases:
            with self.subTest(response=response.status_code, expected=expected.__name__):
                self.responses = [response]
                if not self.provider._access_token:
                    self.responses.insert(0, self.login_response())
                with self.assertRaises(expected) as caught:
                    await self.provider.send('+998900000001', 'Synthetic notification')
                self.assertNotIn('private-provider-response', str(caught.exception))
                self.assertEqual(self.responses, [])

    async def test_send_transport_failure_does_not_blindly_retry(self):
        cases = [(httpx.ConnectError, SMSRetryableError), (httpx.ConnectTimeout, SMSRetryableError),
                 (httpx.PoolTimeout, SMSRetryableError), (httpx.ReadTimeout, sms.SMSDeliveryUnknown),
                 (httpx.WriteTimeout, sms.SMSDeliveryUnknown), (httpx.WriteError, sms.SMSDeliveryUnknown),
                 (httpx.RemoteProtocolError, sms.SMSDeliveryUnknown)]
        for error, expected in cases:
            with self.subTest(error=error.__name__):
                self.responses = [error('private-transport-error')]
                if not self.provider._access_token:
                    self.responses.insert(0, self.login_response())
                with self.assertRaises(expected) as caught:
                    await self.provider.send('+998900000001', 'Synthetic notification')
                self.assertNotIn('private-transport-error', str(caught.exception))
                self.assertEqual(self.responses, [])

    async def test_invalid_local_input_never_contacts_provider(self):
        for phone, text in [('bad-phone', 'Synthetic'), ('+998900000001', '')]:
            with self.assertRaises(sms.SMSRejectedError):
                await self.provider.send(phone, text)
        self.assertEqual(self.requests, [])

    async def test_environment_settings_and_missing_credentials(self):
        with patch.dict(os.environ, {}, clear=True):
            with self.assertRaises(SMSNotConfigured):
                await sms.send_sms('+998900000001', 'Synthetic')
        configuration = ('synthetic@example.invalid', 'synthetic-password', 'template', 'nickname')
        provider = SimpleNamespace(configuration=configuration, send=AsyncMock(return_value='synthetic-sms'))
        environment = dict(zip(('TEXTUP_EMAIL', 'TEXTUP_PASSWORD', 'TEXTUP_TEMPLATE_ID', 'TEXTUP_NICKNAME_ID'), configuration))
        with patch.dict(os.environ, environment, clear=True), patch('sms._textup_client', None), \
                patch('sms.TextUPClient', return_value=provider) as factory:
            self.assertEqual(await sms.send_sms('+998900000001', 'Synthetic'), 'synthetic-sms')
            self.assertEqual(await sms.send_sms('+998900000001', 'Synthetic'), 'synthetic-sms')
            factory.assert_called_once_with(*configuration)
            self.assertEqual(provider.send.await_count, 2)


if __name__ == "__main__":
    unittest.main()
