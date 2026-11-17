"""Integration tests use synthetic data and never send real messages."""
import asyncio
import json
import sqlite3
import subprocess
import sys
import tempfile
import time
import unittest
from dataclasses import replace
from zipfile import ZipFile
from xml.etree import ElementTree
from datetime import datetime
from io import BytesIO
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, patch
from zoneinfo import ZoneInfo

from openpyxl import Workbook, load_workbook
from telegram.error import BadRequest, Forbidden, RetryAfter, TimedOut

from config import Settings
from db import Database
from excel_utils import ValidationError, appeals_workbook, build_workbook, in_shift, parse_excel, phone_number
from handlers.bot import (ADMIN_MENU, USER_MENU, BUTTON_ACTIONS, automatic_schedule, send_now, command_actions, document, export_employees,
                          join_callback, join_contact, my_appeals, new_appeal, status, text_message, templates)
from main import InstanceLock
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

    def test_excel_cross_column_duplicate_has_row_number(self):
        for first, second in [(['A','901234567','911234567'], ['B','911234567',None]),
                              (['A','901234567','911234567'], ['B','931234567','911234567']),
                              (['A','901234567',None], ['B','931234567','901234567'])]:
            with self.assertRaisesRegex(ValidationError, '3-qator.*telefon'):
                parse_excel(workbook_bytes(['ism','telefon','qoshimcha_telefon'], [first, second]), 'employees', TZ)
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
        self.bot = SimpleNamespace(send_message=AsyncMock(return_value=SimpleNamespace(message_id=42)),
                                   send_document=AsyncMock(return_value=SimpleNamespace(message_id=43)))
        self.sms = AsyncMock(return_value="provider-test-id")
        self.service = DeliveryService(self.bot, self.db, self.settings, self.sms)

    def tearDown(self):
        self.temp.cleanup()

    def add_due(self, appeal_id="TEST-1", ref="Test Worker"):
        self.db.import_appeals([appeal(appeal_id, time.time()-10, ref)])

    def channel(self, channel="telegram", appeal_id="TEST-1"):
        return next(row for row in self.db.details(appeal_id)[1] if row["channel"] == channel)

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
        update = SimpleNamespace(effective_user=SimpleNamespace(id=222), effective_message=message)
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

    async def test_cross_employee_phone_duplicates_rollback(self):
        first = employee('First', '333', '+998900000003')
        first['phone_extra'] = '+998900000004'
        self.db.import_employees([first])
        for phone, extra in [('+998900000003', None), ('+998900000004', None),
                             ('+998900000005', '+998900000003'), ('+998900000005', '+998900000004')]:
            other = employee('Other', '444', phone, 7)
            other['phone_extra'] = extra
            with self.assertRaisesRegex(ValidationError, '7-qator.*telefon'):
                self.db.import_employees([employee('Rollback', '555', '+998900000099'), other])
            self.assertNotIn('Rollback', [r['name'] for r in self.db.employee_list()])

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


if __name__ == "__main__":
    unittest.main()
