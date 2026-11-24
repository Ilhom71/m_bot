"""Exercise real navigation with synthetic records and mocked Telegram transport."""
import ast
import inspect
import tempfile
import time
import unittest
from datetime import datetime, timedelta
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, patch
import warnings
from zoneinfo import ZoneInfo

from config import Settings
from db import Database
from excel_utils import parse_excel
from handlers import ui
from main import post_init
from test_bot import appeal, employee, workbook_bytes


TEST_SMS_PASSWORD = "synthetic SMS control password"


class NavigationTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        path = Path(self.temp.name)
        self.settings = Settings('123:synthetic', frozenset({987}), ZoneInfo('Asia/Tashkent'), 9, path/'db', path/'jobs')
        self.db = Database(self.settings.database)
        for i in range(12):
            self.db.import_employees([employee(f'Worker {i}', str(1000+i), f'+99890000{i:04d}')])
        self.db.decide_join(self.db.request_join(1, 1000), True, 987)
        for i in range(16):
            row = appeal(f'LONG-{i}', time.time()-10, 'Worker 0')
            row['body'] = 'Synthetic content. ' * 300
            self.db.import_appeals([row])
        self.db.request_join(2, 1001)
        self.db.request_join(3, 1002)
        self.message = SimpleNamespace(message_id=42, reply_text=AsyncMock(return_value=SimpleNamespace(message_id=42)),
            reply_document=AsyncMock(), delete=AsyncMock(return_value=True), text='', contact=None, document=None)
        self.update = SimpleNamespace(effective_chat=SimpleNamespace(id=987, type='private'),
            effective_user=SimpleNamespace(id=987, username=None), effective_message=self.message)
        self.bot = SimpleNamespace(edit_message_text=AsyncMock(), edit_message_reply_markup=AsyncMock(), set_my_commands=AsyncMock())
        self.context = SimpleNamespace(user_data={}, bot=self.bot, bot_data={'settings': self.settings, 'db': self.db,
            'service': SimpleNamespace(notice=lambda *a: None, start=lambda: None)})

        self.last_screen = {}
        async def capture_reply(*args, **kwargs):
            self.last_screen = dict(kwargs, text=args[0] if args else kwargs.get('text'))
            return self.message.reply_text.return_value
        async def capture_edit(**kwargs):
            self.last_screen = kwargs
        self.message.reply_text.side_effect = capture_reply
        self.bot.edit_message_text.side_effect = capture_edit

    def tearDown(self):
        self.temp.cleanup()

    def markup(self):
        return self.last_screen['reply_markup']

    async def click(self, route, user=987):
        self.update.effective_user.id = user
        self.update.callback_query = SimpleNamespace(data=f"u:{ui.state(self.context)['session']}:{route}",
            message=self.message, answer=AsyncMock())
        await ui.callback(self.update, self.context)
        self.update.callback_query.answer.assert_awaited_once()

    async def enter_sms_password(self, password=TEST_SMS_PASSWORD):
        self.message.text = password
        await ui.input_message(self.update, self.context)
        self.message.text = ''

    async def test_public_commands_only_and_five_sections(self):
        await post_init(SimpleNamespace(bot_data=self.context.bot_data, bot=self.bot))
        for call in self.bot.set_my_commands.call_args_list:
            self.assertEqual([r.command for r in call.args[0]], ['start', 'menu', 'bekor'])
        await ui.start(self.update, self.context)
        await ui.home(self.update, self.context)
        self.assertEqual(sum(map(len, self.markup().keyboard)), 5)
        self.assertEqual(self.message.reply_text.await_count, 2)

    async def test_sms_switch_is_admin_only_and_visible_in_schedule(self):
        await ui.start(self.update, self.context)
        await self.click('schedule')
        self.assertIn('SMS yuborish: ✅ Yoqilgan', self.last_screen['text'])
        self.assertIn('sms:off', str(self.markup()))
        await self.click('sms:off')
        self.assertEqual(self.last_screen['text'], ui.SMS_PASSWORD_UNCONFIGURED)
        self.assertTrue(self.db.sms_enabled())
        self.db.set_sms_password(TEST_SMS_PASSWORD)
        await self.click('sms:off', user=555)
        self.assertTrue(self.db.sms_enabled())
        self.assertNotIn('flow', ui.state(self.context))
        await self.click('sms:off')
        self.assertTrue(self.db.sms_enabled())
        await self.enter_sms_password("wrong password")
        self.assertIn('Parol noto‘g‘ri', self.last_screen['text'])
        self.assertTrue(self.db.sms_enabled())
        await self.enter_sms_password()
        self.assertEqual(self.message.delete.await_count, 2)
        self.assertFalse(self.db.sms_enabled())
        self.assertIn('SMS yuborish: ❌ O‘chirilgan', self.last_screen['text'])
        self.assertIn('sms:on', str(self.markup()))
        self.assertIn('SMS yuborish: o‘chirilgan', self.db.status_text())
        await self.click('sms:off')
        await self.enter_sms_password()
        self.assertFalse(self.db.sms_enabled())
        await self.click('sms:on')
        self.assertFalse(self.db.sms_enabled())
        await self.enter_sms_password()
        self.assertTrue(self.db.sms_enabled())
        await self.click('sms:invalid')
        self.assertEqual(self.last_screen['text'], ui.STALE)
        self.assertTrue(self.db.sms_enabled())

    async def test_settings_keyboard_opens_sms_controls_and_keeps_legacy_shortcut(self):
        self.db.set_sms_password(TEST_SMS_PASSWORD)
        await ui.start(self.update, self.context)
        labels = [key.text for line in self.markup().keyboard for key in line]
        self.assertIn('⚙️ Sozlamalar', labels)
        self.assertNotIn('⚙️ Qo‘shimcha', labels)
        self.message.text = '⚙️ Sozlamalar'
        await ui.input_message(self.update, self.context)
        self.assertIn('⚙️ Sozlamalar', self.last_screen['text'])
        buttons = [b for line in self.markup().inline_keyboard for b in line]
        off = next(b for b in buttons if b.text == '❌ SMS o‘chirish')
        await self.click(off.callback_data.split(':', 2)[2], user=555)
        self.assertTrue(self.db.sms_enabled())
        await self.click(off.callback_data.split(':', 2)[2])
        await self.enter_sms_password()
        self.assertFalse(self.db.sms_enabled())
        self.assertIn('⚙️ Sozlamalar', self.last_screen['text'])
        self.assertIn('SMS yuborish: ❌ O‘chirilgan', self.last_screen['text'])
        self.assertIn('sms:on:extra', str(self.markup()))
        await self.click('sms:on:extra')
        await self.enter_sms_password()
        self.assertTrue(self.db.sms_enabled())
        await self.click('sms:off:status')
        self.assertTrue(self.db.sms_enabled())
        self.assertEqual(self.last_screen['text'], ui.STALE)
        self.message.text = '⚙️ Qo‘shimcha'
        await ui.input_message(self.update, self.context)
        self.assertIn('⚙️ Sozlamalar', self.last_screen['text'])
        self.assertIn('sms:off:extra', str(self.markup()))

    async def test_sms_password_requires_a_fresh_prompt_and_deleted_message(self):
        self.db.set_sms_password(TEST_SMS_PASSWORD)
        await ui.start(self.update, self.context)
        await self.click('sms:off')
        self.message.delete.return_value = False
        await self.enter_sms_password()
        self.assertTrue(self.db.sms_enabled())
        self.assertIn('Parol xabari o‘chmadi', self.last_screen['text'])
        await self.click('sms:off')
        ui.state(self.context)['flow']['expires'] = 0
        self.message.delete.return_value = True
        await self.enter_sms_password()
        self.assertTrue(self.db.sms_enabled())
        self.assertIn('muddati tugadi', self.last_screen['text'])

    async def test_missing_applicant_buttons_use_appeal_id_and_open_correct_card(self):
        for index, name in enumerate(("Ko'rsatilmagan", 'Ko‘rsatilmagan', '')):
            identifier = f'MISSING-NAME-{index}'
            row = appeal(identifier, time.time()-10, 'Worker 0')
            row['client_name'] = name
            self.db.import_appeals([row])
        await ui.start(self.update, self.context)
        await self.click('list:all:0')
        buttons = [b for line in self.markup().inline_keyboard for b in line]
        for index in range(3):
            identifier = f'MISSING-NAME-{index}'
            button = next(b for b in buttons if f'Murojaat #{identifier}' in b.text)
            await self.click(button.callback_data.split(':', 2)[2])
            self.assertIn('Synthetic test only', self.last_screen['text'])
            number = button.callback_data.rsplit(':', 1)[1]
            self.assertEqual(ui.appeal_id(self.context, number), identifier)
        self.assertFalse(any('Ko‘rsatilmagan' in b.text or "Ko'rsatilmagan" in b.text for b in buttons))

    async def test_saved_legacy_sender_appears_in_both_cards_and_list(self):
        row = appeal('LEGACY-SENDER', time.time()-10, 'Worker 0')
        row.update(client_name="Ko'rsatilmagan", legacy=True, source_data={'юборган ижрочи': 'Synthetic Sender'})
        self.db.import_appeals([row])
        await ui.start(self.update, self.context)
        await self.click('list:all:0')
        buttons = [b for line in self.markup().inline_keyboard for b in line]
        sender = next(b for b in buttons if 'Synthetic Sender' in b.text)
        route = sender.callback_data.split(':', 2)[2]
        await self.click(route)
        self.assertIn('Юборган ижрочи: Synthetic Sender', self.last_screen['text'])
        await self.click(route.replace('card:', 'own:'), user=1000)
        self.assertIn('Юборган ижрочи: Synthetic Sender', self.last_screen['text'])

    async def test_sent_telegram_is_not_attention_regardless_of_sms_switch(self):
        self.db.set_sms_password(TEST_SMS_PASSWORD)
        with self.db.connect() as conn:
            conn.execute("UPDATE deliveries SET status='sent' WHERE channel='telegram'")
            conn.execute("UPDATE deliveries SET status='failed',attempts=4 WHERE channel='sms'")
        await ui.start(self.update, self.context)
        await self.click('status')
        self.assertIn('E’tibor talab qiladigan murojaatlar: 0', self.last_screen['text'])
        self.assertNotIn('SMS', self.last_screen['text'])
        await self.click('sms:off')
        await self.enter_sms_password()
        await self.click('status')
        self.assertIn('E’tibor talab qiladigan murojaatlar: 0', self.last_screen['text'])
        self.assertIn('✅ Yuborishda e’tibor talab qiladigan muammo yo‘q.', self.last_screen['text'])
        self.assertNotIn('list:attention:0', str(self.markup()))
        await self.click('list:attention:0')
        self.assertIn('murojaatlar: 0', self.last_screen['text'])
        await self.click('list:all:0')
        labels = [b.text for line in self.markup().inline_keyboard for b in line if ':card:' in b.callback_data]
        self.assertTrue(labels)
        self.assertTrue(all(label.startswith('✅') for label in labels))
        await self.click('sms:on')
        await self.enter_sms_password()
        await self.click('status')
        self.assertIn('E’tibor talab qiladigan murojaatlar: 0', self.last_screen['text'])

    async def test_disabled_sms_preserves_real_telegram_attention_and_pagination(self):
        with self.db.connect() as conn:
            conn.execute("UPDATE deliveries SET status='sent' WHERE channel='telegram'")
            conn.execute("UPDATE deliveries SET status='unknown' WHERE channel='sms'")
            conn.execute("UPDATE deliveries SET status='failed' WHERE channel='telegram' AND appeal_id IN ('LONG-0','LONG-1','LONG-2','LONG-3','LONG-4','LONG-5')")
        self.db.set_sms_enabled(False)
        await ui.start(self.update, self.context)
        await self.click('status')
        self.assertIn('E’tibor talab qiladigan murojaatlar: 6', self.last_screen['text'])
        self.assertIn('Telegram: yuborishda xato: 6 ta', self.last_screen['text'])
        self.assertNotIn('SMS: yuborish natijasi noma’lum', self.last_screen['text'])
        await self.click('card:1')
        self.assertIn('⚠️ E’tibor sababi:\nTelegram: yuborishda xato', self.last_screen['text'])
        self.assertNotIn('SMS: yuborish natijasi noma’lum', self.last_screen['text'])
        await self.click('list:attention:0')
        self.assertIn('murojaatlar: 6', self.last_screen['text'])
        await self.click('pages:attention:0:list')
        self.assertIn('list:attention:1', str(self.markup()))
        await self.click('list:attention:1')
        self.assertEqual(sum(':card:' in b.callback_data for line in self.markup().inline_keyboard for b in line), 1)

    async def test_sent_card_hides_sms_and_attempts_without_changing_history(self):
        await ui.start(self.update, self.context)
        for status in ('pending', 'retry', 'blocked', 'failed', 'unknown', 'sent'):
            with self.db.connect() as conn:
                conn.execute("UPDATE deliveries SET status='sent',attempts=1 WHERE channel='telegram' AND appeal_id='LONG-0'")
                conn.execute("UPDATE deliveries SET status=?,attempts=4 WHERE channel='sms' AND appeal_id='LONG-0'", (status,))
            before = self.db.details('LONG-0')
            await self.click('card:1')
            self.assertIn('Holat: ✅ Yuborildi', self.last_screen['text'])
            for hidden in ('SMS', 'urinish', 'Urinishlar', 'Kutilmoqda', 'E’tibor sababi'):
                self.assertNotIn(hidden, self.last_screen['text'])
            self.assertNotIn('retry:', str(self.markup()))
            self.assertNotIn('unknown:', str(self.markup()))
            self.assertEqual(self.db.details('LONG-0'), before)

    async def test_sms_list_has_due_unsent_sms_and_independent_pagination(self):
        with self.db.connect() as conn:
            conn.execute("UPDATE deliveries SET status='sent' WHERE channel='telegram' AND appeal_id IN ('LONG-0','LONG-1')")
            conn.execute("UPDATE deliveries SET status='failed' WHERE channel='sms'")
            conn.execute("UPDATE deliveries SET status='sent' WHERE channel='sms' AND appeal_id='LONG-2'")
            conn.execute("UPDATE appeals SET due_at=? WHERE id='LONG-3'", (time.time()+3600,))
            conn.execute("UPDATE appeals SET due_at=NULL WHERE id='LONG-4'")
            conn.execute("UPDATE appeals SET completed_at=? WHERE id='LONG-5'", (time.time(),))
        await ui.start(self.update, self.context)
        self.assertEqual(len(ui.appeal_rows(self.context, 'sent')), 2)
        self.assertEqual(len(ui.appeal_rows(self.context, 'unsent')), 13)
        records = ui.appeal_rows(self.context, 'sms')
        self.assertEqual(len(records), 13)
        self.assertTrue({'LONG-0', 'LONG-1', 'LONG-5'}.issubset({row['id'] for row in records}))
        await self.click('smssettings')
        self.assertIn('SMSda bormaganlar: 13', self.last_screen['text'])
        self.db.set_sms_password(TEST_SMS_PASSWORD)
        await self.click('sms:off:smssettings')
        await self.enter_sms_password()
        self.assertFalse(self.db.sms_enabled())
        self.assertIn('SMS sozlamalari', self.last_screen['text'])
        await self.click('list:sms:0')
        self.assertIn('SMSda bormaganlar: 13', self.last_screen['text'])
        self.assertEqual(sum(':smscard:' in b.callback_data for line in self.markup().inline_keyboard for b in line), 5)
        await self.click('pages:sms:0:list')
        self.assertIn('list:sms:1', str(self.markup()))
        await self.click('list:sms:2')
        self.assertEqual(sum(':smscard:' in b.callback_data for line in self.markup().inline_keyboard for b in line), 3)
        await self.click('smscard:1', user=555)
        self.assertTrue(self.update.callback_query.answer.call_args.kwargs['show_alert'])

    async def test_sms_recovery_stays_in_sms_screen_and_preserves_telegram(self):
        with self.db.connect() as conn:
            conn.execute("UPDATE employees SET phone_extra='+998900000099' WHERE id=1")
        self.db.ready_deliveries(time.time())
        with self.db.connect() as conn:
            conn.execute("UPDATE deliveries SET status='sent' WHERE channel='telegram' AND appeal_id='LONG-0'")
            conn.execute("UPDATE sms_deliveries SET status='sent',attempts=1 WHERE appeal_id='LONG-0'")
            conn.execute("UPDATE sms_deliveries SET status='failed',attempts=4 WHERE appeal_id='LONG-0' AND phone='+998900000099'")
            self.db._aggregate_sms(conn, 'LONG-0')
        await ui.start(self.update, self.context)
        await self.click('smscard:1')
        self.assertIn('+998900000099', self.last_screen['text'])
        self.assertIn('urinish: 4', self.last_screen['text'])
        await self.click('retry:1:sms')
        await self.click('execute')
        self.assertIn('SMS — Murojaat #LONG-0', self.last_screen['text'])
        details = {row['phone']: row for row in self.db.sms_details('LONG-0')}
        self.assertEqual(details['+998900000099']['status'], 'pending')
        self.assertEqual(details['+998900000000']['status'], 'sent')
        with self.db.connect() as conn:
            conn.execute("UPDATE sms_deliveries SET status='unknown' WHERE appeal_id='LONG-0' AND phone='+998900000099'")
            self.db._aggregate_sms(conn, 'LONG-0')
        await self.click('unknown:1:sms')
        self.assertIn('smscard:1', str(self.markup()))
        await self.click('confirm:1:sms:yes')
        self.assertIn('smscard:1', str(self.markup()))
        await self.click('execute')
        self.assertIn('Holat: yuborildi', self.last_screen['text'])
        self.assertNotIn('LONG-0', {row['id'] for row in ui.appeal_rows(self.context, 'sms')})
        self.assertEqual(next(c for c in self.db.details('LONG-0')[1] if c['channel'] == 'telegram')['status'], 'sent')

    async def test_all_generated_buttons_have_dispatch_branches_and_fit_limits(self):
        tree = ast.parse(inspect.getsource(ui.dispatch))
        handled = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Compare) and isinstance(node.left, ast.Name) and node.left.id == 'route':
                for value in node.comparators:
                    handled.update(n.value for n in ast.walk(value) if isinstance(n, ast.Constant) and isinstance(n.value, str))
        self.assertEqual(ui.ROUTES, handled)
        await ui.start(self.update, self.context)
        routes = ['home', 'appeals', 'list:all:0', 'list:all:1', 'list:today:0', 'list:attention:0',
            'pages:all:0:list', 'card:1', 'moretext:1:1', 'staff', 'workers:0', 'workers:1',
            'pages:staff:1:workers', 'worker:1', 'assign:1:0', 'pages:staff:0:assign:1',
            'joins:0', 'joins:1', 'request:2', 'schedule', 'status', 'extra', 'smssettings', 'list:sms:0', 'smscard:1',
            'list:sent:0', 'list:unsent:0', 'day:1', 'clock:1:today',
            'clock:global', 'custom', 'date:1', 'add', 'upload:employees', 'upload:appeals',
            'retry:1:telegram', 'unknown:1:telegram', 'confirm:1:telegram:no', 'reject:2', 'attempts:1', 'appeal',
            'file:templates', 'file:employees', 'file:appeals']
        for route in routes:
            name, *args = route.split(':')
            await ui.dispatch(self.update, self.context, name, args)
            markup = self.markup()
            if not hasattr(markup, 'inline_keyboard'):
                for line in markup.keyboard:
                    for key in line:
                        self.assertIn(key.text, ui.ADMIN_KEYS | ui.USER_KEYS)
                continue
            self.assertLessEqual(sum(map(len, markup.inline_keyboard)), 8, route)
            for line in markup.inline_keyboard:
                for button in line:
                    self.assertLessEqual(len(button.callback_data.encode()), 64)
                    self.assertIn(button.callback_data.split(':')[2], handled)
        # State-dependent controls also have real handlers.
        with self.db.connect() as conn:
            conn.execute("UPDATE deliveries SET status='failed',attempts=4 WHERE appeal_id='LONG-0' AND channel='telegram'")
            conn.execute("UPDATE deliveries SET status='unknown',attempts=1 WHERE appeal_id='LONG-0' AND channel='sms'")
        await ui.card(self.update, self.context, '1')
        routes = [b.callback_data.split(':')[2] for line in self.markup().inline_keyboard for b in line]
        self.assertIn('retry', routes)
        self.assertNotIn('unknown', routes)
        self.assertNotIn('send', routes)
        await ui.sms_card(self.update, self.context, '1')
        self.assertIn('unknown', [b.callback_data.split(':')[2] for line in self.markup().inline_keyboard for b in line])

    async def test_every_admin_route_denied_before_database_access(self):
        await ui.start(self.update, self.context)
        old = self.context.bot_data['db']
        self.context.bot_data['db'] = object()
        for route in ui.ROUTES - ui.PUBLIC_ROUTES:
            await self.click(route+':1', user=555)
            self.assertTrue(self.update.callback_query.answer.call_args.kwargs['show_alert'])
        self.context.bot_data['db'] = old
        self.bot.edit_message_text.assert_not_awaited()

    async def test_restart_stale_message_and_malformed_callback_answered(self):
        await ui.start(self.update, self.context)
        self.context.user_data.clear()
        query = SimpleNamespace(data='u:old:card:1', message=self.message, answer=AsyncMock())
        self.update.callback_query = query
        await ui.callback(self.update, self.context)
        self.assertEqual(query.answer.call_args.args[0], ui.STALE)
        await ui.start(self.update, self.context)
        await self.click('card:missing')
        self.assertEqual(self.last_screen['text'], ui.STALE)

    async def test_employee_wizard_validation_keeps_current_step(self):
        await ui.start(self.update, self.context)
        await self.click('add')
        for text in ['New Worker', 'bad']:
            self.message.text = text
            await ui.input_message(self.update, self.context)
        self.assertEqual(ui.state(self.context)['flow']['values'], ['New Worker'])
        for text in ['91 000 00 99', '-', '-', "yo'q"]:
            self.message.text = text
            await ui.input_message(self.update, self.context)
        self.assertTrue(any(r['name'] == 'New Worker' for r in self.db.employee_list()))
        self.assertNotIn('flow', ui.state(self.context))

    async def test_time_assignment_and_pagination_without_typed_ids(self):
        await ui.start(self.update, self.context)
        await self.click('list:all:0')
        self.assertEqual(sum(b.callback_data.split(':')[2] == 'card' for line in self.markup().inline_keyboard for b in line), 5)
        await self.click('pick:1:2')
        self.assertEqual(self.db.details('LONG-0')[0]['employee_id'], 2)
        await self.click('clock:1:tomorrow')
        await self.click('hour:10:00')
        when = self.db.details('LONG-0')[0]['due_at']
        self.assertEqual(__import__('datetime').datetime.fromtimestamp(when, self.settings.timezone).strftime('%H:%M'), '10:00')

    async def test_retry_requires_confirmation_and_does_not_replay(self):
        with self.db.connect() as conn:
            conn.execute("UPDATE deliveries SET status='failed',attempts=4 WHERE appeal_id='LONG-0' AND channel='telegram'")
        await ui.start(self.update, self.context)
        await self.click('retry:1:telegram')
        self.assertEqual(next(r for r in self.db.details('LONG-0')[1] if r['channel'] == 'telegram')['attempts'], 4)
        await self.click('execute')
        telegram = next(r for r in self.db.details('LONG-0')[1] if r['channel'] == 'telegram')
        self.assertEqual(telegram['status'], 'pending')
        await self.click('execute')
        self.assertEqual(self.last_screen['text'], ui.STALE)

    async def test_reject_confirmation_and_stale_pending_action(self):
        await ui.start(self.update, self.context)
        await self.click('reject:2')
        self.assertEqual(len(self.db.pending_joins()), 2)
        await self.click('execute')
        self.assertEqual(len(self.db.pending_joins()), 1)
        with self.db.connect() as conn:
            conn.execute("UPDATE deliveries SET status='unknown',attempts=1 WHERE appeal_id='LONG-0' AND channel='telegram'")
        await self.click('confirm:1:telegram:no')
        with self.db.connect() as conn:
            conn.execute("UPDATE deliveries SET updated_at=updated_at+1 WHERE appeal_id='LONG-0'")
        await self.click('execute')
        self.assertEqual(self.last_screen['text'], ui.STALE)

    async def test_upload_offers_schedule_in_same_message(self):
        await ui.start(self.update, self.context)
        await self.click('upload:appeals')
        payload = workbook_bytes(['id','mijoz_ismi','telefon','murojaat_matni','yuborish_vaqti'],
            [['NEW', 'Synthetic', '+998900000002', 'Synthetic text', None]])
        self.message.document = SimpleNamespace(file_name='test.xlsx', file_size=len(payload), get_file=AsyncMock(
            return_value=SimpleNamespace(download_as_bytearray=AsyncMock(return_value=bytearray(payload)))))
        await ui.input_message(self.update, self.context)
        labels = [b.text for line in self.markup().inline_keyboard for b in line]
        self.assertIn('Darhol', labels)
        self.assertIn('09:00', labels)
        self.assertEqual(self.message.reply_text.await_count, 3)
        await self.click('mode:09:00')
        self.assertEqual(self.db.default_schedule(), '09:00')
        await self.click('off')
        self.assertIsNone(self.db.default_schedule())

    async def test_employee_contact_and_owned_text_only(self):
        await ui.start(self.update, self.context)
        self.update.effective_user.id = 1001
        await self.click('join', user=1001)
        self.message.contact = SimpleNamespace(user_id=1001, phone_number='+998900000001')
        await ui.input_message(self.update, self.context)
        self.assertNotIn('flow', ui.state(self.context))
        await self.click('own:1', user=1001)
        self.assertEqual(self.last_screen['text'], ui.STALE)

    async def test_sms_link_opens_owned_appeal_and_its_full_continuation(self):
        self.update.effective_user.id = 1000
        self.context.args = ['task_1']
        body = 'Synthetic content. ' + '😀' * 3000 + ' Synthetic tail.'
        with self.db.connect() as conn:
            conn.execute("UPDATE appeals SET body=? WHERE id='LONG-0'", (body,))
        await ui.start(self.update, self.context)
        self.assertIn('Murojaat #LONG-0', self.last_screen['text'])
        body = self.db.details('LONG-0')[0]['body']
        size = ui.BODY_PAGE_SIZE
        self.assertTrue(self.last_screen['text'].endswith(body[:size]))
        for page in range(1, (len(body)+size-1)//size):
            await self.click(f'moretext:1:{page}', user=1000)
            self.assertEqual(self.last_screen['text'], body[page*size:(page+1)*size])
        await self.click('mine:0', user=1000)
        self.assertIn('Murojaatingizni tanlang', self.last_screen['text'])

    async def test_sms_link_denies_other_unapproved_future_and_missing_appeals(self):
        self.context.args = ['task_1']
        for user in (1001, 555, 987):
            self.update.effective_user.id = user
            await ui.start(self.update, self.context)
            self.assertIn('sizga biriktirilmagan', self.last_screen['text'])
            self.assertNotIn('Synthetic content.', self.last_screen['text'])
        self.update.effective_user.id = 1000
        with self.db.connect() as conn:
            conn.execute('UPDATE employees SET approved=0 WHERE id=1')
        await ui.start(self.update, self.context)
        self.assertIn('sizga biriktirilmagan', self.last_screen['text'])
        with self.db.connect() as conn:
            conn.execute('UPDATE employees SET approved=1 WHERE id=1')
            conn.execute("UPDATE appeals SET due_at=? WHERE id='LONG-0'", (time.time()+3600,))
        await ui.start(self.update, self.context)
        self.assertIn('sizga biriktirilmagan', self.last_screen['text'])
        self.context.args = ['task_9999']
        await ui.start(self.update, self.context)
        self.assertIn('sizga biriktirilmagan', self.last_screen['text'])

    async def test_sms_link_rejects_malformed_payloads(self):
        self.update.effective_user.id = 1000
        for args in (['task_0'], ['task_-1'], ['task_invalid'], ['task_1', 'extra'], ['task_'+'9'*20]):
            self.context.args = args
            await ui.start(self.update, self.context)
            self.assertIn('havolasi noto‘g‘ri', self.last_screen['text'])

    async def test_optional_steps_cancel_and_unknown_resolution(self):
        await ui.start(self.update, self.context)
        await self.click('add')
        for value in ['Added worker', '93 123 45 67']:
            self.message.text = value
            await ui.input_message(self.update, self.context)
        for _ in range(3):
            await self.click('skip')
        self.assertTrue(any(r['name'] == 'Added worker' for r in self.db.employee_list()))
        await self.click('date:1')
        self.message.text = 'invalid'
        await ui.input_message(self.update, self.context)
        self.assertIn('flow', ui.state(self.context))
        await ui.cancel(self.update, self.context)
        self.assertNotIn('flow', ui.state(self.context))
        with self.db.connect() as conn:
            conn.execute("UPDATE deliveries SET status='unknown',attempts=1 WHERE appeal_id='LONG-0' AND channel='telegram'")
        await self.click('confirm:1:telegram:yes')
        await self.click('execute')
        self.assertEqual(next(r for r in self.db.details('LONG-0')[1] if r['channel'] == 'telegram')['status'], 'sent')

    async def test_conversation_registered_for_all_input_types(self):
        from telegram.ext import Application, ConversationHandler, CallbackQueryHandler
        app = Application.builder().token('123:synthetic').build()
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter('always')
            ui.register_handlers(app)
        self.assertFalse(any('per_message' in str(w.message) for w in caught))
        conversation = app.handlers[0][0]
        self.assertIsInstance(conversation, ConversationHandler)
        self.assertIn(ui.WAIT, conversation.states)
        self.assertTrue(conversation.allow_reentry)
        self.assertIsInstance(app.handlers[0][1], CallbackQueryHandler)

    async def test_reply_keyboard_home_and_inline_submenu_are_not_mixed(self):
        from telegram import InlineKeyboardMarkup, ReplyKeyboardMarkup
        from telegram.error import BadRequest
        self.bot.edit_message_reply_markup.side_effect = BadRequest("Message can't be edited")
        await ui.start(self.update, self.context)
        self.assertIsInstance(self.markup(), ReplyKeyboardMarkup)
        self.assertTrue(self.markup().is_persistent)
        self.assertNotIn('message', ui.state(self.context))
        await ui.dispatch(self.update, self.context, 'staff', [])
        self.assertIsInstance(self.markup(), InlineKeyboardMarkup)
        self.bot.edit_message_text.assert_not_awaited()
        self.bot.edit_message_text.side_effect = BadRequest('Message to edit not found')
        self.message.reply_text.return_value = SimpleNamespace(message_id=43)
        await ui.dispatch(self.update, self.context, 'staff', [])
        self.assertEqual(ui.state(self.context)['message'], 43)
        self.bot.edit_message_reply_markup.assert_not_awaited()

    async def test_employee_list_is_spaced_and_details_stay_in_card(self):
        await ui.start(self.update, self.context)
        await self.click('workers:1')
        text = self.last_screen['text']
        self.assertIn('jami 12 ta', text)
        self.assertIn('Sahifa 2/3', text)
        visible = self.db.employee_list()[5:10]
        self.assertIn(f"6. {visible[0]['name']} — ulanmagan\n\n7. {visible[1]['name']} — ulanmagan", text)
        self.assertIn(f"6. {visible[0]['name']} — ulanmagan",
                      [b.text for line in self.markup().inline_keyboard for b in line])
        self.assertNotIn('tasdig', text)
        self.assertNotIn('Smena:', text)
        self.assertNotIn(' | ', text)
        await self.click(f"worker:{visible[0]['id']}")
        text = self.last_screen['text']
        self.assertIn(visible[0]['phone'], text)
        self.assertIn('Ulanish:', text)
        self.assertIn('❌ Ulanmagan', text)
        self.assertIn('workers:1', str(self.markup()))
        from handlers.bot import employees
        await employees(self.update, self.context)
        self.assertIn('Sahifa 1/3', self.last_screen['text'])
        self.assertIn('1. Worker 0\n\n', self.last_screen['text'])
        self.assertIn('2. Worker 1 — ulanmagan', self.last_screen['text'])
        request = next(r for r in self.db.pending_joins() if r['telegram_id'] == 1001)
        self.db.decide_join(request['id'], True, 987)
        await self.click('workers:0')
        self.assertIn('2. Worker 1\n\n', self.last_screen['text'])
        self.assertIn('2. Worker 1', [b.text for line in self.markup().inline_keyboard for b in line])
        await self.click('worker:2')
        self.assertIn('✅ Ulangan', self.last_screen['text'])

    async def test_bottom_keyboard_routes_and_questions_follow_typed_answers(self):
        await ui.start(self.update, self.context)
        for label, route in ui.ADMIN_KEYS.items():
            self.message.text = label
            await ui.input_message(self.update, self.context)
            self.assertIn(route, ui.ROUTES)
            self.assertTrue(self.markup().inline_keyboard)
        await self.click('add')
        count = self.message.reply_text.await_count
        self.message.text = 'Synthetic worker'
        await ui.input_message(self.update, self.context)
        self.assertEqual(self.message.reply_text.await_count, count+1)
        self.assertIn('telefon', self.last_screen['text'].lower())
        self.message.text = '94 041 23 04'
        await ui.input_message(self.update, self.context)
        self.assertIn('Telegram ID', self.last_screen['text'])
        self.update.effective_user.id = 555
        for label in ui.ADMIN_KEYS:
            self.message.text = label
            await ui.input_message(self.update, self.context)
            self.assertIn('faqat adminlar', self.last_screen['text'])

    async def test_admin_downloads_employee_excel_from_staff_menu(self):
        await ui.start(self.update, self.context)
        await self.click('staff')
        export = next(button for row in self.markup().inline_keyboard for button in row
                      if button.text == 'Excelga chiqarish')
        action = export.callback_data.split(':', 2)[2]
        await self.click(action, user=555)
        self.message.reply_document.assert_not_awaited()
        await self.click(action)
        self.message.reply_document.assert_awaited_once()
        download = self.message.reply_document.await_args.kwargs
        self.assertEqual(download['filename'], 'xodimlar.xlsx')
        employees = parse_excel(download['document'].getvalue(), 'employees', self.settings.timezone)
        self.assertEqual(len(employees), 12)
        self.assertEqual(employees[0]['name'], 'Worker 0')

    async def test_sent_today_lists_telegram_success_by_tashkent_day(self):
        start = datetime(2026, 11, 24, tzinfo=self.settings.timezone)
        noon = start + timedelta(hours=12)
        yesterday = (start - timedelta(seconds=1)).timestamp()
        tomorrow = (start + timedelta(days=1)).timestamp()
        for identifier, due in (('FIRST-SENT', yesterday), ('LATEST-SENT', yesterday),
                                ('YESTERDAY-SENT', noon.timestamp()), ('SMS-ONLY', noon.timestamp()),
                                ('NEXT-DAY', noon.timestamp())):
            self.db.import_appeals([appeal(identifier, due, 'Worker 0')])
        with self.db.connect() as connection:
            connection.executemany("UPDATE deliveries SET status='sent',updated_at=? WHERE appeal_id=? AND channel=?",
                                   [(start.timestamp(), 'FIRST-SENT', 'telegram'),
                                    (noon.timestamp(), 'LATEST-SENT', 'telegram'),
                                    (yesterday, 'YESTERDAY-SENT', 'telegram'),
                                    (noon.timestamp(), 'SMS-ONLY', 'sms'),
                                    (tomorrow, 'NEXT-DAY', 'telegram')])

        with patch("db.datetime") as clock:
            clock.now.return_value = noon
            await ui.start(self.update, self.context)
            await self.click('appeals')
            self.assertTrue(any(b.text == 'Bugun yuborilgan murojaatlar' for line in self.markup().inline_keyboard for b in line))
            await self.click('list:sent_today:0')
            self.assertIn('Bugun yuborilgan murojaatlar: 2', self.last_screen['text'])
            ids = [row['id'] for row in ui.appeal_rows(self.context, 'sent_today')]
            self.assertEqual(ids, ['LATEST-SENT', 'FIRST-SENT'])
            self.assertIn('YESTERDAY-SENT', {row['id'] for row in ui.appeal_rows(self.context, 'today')})

    async def test_status_has_separate_assignment_problem_lists(self):
        row = appeal('UNCONNECTED', time.time()-10, 'Worker 1')
        missing = appeal('MISSING', time.time()-10, 'Missing Worker')
        missing['legacy'] = True
        self.db.import_appeals([row, missing])
        self.context.args = ['assignment_status']
        await ui.start(self.update, self.context)
        actions = [b.callback_data.split(':', 2)[2] for line in self.markup().inline_keyboard for b in line]
        self.assertIn('list:unconnected:0', actions)
        self.assertIn('list:unassigned:0', actions)
        await self.click('list:unconnected:0')
        self.assertIn('Xodimi ulanmagan murojaatlar: 1', self.last_screen['text'])
        self.assertIn('Botga ulanmagan xodimlar: 1 ta', self.last_screen['text'])
        card_action = next(b.callback_data.split(':', 2)[2] for line in self.markup().inline_keyboard
                           for b in line if ':card:' in b.callback_data)
        await self.click(card_action)
        self.assertIn('Murojaat #UNCONNECTED', self.last_screen['text'])
        await self.click('list:unassigned:0')
        self.assertIn('Mos xodim topilmagan murojaatlar: 1', self.last_screen['text'])
        self.update.effective_user.id = 555
        await ui.start(self.update, self.context)
        self.assertIn('faqat adminlar uchun', self.last_screen['text'])

    async def test_legacy_deadline_upload_warns_and_does_not_schedule_from_deadline(self):
        await ui.start(self.update, self.context)
        await self.click('upload:appeals')
        payload = workbook_bytes(['Кириш рақами', 'Аризачи', 'Ҳужжат мазмуни', 'Кимга юборилган', 'Муддати'],
                                 [['SYNTHETIC/1', 'Synthetic Client', 'Synthetic body', 'Worker 0', '2027-01-01']])
        self.message.document = SimpleNamespace(file_name='synthetic.xlsx', file_size=len(payload), get_file=AsyncMock(
            return_value=SimpleNamespace(download_as_bytearray=AsyncMock(return_value=bytearray(payload)))))
        await ui.input_message(self.update, self.context)
        self.assertIn('Муддати', self.last_screen['text'])
        self.assertIn('yuborish vaqti sifatida olinmadi', self.last_screen['text'])
        row = next(r for r in self.db.export_appeals() if r['legacy'])
        self.assertIsNone(row['due_at'])

    async def test_phone_export_upload_assigns_employee_from_telefon_column(self):
        await ui.start(self.update, self.context)
        await self.click('upload:appeals')
        headers = ['id', 'mijoz_ismi', 'telefon', 'murojaat_matni', 'yuborish_vaqti', 'xodim']
        payload = workbook_bytes(headers, [
            ['L-TEST/2026-1', 'Test Client', 900000000, 'Test body', None, 'WORKER 0'],
        ])
        self.message.document = SimpleNamespace(
            file_name='murojaatlar_bot_murojatlar asil_20261124_120154.xlsx', file_size=len(payload),
            get_file=AsyncMock(return_value=SimpleNamespace(
                download_as_bytearray=AsyncMock(return_value=bytearray(payload)))))
        await ui.input_message(self.update, self.context)

        self.assertIn('1 ta murojaat saqlandi', self.last_screen['text'])
        row, _ = self.db.details('L-TEST/2026-1')
        self.assertEqual(row['employee_id'], 1)
        self.assertEqual(row['phone'], "Ko'rsatilmagan")

    async def test_pp100_upload_saves_unique_appeals_and_reports_duplicates(self):
        await ui.start(self.update, self.context)
        await self.click('upload:appeals')
        payload = workbook_bytes(['Ариза рақами', 'Фамиля', 'Исми', 'Мурожат таснифи'],
                                 [['PP/1', 'Test surname', 'Test name', 'First body'],
                                  ['PP/1', 'Test surname', 'Test name', 'Duplicate body'],
                                  ['PP/2', 'Test surname', 'Test name', 'Unique body']])
        self.message.document = SimpleNamespace(file_name='pp100.xlsx', file_size=len(payload), get_file=AsyncMock(
            return_value=SimpleNamespace(download_as_bytearray=AsyncMock(return_value=bytearray(payload)))))
        await ui.input_message(self.update, self.context)

        self.assertIn('2 ta murojaat saqlandi', self.last_screen['text'])
        self.assertIn('1 ta takroriy qator', self.last_screen['text'])
        appeals = [row for row in self.db.export_appeals() if row['id'].startswith('PP100-')]
        self.assertEqual(len(appeals), 2)
        self.assertTrue(any(row['body'].endswith('First body') for row in appeals))

    async def test_disconnect_requires_admin_confirmation_and_approval_notifies_worker(self):
        await ui.start(self.update, self.context)
        await self.click('disconnect:1', user=555)
        self.assertTrue(self.db.employee_list()[0]['approved'])
        await self.click('worker:1')
        await self.click('disconnect:1')
        self.assertTrue(next(r for r in self.db.employee_list() if r['id'] == 1)['approved'])
        await self.click('execute')
        self.assertFalse(next(r for r in self.db.employee_list() if r['id'] == 1)['approved'])
        self.assertIn('❌ Ulanmagan', self.last_screen['text'])
        await self.click('approve:2')
        notices = self.db.ready_notices(time.time())
        self.assertTrue(any(r['admin_id'] == 1001 and 'tasdiqlandi' in r['text'] for r in notices))
        self.assertTrue(any(r['admin_id'] == 1000 and 'uzildi' in r['text'] for r in notices))

    async def test_done_button_is_owned_and_visible_to_admin(self):
        await ui.start(self.update, self.context)
        await self.click('own:1', user=1000)
        self.assertIn('✅ Bajarildi', [b.text for line in self.markup().inline_keyboard for b in line])
        await self.click('done:1', user=1001)
        self.assertIsNone(self.db.details('LONG-0')[0]['completed_at'])
        await self.click('own:1', user=1000)
        await self.click('done:1', user=1000)
        self.assertIn('Ijro: ✅ Bajarildi', self.last_screen['text'])
        self.assertNotIn('✅ Bajarildi', [b.text for line in self.markup().inline_keyboard for b in line])
        await self.click('card:1')
        self.assertIn('Ijro: ✅ Bajarildi', self.last_screen['text'])
        await self.click('list:completed:0')
        self.assertIn('Bajarilgan murojaatlar: 1', self.last_screen['text'])

