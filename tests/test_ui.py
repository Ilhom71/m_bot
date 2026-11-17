"""Exercise real navigation with synthetic records and mocked Telegram transport."""
import ast
import inspect
import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock
import warnings
from zoneinfo import ZoneInfo

from config import Settings
from db import Database
from handlers import ui
from main import post_init
from test_bot import appeal, employee, workbook_bytes


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
            reply_document=AsyncMock(), text='', contact=None, document=None)
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

    async def test_public_commands_only_and_five_sections(self):
        await post_init(SimpleNamespace(bot_data=self.context.bot_data, bot=self.bot))
        for call in self.bot.set_my_commands.call_args_list:
            self.assertEqual([r.command for r in call.args[0]], ['start', 'menu', 'bekor'])
        await ui.start(self.update, self.context)
        await ui.home(self.update, self.context)
        self.assertEqual(sum(map(len, self.markup().keyboard)), 5)
        self.assertEqual(self.message.reply_text.await_count, 2)

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
            'joins:0', 'joins:1', 'request:2', 'schedule', 'status', 'extra', 'day:1', 'clock:1:today',
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
        self.assertIn('unknown', routes)
        self.assertNotIn('send', routes)

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
        self.assertIn(f"6. {visible[0]['name']}\n\n7. {visible[1]['name']}", text)
        self.assertNotIn('tasdig', text)
        self.assertNotIn('Smena:', text)
        self.assertNotIn(' | ', text)
        await self.click(f"worker:{visible[0]['id']}")
        text = self.last_screen['text']
        self.assertIn(visible[0]['phone'], text)
        self.assertIn('Ulanish:', text)
        self.assertIn('workers:1', str(self.markup()))
        from handlers.bot import employees
        await employees(self.update, self.context)
        self.assertIn('Sahifa 1/3', self.last_screen['text'])

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

