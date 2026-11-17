"""Button navigation and conversations over the existing application services."""
import asyncio
import hashlib
import json
import re
import secrets
import time
import uuid
from datetime import datetime, timedelta

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, KeyboardButton, ReplyKeyboardMarkup
from telegram.error import BadRequest, TelegramError
from telegram.ext import CallbackQueryHandler, CommandHandler, ConversationHandler, MessageHandler, filters

from db import STATUSES
from excel_utils import MAX_FILE_SIZE, ValidationError, parse_datetime, parse_excel, phone_number, shift_value
from . import bot as legacy

WAIT = 1
STALE = "Bu tugma eskirgan, /menu bosing"
PUBLIC_ROUTES = {"home", "cancel", "join", "mine", "appeal", "moretext", "own"}
ROUTES = PUBLIC_ROUTES | {"appeals", "list", "pages", "card", "staff", "workers", "worker",
    "add", "upload", "joins", "request", "approve", "reject", "schedule", "mode", "off",
    "status", "extra", "file", "send", "assign", "pick", "day", "date", "clock", "hour",
    "custom", "retry", "unknown", "attempts", "confirm", "execute", "skip"}
ADMIN_KEYS = {'📋 Murojaatlar': 'appeals', '👥 Xodimlar': 'staff', '⏰ Yuborish': 'schedule',
              '📊 Holat': 'status', '⚙️ Qo‘shimcha': 'extra'}
USER_KEYS = {'📋 Murojaatlarim': 'mine', '📱 Ulanish': 'join', '📱 Xodim sifatida ulanish': 'join',
             '✉️ Murojaat yuborish': 'appeal'}


def state(context):
    return context.user_data.setdefault("ui", {"session": secrets.token_hex(3)})


def button(context, label, route):
    assert route.split(":")[0] in ROUTES, route
    data = f"u:{state(context)['session']}:{route}"
    assert len(data.encode()) <= 64
    return InlineKeyboardButton(label, callback_data=data)


def nav(context, back="home"):
    return [[button(context, "⬅️ Orqaga", back), button(context, "🏠 Menyu", "home")]]


async def show(update, context, text, rows=None, back="home"):
    keyboard = InlineKeyboardMarkup((rows or []) + (nav(context, back) if back is not None else []))
    assert sum(len(row) for row in keyboard.inline_keyboard) <= 8
    ui = state(context)
    text = text.encode('utf-16-le')[:8000].decode('utf-16-le', errors='ignore')
    message_id = ui.get("message")
    if message_id:
        try:
            await context.bot.edit_message_text(chat_id=update.effective_chat.id, message_id=message_id,
                                                text=text, reply_markup=keyboard)
            return
        except BadRequest as exc:
            if "message is not modified" in str(exc).lower():
                return
        except TelegramError:
            pass
    # Telegram only permits editing messages with inline keyboards or no markup.
    # Do not send ReplyKeyboardRemove and then try to attach an inline keyboard.
    message = await update.effective_message.reply_text(text, reply_markup=keyboard)
    ui["message"] = message.message_id


def admin(update, context):
    return update.effective_chat.type == "private" and update.effective_user.id in context.bot_data["settings"].admin_ids


def db(context):
    return context.bot_data["db"]


def clear_flow(context):
    state(context).pop("flow", None)
    state(context).pop("pending", None)


async def home(update, context):
    clear_flow(context)
    state(context).pop('message', None)
    state(context)['session'] = secrets.token_hex(3)
    if admin(update, context):
        rows = [['📋 Murojaatlar', '👥 Xodimlar'], ['⏰ Yuborish', '📊 Holat'], ['⚙️ Qo‘shimcha']]
    else:
        workers = db(context).employee_list()
        employee = any(r["telegram_id"] == update.effective_user.id for r in workers)
        rows = [['📋 Murojaatlarim', '📱 Ulanish']] if employee else [['✉️ Murojaat yuborish'], ['📱 Xodim sifatida ulanish']]
    await update.effective_message.reply_text('Pastdagi klaviaturadan kerakli bo‘limni tanlang.',
        reply_markup=ReplyKeyboardMarkup(rows, resize_keyboard=True, is_persistent=True,
                                        input_field_placeholder='Pastdagi tugmalardan tanlang'))


async def start(update, context):
    if update.effective_chat.type != "private":
        return ConversationHandler.END
    db(context).register_user(update.effective_user.id, update.effective_user.username)
    await home(update, context)
    return WAIT


async def cancel(update, context):
    await home(update, context)
    return WAIT


def appeal_rows(context, kind="all"):
    database = db(context)
    statuses = {r["id"]: r for r in database.list_appeals(context.bot_data["settings"].timezone, today=kind == "today")}
    with database.connect() as conn:
        rows = [dict(r) for r in conn.execute("SELECT rowid AS number,id,client_name FROM appeals ORDER BY created_at DESC,rowid DESC")]
    return [dict(r, **{k: v for k, v in statuses[r["id"]].items() if k != "id"}) for r in rows if r["id"] in statuses]


def attention(row):
    return (row["telegram"] in ("unknown", "failed") or row["sms"] in ("unknown", "failed")
            or not row["employee"] or not row["employee_approved"] or row["due_at"] is None)


def appeal_id(context, number):
    with db(context).connect() as conn:
        row = conn.execute("SELECT id FROM appeals WHERE rowid=?", (int(number),)).fetchone()
    if not row:
        raise ValueError(STALE)
    return row[0]


def snapshot(context, number):
    identifier = appeal_id(context, number)
    return hashlib.sha256(json.dumps(db(context).details(identifier), sort_keys=True).encode()).hexdigest()


async def listing(update, context, kind, page=0):
    rows = appeal_rows(context, kind)
    if kind == "attention":
        rows = [r for r in rows if attention(r)]
    total = max(1, (len(rows) + 4) // 5)
    page = max(0, min(int(page), total - 1))
    buttons = []
    for row in rows[page*5:page*5+5]:
        icon = "⚠️" if attention(row) else "✅" if row["telegram"] == row["sms"] == "sent" else "⏳"
        buttons.append([button(context, f"{icon} {row['client_name'][:40]}", f"card:{row['number']}")])
    if total > 1:
        buttons.append([button(context, f"◀️ {page+1}/{total} ▶️", f"pages:{kind}:{page}:list")])
    state(context)["list_back"] = f"list:{kind}:{page}"
    await show(update, context, f"{'Bugungi' if kind == 'today' else 'Barcha' if kind == 'all' else 'E’tibor talab qiladigan'} murojaatlar: {len(rows)}", buttons, "appeals")


async def card(update, context, number, attempts=False):
    identifier = appeal_id(context, number)
    row, channels = db(context).details(identifier)
    employees = {e["id"]: e for e in db(context).employee_list()}
    employee = employees.get(row["employee_id"])
    zone = context.bot_data["settings"].timezone
    when = datetime.fromtimestamp(row["due_at"], zone).strftime("%d.%m.%Y %H:%M") if row["due_at"] else "belgilanmagan"
    text = f"Mijoz: {row['client_name'][:150]}\nTelefon: {row['phone']}\nXodim: {employee['name'] if employee else row['employee_hint'] or 'biriktirilmagan'}\nVaqt: {when}\n\n"
    summary = "\n".join(f"{c['channel'].upper()}: {STATUSES[c['status']]} | urinish: {c['attempts']}" for c in channels)
    body = row['body']
    text += body[:2400] + ("\n… Davomi tugmasida." if len(body) > 2400 else "") + "\n\n" + summary
    buttons = []
    unlocked = all(c["attempts"] == 0 and c["status"] not in ("sent", "sending", "unknown") for c in channels)
    if unlocked:
        buttons += [[button(context, "📨 Hozir yuborish", f"send:{number}"), button(context, "📅 Vaqt belgilash", f"day:{number}")],
                    [button(context, "👤 Xodim biriktirish", f"assign:{number}:0")]]
    if attempts:
        text += "\n\n" + "\n".join(f"{c['channel']}: {c['last_error'] or 'xato yo‘q'}" for c in channels)
        text += "\n" + "\n".join(f"SMS {c['phone']}: {STATUSES[c['status']]} | urinish: {c['attempts']} | {c['last_error'] or 'xato yo‘q'}" for c in db(context).sms_details(identifier))
        buttons = []
    else:
        for channel in channels:
            name = channel['channel']
            if channel["status"] == "failed":
                buttons.append([button(context, f"🔄 Qayta urinish ({name.upper()})", f"retry:{number}:{name}")])
            if channel["status"] == "unknown":
                buttons.append([button(context, f"{name.upper()}: natijani belgilash", f"unknown:{number}:{name}")])
        if any(c['attempts'] or c['status'] == 'blocked' for c in channels):
            buttons.append([button(context, "Urinishlar", f"attempts:{number}")])
        if len(body) > 2400:
            buttons.append([button(context, "Matn davomi", f"moretext:{number}:1")])
    await show(update, context, text, buttons, state(context).get("list_back", "list:all:0"))


async def choose_workers(update, context, number=None, page=0):
    rows = db(context).employee_list()
    total = max(1, (len(rows) + 4) // 5)
    page = max(0, min(int(page), total - 1))
    visible = rows[page*5:page*5+5]
    buttons = [[button(context, f"{page*5+i}. {r['name'][:48]}",
                       f"pick:{number}:{r['id']}" if number is not None else f"worker:{r['id']}")]
               for i, r in enumerate(visible, 1)]
    if total > 1:
        route = f"assign:{number}" if number is not None else "workers"
        buttons.append([button(context, f"◀️ Sahifa {page+1}/{total} ▶️", f"pages:staff:{page}:{route}")])
    if number is None:
        state(context)['workers_back'] = f'workers:{page}'
    text = (f"👥 Xodimlar — jami {len(rows)} ta\nSahifa {page+1}/{total}\n\n" +
            "\n\n".join(f"{page*5+i}. {r['name']}" for i, r in enumerate(visible, 1)) +
            ("\n\nBiriktirish uchun xodimni tanlang." if number is not None else "\n\nTelefon va ulanish holatini ko‘rish uchun xodimni bosing.")) if rows else "Xodimlar ro‘yxati bo‘sh."
    await show(update, context, text, buttons, f"card:{number}" if number is not None else "staff")


async def hours(update, context, target, day="today"):
    state(context)["time_target"] = (target, day)
    buttons = [[button(context, value, f"hour:{value}") for value in ("09:00", "10:00", "14:00")],
               [button(context, value, f"hour:{value}") for value in ("15:00", "18:00")],
               [button(context, "Boshqa vaqt", "custom")]]
    await show(update, context, "Soatni tanlang (Asia/Tashkent).", buttons, "schedule" if target == "global" else f"day:{target}")


async def prompt(update, context, flow, question, back="home", extra=None):
    state(context)["flow"] = flow
    buttons = extra or []
    buttons.append([button(context, "❌ Bekor", "cancel")])
    await show(update, context, question, buttons, back)


async def confirm_action(update, context, action, number=None, channel=None, sent=None):
    state(context)["pending"] = dict(action=action, number=number, channel=channel, sent=sent,
        snapshot=snapshot(context, number) if action != "reject" else None)
    await show(update, context, "Ishonchingiz komilmi? Qabul qiluvchidan natijani tekshiring. Bu amal qayta yuborishga olib kelishi mumkin.",
               [[button(context, "✅ Ha", "execute"), button(context, "❌ Yo‘q", f"card:{number}" if action != "reject" else "joins:0")]])


async def callback(update, context):
    query = update.callback_query
    parts = (query.data or "").split(":")
    route = parts[2] if len(parts) >= 3 else ""
    # Authorization precedes all database access and mutation, including forged routes.
    if update.effective_chat.type != "private" or (route not in PUBLIC_ROUTES and not admin(update, context)):
        await query.answer("Bu amal faqat adminlar uchun.", show_alert=True)
        return WAIT
    ui = context.user_data.get("ui")
    if not ui or len(parts) < 3 or parts[0] != "u" or parts[1] != ui['session'] or route not in ROUTES or (ui.get('message') and query.message.message_id != ui['message']):
        await query.answer(STALE, show_alert=True)
        return WAIT
    await query.answer()
    if route != 'skip':
        ui.pop("flow", None)
    try:
        await dispatch(update, context, route, parts[3:])
    except (ValueError, IndexError, KeyError, TypeError, StopIteration):
        await show(update, context, STALE)
    return WAIT


async def dispatch(update, context, route, args):
    database, ui = db(context), state(context)
    settings = context.bot_data['settings']
    if route in ('home', 'cancel'):
        await home(update, context)
    elif route == 'appeals':
        await show(update, context, "Murojaatlar", [[button(context, label, action)] for label, action in
            (("Bugun", "list:today:0"), ("Hammasi", "list:all:0"), ("Excel yuklash", "upload:appeals"))])
    elif route == 'list':
        await listing(update, context, args[0], int(args[1]))
    elif route == 'pages':
        kind, page = args[0], int(args[1])
        destination = ':'.join(args[2:])
        records = database.employee_list() if kind == 'staff' else appeal_rows(context, kind)
        count = sum(attention(r) for r in records) if kind == 'attention' else len(records)
        total = max(1, (count+4)//5)
        choices = []
        for p, label in ((page-1, '◀️ Oldingi'), (page+1, 'Keyingi ▶️')):
            if 0 <= p < total:
                action = f"list:{kind}:{p}" if destination == 'list' else f"{destination}:{p}"
                choices.append([button(context, label, action)])
        await show(update, context, "Sahifani almashtiring.", choices)
    elif route in ('card', 'attempts'):
        await card(update, context, args[0], route == 'attempts')
    elif route == 'moretext':
        identifier = appeal_id(context, args[0])
        row = database.details(identifier)[0]
        if not admin(update, context) and not any(r['id'] == identifier for r in database.export_appeals(telegram_id=update.effective_user.id)):
            raise ValueError(STALE)
        page = int(args[1]); text = row['body'][page*2400:(page+1)*2400]
        buttons = [[button(context, 'Davomi ▶️', f"moretext:{args[0]}:{page+1}")]] if len(row['body']) > (page+1)*2400 else []
        await show(update, context, text or 'Matn tugadi.', buttons, f'card:{args[0]}' if admin(update, context) else 'mine:0')
    elif route == 'staff':
        await show(update, context, "Xodimlar", [[button(context, label, action)] for label, action in
            (("Ro‘yxat", "workers:0"), ("Qo‘shish", "add"), ("Excel yuklash", "upload:employees"),
             (f"Ulanish so‘rovlari ({len(database.pending_joins())})", "joins:0"))])
    elif route == 'workers':
        await choose_workers(update, context, page=int(args[0]))
    elif route == 'assign':
        await choose_workers(update, context, args[0], int(args[1]))
    elif route in ('worker', 'pick'):
        employee = next((r for r in database.employee_list() if r['id'] == int(args[-1])), None)
        if not employee:
            raise ValueError(STALE)
        if route == 'pick':
            database.assign(appeal_id(context, args[0]), employee['name'])
            await card(update, context, args[0])
        else:
            await show(update, context, f"👤 {employee['name']}\n\n📱 Asosiy telefon: {employee['phone']}\nQo‘shimcha telefon: {employee['phone_extra'] or 'yo‘q'}\n\nUlanish: {'✅ Tasdiqlangan' if employee['approved'] else '⏳ Ulanish yoki admin tasdig‘i kutilmoqda'}\nSmena: {employee['shift'] or 'cheklanmagan'}", back=ui.get('workers_back', 'workers:0'))
    elif route == 'add':
        await prompt(update, context, {'kind': 'employee', 'values': []}, EMPLOYEE_QUESTIONS[0], 'staff')
    elif route == 'upload':
        kind = args[0]
        if kind not in ('appeals', 'employees'):
            raise ValueError(STALE)
        await prompt(update, context, {'kind': 'upload', 'type': kind}, "Excel (.xlsx) faylini yuboring. Namuna: " + ('murojaatlar.xlsx' if kind == 'appeals' else 'xodimlar.xlsx'), 'appeals' if kind == 'appeals' else 'staff')
    elif route in ('joins', 'request'):
        rows = database.pending_joins()
        index = min(int(args[0]), max(0, len(rows)-1)) if route == 'joins' else next((i for i,r in enumerate(rows) if r['id'] == int(args[0])), -1)
        if index < 0:
            raise ValueError(STALE)
        if not rows:
            await show(update, context, 'Kutilayotgan ulanish so‘rovlari yo‘q.', back='staff')
            return
        row = rows[index]
        buttons = [[button(context, '✅ Tasdiqlash', f"approve:{row['id']}"), button(context, '❌ Rad etish', f"reject:{row['id']}")]]
        if index:
            buttons.append([button(context, '◀️ Oldingi', f'joins:{index-1}')])
        if index+1 < len(rows):
            buttons.append([button(context, 'Keyingi ▶️', f'joins:{index+1}')])
        await show(update, context, f"So‘rov {index+1}/{len(rows)}\n{row['name']}\nAsosiy: {row['phone']}\nQo‘shimcha: {row['phone_extra'] or 'yo‘q'}\nTelegram ID: {row['telegram_id']}", buttons, 'staff')
    elif route == 'approve':
        legacy.decide_request(context, int(args[0]), True, update.effective_user.id)
        await dispatch(update, context, 'joins', ['0'])
    elif route == 'reject':
        await confirm_action(update, context, 'reject', args[0])
    elif route == 'schedule':
        mode = database.default_schedule()
        text = 'Hozir avtomatik: yoqilgan' if mode == 'darhol' else f'Har kuni {mode}' if mode else 'Avtomatik yuborish: o‘chirilgan'
        await show(update, context, text + '\nExcelda alohida vaqt bo‘lsa, o‘sha vaqt ishlatiladi.',
            [[button(context, 'Darhol', 'mode:darhol')], [button(context, 'Soat tanlash', 'clock:global')], [button(context, 'O‘chirish', 'off')]])
    elif route == 'mode':
        database.set_default_schedule(':'.join(args), settings.timezone)
        database.assign_due(time.time(), settings.timezone)
        await dispatch(update, context, 'schedule', [])
    elif route == 'off':
        # Disable future default scheduling; explicit/already scheduled jobs retain their times.
        with database.connect() as conn:
            conn.execute("DELETE FROM settings WHERE key='default_send_time'")
        await show(update, context, 'Avtomatik jadval o‘chirildi. Oldin belgilangan yuborish vaqtlari saqlanadi.', back='schedule')
    elif route == 'status':
        count = sum(attention(r) for r in appeal_rows(context))
        await show(update, context, database.status_text().replace('/jadval ni tanlang', 'Yuborish bo‘limidan tanlang') + f'\n\nE’tibor talab qiladigan murojaatlar: {count}',
            [[button(context, f'⚠️ E’tibor kerak ({count})', 'list:attention:0')]])
    elif route == 'extra':
        await show(update, context, 'Shablonlar va Excel tahrirlash', [[button(context, label, action)] for label, action in
            (('Bo‘sh shablonlar', 'file:templates'), ('Xodimlarni Excel olish', 'file:employees'), ('Murojaatlarni Excel olish', 'file:appeals'))])
    elif route == 'file':
        await {'templates': legacy.templates, 'employees': legacy.export_employees, 'appeals': legacy.export_all}[args[0]](update, context)
    elif route == 'send':
        identifier = appeal_id(context, args[0]); row = database.details(identifier)[0]
        worker = next((r for r in database.employee_list() if r['id'] == row['employee_id'] and r['approved']), None)
        if not worker:
            await show(update, context, 'Avval xodim biriktiring va uning ulanishini tasdiqlang.', back=f'card:{args[0]}')
            return
        database.set_time(identifier, time.time())
        await card(update, context, args[0])
    elif route == 'day':
        await show(update, context, 'Qaysi kuni yuborilsin?', [[button(context, 'Bugun', f'clock:{args[0]}:today'), button(context, 'Ertaga', f'clock:{args[0]}:tomorrow')],
            [button(context, 'Sana yozish', f'date:{args[0]}')]], f'card:{args[0]}')
    elif route == 'date':
        await prompt(update, context, {'kind': 'date', 'target': args[0]}, 'Sanani yozing. Masalan: 2026-12-01', f'day:{args[0]}')
    elif route == 'clock':
        await hours(update, context, args[0], args[1] if len(args) > 1 else 'today')
    elif route == 'custom':
        await prompt(update, context, {'kind': 'clock'}, 'Soatni yozing. Masalan: 09:30')
    elif route == 'hour':
        await save_hour(update, context, ':'.join(args))
    elif route == 'retry':
        await confirm_action(update, context, 'retry', args[0], args[1])
    elif route == 'unknown':
        await show(update, context, f'{args[1].upper()} natijasini qabul qiluvchidan tekshiring.', [[button(context, 'Yuborildi', f'confirm:{args[0]}:{args[1]}:yes'),
            button(context, 'Yuborilmadi', f'confirm:{args[0]}:{args[1]}:no')]], f'card:{args[0]}')
    elif route == 'confirm':
        await confirm_action(update, context, 'unknown', args[0], args[1], args[2] == 'yes')
    elif route == 'execute':
        pending = ui.pop('pending')
        number = pending['number']
        if pending['action'] == 'reject':
            legacy.decide_request(context, int(number), False, update.effective_user.id)
            await dispatch(update, context, 'joins', ['0'])
        else:
            if snapshot(context, number) != pending['snapshot']:
                raise ValueError(STALE)
            identifier = appeal_id(context, number)
            if pending['action'] == 'retry':
                database.retry(identifier, pending['channel'])
            else:
                database.resolve_unknown(identifier, pending['channel'], pending['sent'])
            await card(update, context, number)
    elif route == 'appeal':
        await prompt(update, context, {'kind': 'appeal', 'values': []}, APPEAL_QUESTIONS[0])
    elif route == 'skip':
        flow = ui.get('flow')
        if not flow or flow['kind'] != 'employee' or len(flow['values']) < 2:
            raise ValueError(STALE)
        await input_message(update, context, '-')
    elif route == 'join':
        ui['flow'] = {'kind': 'join'}
        # Telegram only supports contact sharing on reply keyboards.
        await update.effective_message.reply_text('O‘z kontaktingizni yuboring.', reply_markup=ReplyKeyboardMarkup(
            [[KeyboardButton('📱 Telefonimni ulashish', request_contact=True)], ['❌ Bekor']], resize_keyboard=True, one_time_keyboard=True))
        await show(update, context, 'Telefon yuborish tugmasini bosing. Keyin admin tasdiqlaydi.', [[button(context, '❌ Bekor', 'cancel')]])
    elif route == 'mine':
        rows = database.export_appeals(telegram_id=update.effective_user.id)
        page = int(args[0]) if args else 0
        buttons = []
        for row in rows[page*4:page*4+4]:
            with database.connect() as conn:
                number = conn.execute('SELECT rowid FROM appeals WHERE id=?', (row['id'],)).fetchone()[0]
            buttons.append([button(context, row['client_name'][:40], f'own:{number}')])
        if page > 0:
            buttons.append([button(context, '◀️ Oldingi', f'mine:{page-1}')])
        if (page+1)*4 < len(rows):
            buttons.append([button(context, 'Keyingi ▶️', f'mine:{page+1}')])
        await show(update, context, 'Murojaatingizni tanlang.' if rows else 'Sizga biriktirilgan va vaqti kelgan murojaat yo‘q.', buttons)
    elif route == 'own':
        identifier = appeal_id(context, args[0])
        row = next((r for r in database.export_appeals(telegram_id=update.effective_user.id) if r['id'] == identifier), None)
        if not row:
            raise ValueError(STALE)
        buttons = [[button(context, 'Davomi ▶️', f'moretext:{args[0]}:1')]] if len(row['body']) > 2400 else []
        await show(update, context, f"Mijoz: {row['client_name'][:150]}\nTelefon: {row['phone']}\n\n{row['body'][:2400]}", buttons, 'mine:0')


EMPLOYEE_QUESTIONS = ['Xodim ismini yozing. Masalan: Aliyev Vali', 'Asosiy telefonni yozing. Masalan: 90 123 45 67',
    'Telegram ID yoki username yozing. Masalan: @vali. Bilmasangiz - yozing.',
    'Smenani yozing. Masalan: 09:00-18:00. Cheklov bo‘lmasa - yozing.',
    'Qo‘shimcha telefonni yozing. Masalan: 91 123 45 67. Bo‘lmasa yo‘q yoki - yozing.']
APPEAL_QUESTIONS = ['Ism-familiyangizni yozing. Masalan: Aliyev Vali', 'Telefoningizni yozing. Masalan: 90 123 45 67', 'Murojaatingizni yozing (1–1500 belgi). Masalan: Hujjatim bo‘yicha ma’lumot kerak.']


async def save_hour(update, context, value):
    if not re.fullmatch(r'(?:[01]\d|2[0-3]):[0-5]\d', value):
        raise ValueError('Soat noto‘g‘ri. Masalan: 09:30')
    target, day = state(context)['time_target']
    settings = context.bot_data['settings']
    if target == 'global':
        db(context).set_default_schedule(value, settings.timezone)
        await dispatch(update, context, 'schedule', [])
    else:
        date = datetime.now(settings.timezone).date()
        if day == 'tomorrow':
            date += timedelta(days=1)
        elif day != 'today':
            date = datetime.strptime(day, '%Y-%m-%d').date()
        moment = parse_datetime(f'{date} {value}', settings.timezone)
        db(context).set_time(appeal_id(context, target), moment.timestamp())
        await card(update, context, target)


async def input_message(update, context, text_override=None):
    flow = state(context).get('flow')
    text = text_override if text_override is not None else (update.effective_message.text or '').strip()
    if text_override is None:
        # Replies to typed input must be visible below the user's latest answer.
        state(context).pop('message', None)
    if text == '❌ Bekor':
        return await cancel(update, context)
    if text in ADMIN_KEYS or text in USER_KEYS or text == '🏠 Menyu':
        if text in ADMIN_KEYS and not admin(update, context):
            await update.effective_message.reply_text('Bu amal faqat adminlar uchun.')
            return WAIT
        clear_flow(context)
        if text == '🏠 Menyu':
            return await cancel(update, context)
        await dispatch(update, context, (ADMIN_KEYS | USER_KEYS)[text], [])
        return WAIT
    if not flow:
        await home(update, context)
        return WAIT
    if flow['kind'] not in ('appeal', 'join') and not admin(update, context):
        await home(update, context)
        return WAIT
    try:
        kind = flow['kind']
        if kind == 'upload':
            await import_document(update, context, flow['type'])
        elif kind == 'join':
            contact = update.effective_message.contact
            if not contact or contact.user_id != update.effective_user.id:
                raise ValueError('O‘z kontaktingizni pastdagi telefon yuborish tugmasi orqali ulashing.')
            request_id = db(context).join_by_phone(phone_number(contact.phone_number), update.effective_user.id)
            if request_id:
                legacy.notify_join(context, request_id)
            clear_flow(context)
            await show(update, context, 'Ulanish so‘rovi yuborildi. Admin tasdig‘ini kuting.' if request_id else 'Ulanishingiz tasdiqlangan.')
        elif kind == 'clock':
            await save_hour(update, context, text)
            state(context).pop('flow', None)
        elif kind == 'date':
            date = datetime.strptime(text, '%Y-%m-%d').date()
            if date < datetime.now(context.bot_data['settings'].timezone).date():
                raise ValueError('O‘tgan sanani tanlamang. Masalan: 2026-12-01')
            state(context).pop('flow', None)
            await hours(update, context, flow['target'], date.isoformat())
        elif kind in ('employee', 'appeal'):
            values = flow['values']; index = len(values)
            if not text:
                raise ValueError('Matn yozing.')
            if index == 0 and len(text) > 120:
                raise ValueError('Ism 1–120 belgi bo‘lsin.')
            value = text
            if index == 1:
                value = phone_number(text)
                if kind == 'employee' and any(value in (r['phone'], r['phone_extra']) and r['name'].casefold() != values[0].casefold()
                                               for r in db(context).employee_list()):
                    raise ValueError('Bu telefon boshqa xodimga tegishli. Boshqa raqam yozing.')
            if kind == 'employee':
                if index == 2:
                    value = '' if text == '-' else text.lower()
                    if value and not (value.isdigit() and int(value) > 0 or re.fullmatch(r'@[a-zA-Z][a-zA-Z0-9_]{4,31}', value)):
                        raise ValueError('Telegram ID, @username yoki - yozing. Masalan: @vali_ali')
                if index == 3:
                    value = shift_value('' if text == '-' else text)
                if index == 4:
                    value = None if text.lower().replace('‘', "'").replace('’', "'") in ('-', "yo'q", 'yoq') else phone_number(text)
                if index == 4:
                    name, phone, reference, shift = values
                    db(context).import_employees([dict(name=name, phone=phone, phone_extra=value, telegram_ref=reference or 'phone:'+phone, shift=shift, row_number=1)])
                    clear_flow(context)
                    await show(update, context, 'Xodim saqlandi. Xodim Ulanish tugmasi orqali kontaktini yuborsin, so‘ng uni tasdiqlang.', back='staff')
                    return WAIT
            elif index == 2:
                if len(text) > 1500:
                    raise ValueError('Murojaat 1–1500 belgi bo‘lsin.')
                identifier = 'U-' + uuid.uuid4().hex[:12]
                db(context).import_appeals([dict(id=identifier, client_name=values[0], phone=values[1], body=text, due_at=None, employee_ref='', row_number=1, source_user=update.effective_user.id)])
                db(context).apply_default_schedule(context.bot_data['settings'].timezone)
                context.bot_data['service'].notice('new:'+identifier, 'Yangi murojaat qabul qilindi. /menu → Murojaatlar bo‘limida ko‘ring.')
                clear_flow(context)
                await show(update, context, 'Murojaatingiz qabul qilindi. Admin ko‘rib chiqadi.')
                return WAIT
            values.append(value)
            questions = EMPLOYEE_QUESTIONS if kind == 'employee' else APPEAL_QUESTIONS
            extra = [[button(context, 'O‘tkazib yuborish', 'skip')]] if kind == 'employee' and len(values) >= 2 else None
            await prompt(update, context, flow, questions[len(values)], extra=extra)
    except (ValueError, ValidationError) as exc:
        message = '\n'.join(exc.errors) if isinstance(exc, ValidationError) else str(exc)
        if flow['kind'] == 'date':
            message = 'Sana noto‘g‘ri. Masalan: 2026-12-01'
        question = (EMPLOYEE_QUESTIONS if flow['kind'] == 'employee' else APPEAL_QUESTIONS)[len(flow['values'])] if flow['kind'] in ('employee', 'appeal') else 'Qayta kiriting yoki ❌ Bekor tugmasini bosing.'
        await prompt(update, context, flow, message + '\n\n' + question)
    return WAIT


async def import_document(update, context, kind):
    document = update.effective_message.document
    if not document or not (document.file_name or '').lower().endswith('.xlsx'):
        raise ValueError('Faqat .xlsx fayl yuboring. Masalan: xodimlar.xlsx')
    if document.file_size and document.file_size > MAX_FILE_SIZE:
        raise ValueError('Fayl 5 MB dan oshmasin.')
    remote = await document.get_file()
    content = bytes(await remote.download_as_bytearray())
    rows = await asyncio.to_thread(parse_excel, content, kind, context.bot_data['settings'].timezone)
    database = db(context)
    if kind == 'employees':
        database.import_employees(rows)
        text = f'{len(rows)} ta xodim saqlandi. Xodimlar Ulanish tugmasi orqali kontakt yuborsin. Keyin Xodimlar → Ulanish so‘rovlaridan tasdiqlang.'
    else:
        added, skipped = database.import_appeals(rows)
        scheduled = database.apply_default_schedule(context.bot_data['settings'].timezone)
        text = f'{added} ta murojaat saqlandi; {skipped} ta o‘zgarmagan qator. {scheduled} ta murojaat jadvalga olindi.'
    database.assign_due(time.time(), context.bot_data['settings'].timezone)
    clear_flow(context)
    buttons = []
    if kind == 'appeals' and not database.default_schedule():
        text += '\nQachon yuborilsin? Bir marta tanlang.'
        buttons = [[button(context, 'Darhol', 'mode:darhol'), button(context, '09:00', 'mode:09:00')], [button(context, 'Boshqa vaqt', 'clock:global')]]
    await show(update, context, text, buttons, 'staff' if kind == 'employees' else 'appeals')


def register_handlers(application):
    private = filters.ChatType.PRIVATE
    commands = [CommandHandler(['start', 'menu'], start, filters=private), CommandHandler('bekor', cancel, filters=private)]
    inputs = MessageHandler(private & (filters.TEXT & ~filters.COMMAND | filters.Document.ALL | filters.CONTACT), input_message)
    # Navigation is independent of message IDs; conversations handle typed input.
    conversation = ConversationHandler(entry_points=commands + [inputs], states={WAIT: commands + [inputs]},
        fallbacks=[CommandHandler('bekor', cancel, filters=private)], allow_reentry=True)
    application.add_handler(conversation)
    application.add_handler(CallbackQueryHandler(callback))
    application.add_handler(MessageHandler(private & filters.COMMAND, legacy.unknown_command))
    application.add_error_handler(legacy.error_handler)
