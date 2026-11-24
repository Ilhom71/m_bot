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

from db import STATUSES, assignment_problem
from excel_utils import APPEAL_UPLOAD_PROMPT, MAX_FILE_SIZE, ValidationError, appeal_contact, import_notes, parse_datetime, parse_excel, phone_number, shift_value
from . import bot as legacy

WAIT = 1
# Leave room for headings even when every body character uses two UTF-16 units.
BODY_PAGE_SIZE = 1600
STALE = "Bu tugma eskirgan, /menu bosing"
SMS_PASSWORD_UNCONFIGURED = "SMS boshqaruvi qulflangan. Parolni bot egasi terminal orqali belgilaydi."
PUBLIC_ROUTES = {"home", "cancel", "join", "mine", "appeal", "moretext", "own", "done"}
ROUTES = PUBLIC_ROUTES | {"appeals", "list", "pages", "card", "staff", "workers", "worker",
    "add", "upload", "joins", "request", "approve", "reject", "schedule", "mode", "off",
    "status", "extra", "file", "send", "assign", "pick", "day", "date", "clock", "hour",
    "custom", "retry", "unknown", "attempts", "confirm", "execute", "skip", "disconnect", "sms", "smssettings", "smscard"}
ADMIN_KEYS = {'📋 Murojaatlar': 'appeals', '👥 Xodimlar': 'staff', '⏰ Yuborish': 'schedule',
              '📊 Holat': 'status', '⚙️ Sozlamalar': 'extra', '⚙️ Qo‘shimcha': 'extra'}
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
        rows = [['📋 Murojaatlar', '👥 Xodimlar'], ['⏰ Yuborish', '📊 Holat'], ['⚙️ Sozlamalar']]
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
    args = getattr(context, 'args', None) or []
    if args == ['assignment_status']:
        if admin(update, context):
            await dispatch(update, context, 'status', [])
        else:
            await show(update, context, 'Bu bo‘lim faqat adminlar uchun.')
        return WAIT
    if args:
        target = re.fullmatch(r'task_([1-9][0-9]{0,18})', args[0]) if len(args) == 1 else None
        if target:
            try:
                await dispatch(update, context, 'own', [target[1]])
            except ValueError:
                await show(update, context, 'Bu murojaat sizga biriktirilmagan yoki hozir mavjud emas.', back='mine:0')
        else:
            await show(update, context, 'Murojaat havolasi noto‘g‘ri. Murojaatlarim bo‘limidan tanlang.', back='mine:0')
    return WAIT


async def cancel(update, context):
    await home(update, context)
    return WAIT


def appeal_rows(context, kind="all"):
    database = db(context)
    statuses = {r["id"]: r for r in database.list_appeals(
        context.bot_data["settings"].timezone, today=kind == "today", sent_today=kind == "sent_today")}
    with database.connect() as conn:
        rows = [dict(r) for r in conn.execute("SELECT rowid AS number,id,client_name,source_json FROM appeals ORDER BY created_at DESC,rowid DESC")]
    result = [dict(r, **{k: v for k, v in statuses[r["id"]].items() if k != "id"}) for r in rows if r["id"] in statuses]
    if kind in ('unassigned', 'unconnected'):
        result = [row for row in result if row['telegram'] != 'sent' and assignment_problem(row) == kind]
    elif kind == 'completed':
        result = [row for row in result if row['completed_at'] is not None]
    elif kind == 'sent':
        result = [row for row in result if row['telegram'] == 'sent']
    elif kind == 'sent_today':
        result.sort(key=lambda row: row['telegram_updated_at'], reverse=True)
    elif kind == 'unsent':
        result = [row for row in result if row['telegram'] != 'sent' and row['completed_at'] is None]
    elif kind == 'sms':
        result = [row for row in result if row['due_at'] is not None and row['due_at'] <= time.time() and row['sms'] != 'sent']
    return result


def attention_reasons(row):
    if row.get('completed_at') is not None or row['telegram'] == 'sent':
        return []
    reasons = []
    if row['telegram'] == 'failed':
        reasons.append('Telegram: yuborishda xato')
    elif row['telegram'] == 'unknown':
        reasons.append('Telegram: yuborish natijasi noma’lum')
    problem = assignment_problem(row)
    if problem == 'unassigned':
        reasons.append('Xodim biriktirilmagan')
    elif problem == 'unconnected':
        reasons.append('Xodimning ulanishi tasdiqlanmagan')
    if row['due_at'] is None:
        reasons.append('Yuborish vaqti belgilanmagan')
    return reasons


def attention(row):
    return bool(attention_reasons(row))


def appeal_id(context, number):
    with db(context).connect() as conn:
        row = conn.execute("SELECT id FROM appeals WHERE rowid=?", (int(number),)).fetchone()
    if not row:
        raise ValueError(STALE)
    return row[0]


def snapshot(context, number):
    identifier = appeal_id(context, number)
    return hashlib.sha256(json.dumps(db(context).details(identifier), sort_keys=True).encode()).hexdigest()


def worker_snapshot(context, identifier):
    worker = next((r for r in db(context).employee_list() if r['id'] == int(identifier)), None)
    if worker is None:
        raise ValueError(STALE)
    return hashlib.sha256(json.dumps(worker, sort_keys=True).encode()).hexdigest()


async def listing(update, context, kind, page=0):
    rows = appeal_rows(context, kind)
    if kind == "attention":
        rows = [r for r in rows if attention(r)]
    total = max(1, (len(rows) + 4) // 5)
    page = max(0, min(int(page), total - 1))
    buttons = []
    for row in rows[page*5:page*5+5]:
        status = '🏁 Bajarildi' if row['completed_at'] is not None else '✅ Yuborildi' if row['telegram'] == 'sent' else '⚠️ E’tibor kerak' if attention(row) else '⏳ Kutilmoqda'
        if kind == 'sms':
            status = 'SMS: ' + STATUSES[row['sms']]
            if row['sms_held']:
                status = 'SMS: qo‘lda yuboriladi' if row['sms'] != 'unknown' else 'SMS: natija noma’lum'
        label = appeal_contact(row)[1] or f"Murojaat #{row['id']}"
        route = 'smscard' if kind == 'sms' else 'card'
        buttons.append([button(context, f"{status} — {label[:80]}", f"{route}:{row['number']}")])
    if total > 1:
        buttons.append([button(context, f"◀️ {page+1}/{total} ▶️", f"pages:{kind}:{page}:list")])
    state(context)['sms_back' if kind == 'sms' else 'list_back'] = f"list:{kind}:{page}"
    title = {'today': 'Bugungi', 'sent_today': 'Bugun yuborilgan', 'all': 'Barcha', 'unassigned': 'Mos xodim topilmagan',
             'unconnected': 'Xodimi ulanmagan', 'completed': 'Bajarilgan', 'sent': 'Yuborilgan',
             'unsent': 'Yuborilmagan'}.get(kind, 'E’tibor talab qiladigan')
    text = f"{title} murojaatlar: {len(rows)}"
    if kind == 'unconnected':
        text += f"\nBotga ulanmagan xodimlar: {len({row['employee_id'] for row in rows})} ta"
    if kind == 'sms':
        text = f"SMSda bormaganlar: {len(rows)}\nVaqti kelgan va SMS qabul qilingani tasdiqlanmagan murojaatlar.\nNoma’lum natijani qabul qiluvchidan tekshiring."
    back = {'sms': 'smssettings', 'unassigned': 'status', 'unconnected': 'status'}.get(kind, 'appeals')
    await show(update, context, text, buttons, back)


async def card(update, context, number):
    identifier = appeal_id(context, number)
    row, channels = db(context).details(identifier)
    employees = {e["id"]: e for e in db(context).employee_list()}
    employee = employees.get(row["employee_id"])
    zone = context.bot_data["settings"].timezone
    when = datetime.fromtimestamp(row["due_at"], zone).strftime("%d.%m.%Y %H:%M") if row["due_at"] else "belgilanmagan"
    label, name = appeal_contact(row)
    contact_line = f"{label}: {name[:150]}\n" if name else ""
    text = (f"Murojaat #{identifier}\n{contact_line}Telefon: {row['phone']}\n"
            f"Xodim: {employee['name'] if employee else row['employee_hint'] or 'biriktirilmagan'}\nVaqt: {when}\n\n")
    if row['completed_at'] is not None:
        text += 'Ijro: ✅ Bajarildi — ' + datetime.fromtimestamp(row['completed_at'], zone).strftime('%d.%m.%Y %H:%M') + '\n\n'
    telegram = next(c for c in channels if c['channel'] == 'telegram')
    summary = 'Holat: ' + ('✅ Yuborildi' if telegram['status'] == 'sent' else STATUSES[telegram['status']])
    status_row = dict(row, **{c['channel']: c['status'] for c in channels},
                      employee_approved=employee['approved'] if employee else False,
                      employee_telegram_id=employee['telegram_id'] if employee else None)
    reasons = attention_reasons(status_row)
    if reasons:
        text += '⚠️ E’tibor sababi:\n' + '\n'.join(reasons) + '\n\n'
    body = row['body']
    text += body[:BODY_PAGE_SIZE] + ("\n… Davomi tugmasida." if len(body) > BODY_PAGE_SIZE else "") + "\n\n" + summary
    buttons = []
    unlocked = row['completed_at'] is None and all(c["attempts"] == 0 and c["status"] not in ("sent", "sending", "unknown") for c in channels)
    if unlocked:
        buttons += [[button(context, "📨 Hozir yuborish", f"send:{number}"), button(context, "📅 Vaqt belgilash", f"day:{number}")],
                    [button(context, "👤 Xodim biriktirish", f"assign:{number}:0")]]
    if telegram['status'] == 'failed' and row['completed_at'] is None:
        buttons.append([button(context, '🔄 Qayta urinish', f'retry:{number}:telegram')])
    if telegram['status'] == 'unknown':
        buttons.append([button(context, 'Natijani belgilash', f'unknown:{number}:telegram')])
    if len(body) > BODY_PAGE_SIZE:
        buttons.append([button(context, "Matn davomi", f"moretext:{number}:1")])
    await show(update, context, text, buttons, state(context).get("list_back", "list:all:0"))


async def sms_card(update, context, number):
    identifier = appeal_id(context, number)
    row, channels = db(context).details(identifier)
    channel = next(c for c in channels if c['channel'] == 'sms')
    employee = next((e for e in db(context).employee_list() if e['id'] == row['employee_id']), None)
    text = f"SMS — Murojaat #{identifier}\nXodim: {employee['name'] if employee else 'biriktirilmagan'}\nHolat: {STATUSES[channel['status']]}"
    enabled = db(context).sms_enabled()
    text += '\nSMS yuborish: ' + ('yoqilgan' if enabled else 'o‘chirilgan')
    recipients = db(context).sms_details(identifier)
    if not recipients and not row['sms_held']:
        text += '\nSMS hali navbatga olinmagan. Yuborish vaqti va xodim ulanishini tekshiring.'
    for recipient in recipients:
        text += f"\n\n{recipient['phone']}\n{STATUSES[recipient['status']]} | urinish: {recipient['attempts']}"
        if recipient['last_error']:
            text += '\nSabab: ' + recipient['last_error']
    if channel['status'] == 'unknown':
        text += '\n\nNatija noma’lum: qayta yuborishdan oldin qabul qiluvchidan tekshiring.'
    if row['completed_at'] is not None:
        text += '\n\nMurojaat bajarilgan; qayta yuborilmaydi.'
    buttons = []
    if row['sms_held']:
        text += '\n\nSMS qayta yoqilganda yuborish vaqti o‘tgan edi. Bu SMS avtomatik yuborilmaydi.'
    if row['completed_at'] is None and row['sms_held'] and enabled and any(
            entry['status'] in ('pending', 'retry', 'failed', 'blocked') for entry in (recipients or [channel])):
        buttons.append([button(context, '📨 SMSni navbatga qo‘yish', f'retry:{number}:sms')])
    elif row['completed_at'] is None and not row['sms_held'] and channel['status'] in ('failed', 'blocked'):
        buttons.append([button(context, '🔄 SMS qayta urinish', f'retry:{number}:sms')])
    if channel['status'] == 'unknown':
        buttons.append([button(context, 'SMS natijasini belgilash', f'unknown:{number}:sms')])
    await show(update, context, text, buttons, state(context).get('sms_back', 'list:sms:0'))


async def choose_workers(update, context, number=None, page=0):
    rows = db(context).employee_list()
    total = max(1, (len(rows) + 4) // 5)
    page = max(0, min(int(page), total - 1))
    visible = rows[page*5:page*5+5]
    notes = {r['id']: '' if r['approved'] and r['telegram_id'] else ' — ulanmagan' for r in visible}
    buttons = [[button(context, f"{page*5+i}. {r['name'][:48]}{notes[r['id']]}",
                       f"pick:{number}:{r['id']}" if number is not None else f"worker:{r['id']}")]
               for i, r in enumerate(visible, 1)]
    if total > 1:
        route = f"assign:{number}" if number is not None else "workers"
        buttons.append([button(context, f"◀️ Sahifa {page+1}/{total} ▶️", f"pages:staff:{page}:{route}")])
    if number is None:
        state(context)['workers_back'] = f'workers:{page}'
    text = (f"👥 Xodimlar — jami {len(rows)} ta\nSahifa {page+1}/{total}\n\n" +
            "\n\n".join(f"{page*5+i}. {r['name']}{notes[r['id']]}" for i, r in enumerate(visible, 1)) +
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
        snapshot=worker_snapshot(context, number) if action == 'disconnect' else snapshot(context, number) if action != "reject" else None)
    question = 'Xodimning bot bilan ulanishi uzilsinmi?' if action == 'disconnect' else "Ishonchingiz komilmi? Qabul qiluvchidan natijani tekshiring. Bu amal qayta yuborishga olib kelishi mumkin."
    back = f'worker:{number}' if action == 'disconnect' else 'joins:0' if action == 'reject' else f'smscard:{number}' if channel == 'sms' else f'card:{number}'
    await show(update, context, question,
               [[button(context, "✅ Ha", "execute"), button(context, "❌ Yo‘q", back)]])


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
            (("Bugun", "list:today:0"), ("Bugun yuborilgan murojaatlar", "list:sent_today:0"),
             ("Yuborilgan", "list:sent:0"), ("Yuborilmagan", "list:unsent:0"),
             ("Hammasi", "list:all:0"), ("Excel yuklash", "upload:appeals"))])
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
        await card(update, context, args[0])
    elif route == 'smscard':
        await sms_card(update, context, args[0])
    elif route == 'moretext':
        identifier = appeal_id(context, args[0])
        row = database.details(identifier)[0]
        if not admin(update, context) and not any(r['id'] == identifier for r in database.export_appeals(telegram_id=update.effective_user.id)):
            raise ValueError(STALE)
        page = int(args[1]); text = row['body'][page*BODY_PAGE_SIZE:(page+1)*BODY_PAGE_SIZE]
        buttons = [[button(context, 'Davomi ▶️', f"moretext:{args[0]}:{page+1}")]] if len(row['body']) > (page+1)*BODY_PAGE_SIZE else []
        await show(update, context, text or 'Matn tugadi.', buttons, f'card:{args[0]}' if admin(update, context) else 'mine:0')
    elif route == 'staff':
        await show(update, context, "Xodimlar", [[button(context, label, action)] for label, action in
            (("Ro‘yxat", "workers:0"), ("Qo‘shish", "add"), ("Excel yuklash", "upload:employees"),
             ("Excelga chiqarish", "file:employees"),
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
            buttons = [[button(context, 'Ulanishni uzish', f"disconnect:{employee['id']}")]] if employee['telegram_id'] or employee['approved'] else []
            await show(update, context, f"👤 {employee['name']}\n\n📱 Asosiy telefon: {employee['phone']}\nQo‘shimcha telefon: {employee['phone_extra'] or 'yo‘q'}\n\nUlanish: {'✅ Ulangan' if employee['approved'] and employee['telegram_id'] else '❌ Ulanmagan'}\nSmena: {employee['shift'] or 'cheklanmagan'}", buttons, back=ui.get('workers_back', 'workers:0'))
    elif route == 'disconnect':
        await confirm_action(update, context, 'disconnect', args[0])
    elif route == 'add':
        await prompt(update, context, {'kind': 'employee', 'values': []}, EMPLOYEE_QUESTIONS[0], 'staff')
    elif route == 'upload':
        kind = args[0]
        if kind not in ('appeals', 'employees'):
            raise ValueError(STALE)
        message = APPEAL_UPLOAD_PROMPT if kind == 'appeals' else 'Excel (.xlsx) faylini yuboring. Namuna — xodimlar.xlsx'
        await prompt(update, context, {'kind': 'upload', 'type': kind}, message, 'appeals' if kind == 'appeals' else 'staff')
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
        sms_enabled = database.sms_enabled()
        text = 'Hozir avtomatik: yoqilgan' if mode == 'darhol' else f'Har kuni {mode}' if mode else 'Avtomatik yuborish: o‘chirilgan'
        text += '\nSMS yuborish: ' + ('✅ Yoqilgan' if sms_enabled else '❌ O‘chirilgan')
        text += '\nSMS qayta yoqilganda vaqti o‘tgan SMS’lar faqat qo‘lda navbatga qo‘yiladi.'
        await show(update, context, text + '\nExcelda alohida vaqt bo‘lsa, o‘sha vaqt ishlatiladi.',
            [[button(context, 'Darhol', 'mode:darhol')], [button(context, 'Soat tanlash', 'clock:global')], [button(context, 'O‘chirish', 'off')],
             [button(context, '❌ SMS o‘chirish' if sms_enabled else '✅ SMS yoqish', 'sms:off' if sms_enabled else 'sms:on')]])
    elif route == 'sms':
        if len(args) not in (1, 2) or args[0] not in ('on', 'off') or (len(args) == 2 and args[1] not in ('extra', 'schedule', 'smssettings')):
            raise ValueError(STALE)
        back = args[1] if len(args) == 2 else 'schedule'
        if not database.sms_password_configured():
            await show(update, context, SMS_PASSWORD_UNCONFIGURED, back=back)
            return
        action = 'yoqish' if args[0] == 'on' else 'o‘chirish'
        await prompt(update, context, dict(kind='sms_password', enabled=args[0] == 'on', back=back,
                                          expires=time.monotonic() + 120),
                     f"SMS’ni {action} uchun parolni yuboring. So‘rov 2 daqiqa amal qiladi.", back)
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
        records = appeal_rows(context)
        count = sum(attention(r) for r in records)
        text = database.status_text(include_sms=False).replace('/jadval ni tanlang', 'Yuborish bo‘limidan tanlang') + f'\n\nE’tibor talab qiladigan murojaatlar: {count}'
        buttons = []
        if count:
            reasons = {}
            for row in records:
                for reason in attention_reasons(row):
                    reasons[reason] = reasons.get(reason, 0) + 1
            text += '\nSabablar:\n' + '\n'.join(f'{reason}: {number} ta' for reason, number in reasons.items())
            buttons.append([button(context, f'⚠️ E’tibor kerak ({count})', 'list:attention:0')])
        else:
            text += '\n✅ Yuborishda e’tibor talab qiladigan muammo yo‘q.'
        await show(update, context, text, buttons +
            [[button(context, 'Mos xodim topilmagan', 'list:unassigned:0')],
             [button(context, 'Xodimi ulanmagan', 'list:unconnected:0')],
             [button(context, '✅ Bajarilgan', 'list:completed:0')]])
    elif route == 'extra':
        sms_enabled = database.sms_enabled()
        text = '⚙️ Sozlamalar\n\nSMS yuborish: ' + ('✅ Yoqilgan' if sms_enabled else '❌ O‘chirilgan')
        text += '\nSMS qayta yoqilganda vaqti o‘tgan SMS’lar faqat qo‘lda navbatga qo‘yiladi.\n\nShablonlar va Excel tahrirlash'
        sms_button = button(context, '❌ SMS o‘chirish' if sms_enabled else '✅ SMS yoqish',
                            'sms:off:extra' if sms_enabled else 'sms:on:extra')
        await show(update, context, text, [[button(context, '📱 SMS sozlamalari', 'smssettings')], [sms_button]] + [[button(context, label, action)] for label, action in
            (('Bo‘sh shablonlar', 'file:templates'), ('Xodimlarni Excel olish', 'file:employees'), ('Murojaatlarni Excel olish', 'file:appeals'))])
    elif route == 'smssettings':
        enabled = database.sms_enabled()
        rows = appeal_rows(context, 'sms')
        text = '📱 SMS sozlamalari\nSMS yuborish: ' + ('✅ Yoqilgan' if enabled else '❌ O‘chirilgan')
        text += '\nSMS qayta yoqilganda vaqti o‘tgan SMS’lar faqat qo‘lda navbatga qo‘yiladi.'
        text += f'\n\nSMSda bormaganlar: {len(rows)}\nBu ro‘yxatda faqat yuborish vaqti kelgan murojaatlar bor.'
        await show(update, context, text,
            [[button(context, '❌ SMS o‘chirish' if enabled else '✅ SMS yoqish', 'sms:off:smssettings' if enabled else 'sms:on:smssettings')],
             [button(context, f'SMSda bormaganlar ({len(rows)})', 'list:sms:0')]], 'extra')
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
            button(context, 'Yuborilmadi', f'confirm:{args[0]}:{args[1]}:no')]], f'smscard:{args[0]}' if args[1] == 'sms' else f'card:{args[0]}')
    elif route == 'confirm':
        await confirm_action(update, context, 'unknown', args[0], args[1], args[2] == 'yes')
    elif route == 'execute':
        pending = ui.pop('pending')
        number = pending['number']
        if pending['action'] == 'disconnect':
            if worker_snapshot(context, number) != pending['snapshot']:
                raise ValueError(STALE)
            database.disconnect_employee(int(number))
            await dispatch(update, context, 'worker', [number])
        elif pending['action'] == 'reject':
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
            if pending['channel'] == 'sms':
                await sms_card(update, context, number)
            else:
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
            name = appeal_contact(row)[1] or f"Murojaat #{row['id']}"
            buttons.append([button(context, name[:80], f'own:{number}')])
        if page > 0:
            buttons.append([button(context, '◀️ Oldingi', f'mine:{page-1}')])
        if (page+1)*4 < len(rows):
            buttons.append([button(context, 'Keyingi ▶️', f'mine:{page+1}')])
        await show(update, context, 'Murojaatingizni tanlang.' if rows else 'Sizga biriktirilgan va vaqti kelgan murojaat yo‘q.', buttons)
    elif route == 'done':
        identifier = appeal_id(context, args[0])
        database.complete_appeal(identifier, update.effective_user.id, settings.admin_ids)
        await dispatch(update, context, 'own', args)
    elif route == 'own':
        identifier = appeal_id(context, args[0])
        row = next(iter(database.export_appeals(telegram_id=update.effective_user.id, appeal_id=identifier)), None)
        if not row:
            raise ValueError(STALE)
        buttons = [[button(context, 'Davomi ▶️', f'moretext:{args[0]}:1')]] if len(row['body']) > BODY_PAGE_SIZE else []
        if row['completed_at'] is None:
            buttons.append([button(context, '✅ Bajarildi', f'done:{args[0]}')])
        progress = '✅ Bajarildi' if row['completed_at'] is not None else '⏳ Kutilmoqda'
        label, name = appeal_contact(row)
        contact_line = f"{label}: {name[:150]}\n" if name else ""
        await show(update, context, f"Murojaat #{row['id']}\n{contact_line}Telefon: {row['phone']}\nIjro: {progress}\n\n{row['body'][:BODY_PAGE_SIZE]}", buttons, 'mine:0')


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
    if flow and flow['kind'] == 'sms_password':
        return await sms_password_input(update, context, flow)
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
            request_id = db(context).join_by_phone(contact.phone_number, update.effective_user.id,
                                                   update.effective_user.username)
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


async def sms_password_input(update, context, flow):
    clear_flow(context)
    if not admin(update, context):
        await home(update, context)
        return WAIT
    message = update.effective_message
    if message.text is None:
        await prompt(update, context, flow, "SMS parolini matn sifatida yuboring.", flow['back'])
        return WAIT
    try:
        deleted = await message.delete()
    except TelegramError:
        deleted = False
    if not deleted:
        await show(update, context, "Parol xabari o‘chmadi. Uni chatdan o‘chiring va amalni qaytadan boshlang.",
                   back=flow['back'])
        return WAIT
    if time.monotonic() >= flow['expires']:
        await show(update, context, "Parol so‘rovi muddati tugadi. SMS tugmasini qayta bosing.", back=flow['back'])
        return WAIT
    result = await asyncio.to_thread(db(context).change_sms_with_password, flow['enabled'], message.text,
                                     update.effective_user.id)
    if result == 'changed':
        await dispatch(update, context, flow['back'], [])
    elif result == 'invalid':
        await prompt(update, context, flow, "Parol noto‘g‘ri. Qayta kiriting.", flow['back'])
    else:
        text = ("Urinishlar chegarasiga yetdingiz. 5 daqiqadan keyin qayta urinib ko‘ring."
                if result == 'locked' else SMS_PASSWORD_UNCONFIGURED)
        await show(update, context, text, back=flow['back'])
    return WAIT


async def import_document(update, context, kind):
    document = update.effective_message.document
    if not document or not (document.file_name or '').lower().endswith('.xlsx'):
        raise ValueError('Faqat .xlsx fayl yuboring.')
    if document.file_size and document.file_size > MAX_FILE_SIZE:
        raise ValueError('Fayl 5 MB dan oshmasin.')
    remote = await document.get_file()
    content = bytes(await remote.download_as_bytearray())
    rows = await asyncio.to_thread(parse_excel, content, kind, context.bot_data['settings'].timezone,
                                   filename=document.file_name)
    database = db(context)
    if kind == 'employees':
        database.import_employees(rows)
        text = f'{len(rows)} ta xodim saqlandi. Xodimlar Ulanish tugmasi orqali kontakt yuborsin. Keyin Xodimlar → Ulanish so‘rovlaridan tasdiqlang.'
    else:
        added, skipped = database.import_appeals(rows)
        scheduled = database.apply_default_schedule(context.bot_data['settings'].timezone)
        text = f'{added} ta murojaat saqlandi; {skipped} ta o‘zgarmagan qator. {scheduled} ta murojaat jadvalga olindi.'
        text += '\n\n' + import_notes(rows)
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
