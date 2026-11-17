"""Uzbek-language private-chat workflows."""
import asyncio
import logging
import time
import uuid
from functools import wraps
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, KeyboardButton, ReplyKeyboardMarkup, Update
from telegram.ext import CallbackQueryHandler, CommandHandler, MessageHandler, filters

from db import STATUSES
from excel_utils import MAX_FILE_SIZE, ValidationError, appeals_workbook, build_workbook, parse_datetime, parse_excel, phone_number
from scheduler import appeal_list_text, pages

logger = logging.getLogger(__name__)

# Compatibility helpers; production navigation lives in handlers.ui.
ADMIN_MENU = ReplyKeyboardMarkup([["Murojaatlar", "Xodimlar"], ["Avtomatik yuborish", "Holat"], ["Yordam"]], resize_keyboard=True)
USER_MENU = ReplyKeyboardMarkup([["Xodim sifatida ulanish", "Mening murojaatlarim"], ["Murojaat yuborish"]], resize_keyboard=True)
HELP = "Ochiq buyruqlar: /start, /menu, /bekor. Kerakli amalni /menu tugmalaridan tanlang."


async def reply(update, text, **kwargs):
    for part in pages(text):
        await update.effective_message.reply_text(part, **kwargs)


def admin_only(func):
    @wraps(func)
    async def guarded(update, context, *args, **kwargs):
        if update.effective_chat.type != "private":
            await reply(update, "Bu buyruqni botning shaxsiy chatida yuboring.")
            return
        if update.effective_user.id not in context.bot_data["settings"].admin_ids:
            await reply(update, "Bu amal faqat adminlar uchun. ID raqamingiz: " + str(update.effective_user.id))
            return
        return await func(update, context, *args, **kwargs)
    return guarded


async def start(update, context):
    if update.effective_chat.type != "private":
        await reply(update, "Bot bilan shaxsiy chatda ishlang.")
        return
    user = update.effective_user
    context.bot_data["db"].register_user(user.id, user.username)
    context.user_data.clear()
    is_admin = user.id in context.bot_data["settings"].admin_ids
    text = ("Assalomu alaykum! Admin menyusi tayyor.\n" + HELP if is_admin else
            "Assalomu alaykum! Murojaat uchun /murojaat ni bosing.\n"
            "Xodim bo'lsangiz, 'Xodim sifatida ulanish' tugmasini bosing. Admin tasdiqlagach Excel murojaatlarni olasiz.\n"
            f"Telegram ID: {user.id}")
    if not context.bot_data["settings"].admin_ids:
        text += "\nAdmin hali sozlanmagan. Loyiha egasi .env faylidagi ADMIN_IDS ni to'ldirsin."
    await reply(update, text, reply_markup=ADMIN_MENU if is_admin else USER_MENU)


async def identity(update, context):
    await reply(update, f"Telegram ID raqamingiz: {update.effective_user.id}")


@admin_only
async def help_command(update, context):
    await reply(update, HELP)


@admin_only
async def upload_appeals(update, context):
    context.user_data.clear()
    context.user_data["upload"] = "appeals"
    await reply(update, "Murojaatlar .xlsx faylini yuboring. Ustunlar: id, mijoz_ismi, telefon, murojaat_matni, yuborish_vaqti; xodim — ixtiyoriy.\n"
                "Vaqt katagi bo'sh bo'lsa, umumiy avtomatik jadval ishlatiladi. Jadvalni 'Avtomatik yuborish' tugmasida bir marta tanlang. O'tgan vaqtli qatorlar darhol navbatga tushadi. /bekor — bekor qilish.")


@admin_only
async def upload_employees(update, context):
    context.user_data.clear()
    context.user_data["upload"] = "employees"
    await reply(update, "Xodimlar .xlsx faylini yuboring. Ustunlar: ism, telefon, telegram_id; smena_vaqti — ixtiyoriy.\n"
                "telegram_id: raqamli ID yoki @username; bo'sh qoldirish mumkin. Xodim telefonini ulashib, admin tasdig'ini oladi. Smena: 09:00-18:00. /bekor — bekor qilish.")


@admin_only
async def document(update, context):
    kind = context.user_data.get("upload")
    if not kind:
        await reply(update, "Avval /yuklash_murojaat yoki /yuklash_xodim buyrug'ini tanlang.")
        return
    doc = update.effective_message.document
    if not (doc.file_name or "").lower().endswith(".xlsx"):
        await reply(update, "Faqat .xlsx qabul qilinadi. /shablon orqali namuna oling.")
        return
    if doc.file_size and doc.file_size > MAX_FILE_SIZE:
        await reply(update, "Fayl 5 MB dan oshmasin.")
        return
    try:
        remote = await doc.get_file()
        data = bytes(await remote.download_as_bytearray())
        rows = await asyncio.to_thread(parse_excel, data, kind, context.bot_data["settings"].timezone)
        db = context.bot_data["db"]
        if kind == "employees":
            db.import_employees(rows)
            message = f"{len(rows)} ta xodim saqlandi. Xodimlar /start yoki /xodim_ulan orqali ulansin. /tasdiqlar orqali tasdiqlang."
        else:
            added, skipped = db.import_appeals(rows)
            scheduled = db.apply_default_schedule(context.bot_data["settings"].timezone)
            message = f"{added} ta murojaat saqlandi (yangi yoki tahrirlangan); {skipped} ta o'zgarmagan qator takrorlanmadi.\n{scheduled} ta murojaat umumiy jadvalga avtomatik qo'yildi."
            if db.default_schedule():
                message += "\nVaqti kelganda bot tasdiqlangan xodimlarga o'zi yuboradi. ID kiritish shart emas."
            if any(row.get("legacy") for row in rows):
                message += "\nEski jadval tanildi. Kirish raqami matnda saqlandi; bot ID si L- bilan boshlanadi. 'Muddati' yuborish vaqti sifatida olinmadi."
                message += "\nUzun matnlar to'liq saqlandi. Bo'sh arizachi 'Ko'rsatilmagan' deb belgilandi. Noma'lum xodimlar biriktirishni kutadi."
        db.assign_due(time.time(), context.bot_data["settings"].timezone)
        context.user_data.pop("upload", None)
        await reply(update, message + "\n/bugun yoki /royxat orqali tekshiring.")
        if kind == "appeals" and not db.default_schedule() and any(row["due_at"] is None for row in rows):
            context.args = []
            await automatic_schedule(update, context)
    except ValidationError as exc:
        await reply(update, "Fayl qabul qilinmadi; hech bir qator saqlanmadi.\n" + "\n".join(exc.errors))
    except Exception as exc:
        logger.error("Upload failed error=%s", type(exc).__name__)
        await reply(update, "Faylni qayta ishlashda xato. Qayta urinib ko'ring; /holat orqali natijani tekshiring.")


@admin_only
async def status(update, context):
    await reply(update, context.bot_data["db"].status_text())


@admin_only
async def today(update, context):
    settings, db = context.bot_data["settings"], context.bot_data["db"]
    db.assign_due(time.time(), settings.timezone)
    await reply(update, appeal_list_text(db.list_appeals(settings.timezone, today=True), settings.timezone, "Bugungi yuborishlar"))


@admin_only
async def all_appeals(update, context):
    settings, db = context.bot_data["settings"], context.bot_data["db"]
    await reply(update, appeal_list_text(db.list_appeals(settings.timezone), settings.timezone, "Murojaatlar ro'yxati"))


@admin_only
async def employees(update, context):
    from .ui import choose_workers
    await choose_workers(update, context)


@admin_only
async def set_time(update, context):
    if not context.args:
        await begin_input(update, context, "vaqt")
        return
    if len(context.args) != 3:
        await reply(update, "Namuna: /vaqt A001 2026-12-01 09:00")
        return
    try:
        when = parse_datetime(" ".join(context.args[1:]), context.bot_data["settings"].timezone)
        context.bot_data["db"].set_time(context.args[0], when.timestamp())
        await reply(update, "Yuborish vaqti saqlandi: " + when.strftime("%d.%m.%Y %H:%M") + ". O'tgan vaqt bo'lsa, yaqin navbatda yuboriladi.")
    except ValueError as exc:
        await reply(update, str(exc))


@admin_only
async def assign(update, context):
    if not context.args:
        await begin_input(update, context, "biriktir")
        return
    if len(context.args) < 2:
        await reply(update, "Namuna: /biriktir A001 Xodim ismi")
        return
    try:
        context.bot_data["db"].assign(context.args[0], " ".join(context.args[1:]))
        await reply(update, "Xodim biriktirildi.")
    except ValueError as exc:
        await reply(update, str(exc))


@admin_only
async def details(update, context):
    if not context.args:
        await begin_input(update, context, "batafsil")
        return
    result = context.bot_data["db"].details(context.args[0]) if len(context.args) == 1 else None
    if not result:
        await reply(update, "Murojaat topilmadi. Namuna: /batafsil A001")
        return
    row, channels = result
    text = f"#{row['id']}\n{row['client_name']}\n{row['phone']}\n{row['body']}\n"
    text += "\n".join(f"{c['channel']}: {STATUSES[c['status']]} | urinish: {c['attempts']} | {c['last_error'] or 'xato yo‘q'}" for c in channels)
    text += "\n" + "\n".join(f"SMS {c['phone']}: {STATUSES[c['status']]} | urinish: {c['attempts']} | {c['last_error'] or 'xato yo‘q'}"
                               for c in context.bot_data['db'].sms_details(row['id']))
    await reply(update, text)


@admin_only
async def retry(update, context):
    if not context.args:
        await begin_input(update, context, "qayta")
        return
    if len(context.args) != 2 or context.args[1] not in ("telegram", "sms"):
        await reply(update, "Namuna: /qayta A001 telegram")
        return
    try:
        context.bot_data["db"].retry(*context.args)
        await reply(update, "Qayta yuborish navbatga qo'yildi.")
    except ValueError as exc:
        await reply(update, str(exc))


@admin_only
async def confirm(update, context):
    if not context.args:
        await begin_input(update, context, "tasdiq")
        return
    if len(context.args) != 3 or context.args[1] not in ("telegram", "sms") or context.args[2] not in ("yuborildi", "yuborilmadi"):
        await reply(update, "Qabul qiluvchidan tekshirgandan keyin:\n/tasdiq A001 telegram yuborildi\nyoki\n/tasdiq A001 telegram yuborilmadi\n"
                    "'yuborilmadi' tanlansa, qayta yuboriladi.")
        return
    try:
        context.bot_data["db"].resolve_unknown(context.args[0], context.args[1], context.args[2] == "yuborildi")
        await reply(update, "Tekshirilgan natija saqlandi.")
    except ValueError as exc:
        await reply(update, str(exc))


@admin_only
async def templates(update, context):
    for filename, columns in (
        ("murojaatlar.xlsx", ["id", "mijoz_ismi", "telefon", "murojaat_matni", "yuborish_vaqti", "xodim"]),
        ("xodimlar.xlsx", ["ism", "telefon", "telegram_id", "smena_vaqti", "qoshimcha_telefon"]),
    ):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = filename.removesuffix(".xlsx")
        sheet.append(columns)
        sheet.freeze_panes = "A2"
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="245A81")
            cell.alignment = Alignment(wrap_text=True)
            sheet.column_dimensions[cell.column_letter].width = 24 if cell.value != "murojaat_matni" else 55
        sheet.auto_filter.ref = sheet.dimensions
        stream = BytesIO()
        workbook.save(stream)
        stream.seek(0)
        await update.effective_message.reply_document(document=stream, filename=filename)


async def cancel(update, context):
    context.user_data.clear()
    menu = ADMIN_MENU if update.effective_user.id in context.bot_data["settings"].admin_ids else USER_MENU
    await reply(update, "Joriy kiritish bekor qilindi.", reply_markup=menu)


async def new_appeal(update, context):
    context.user_data.clear()
    context.user_data["form"] = {"step": "name"}
    await reply(update, "Ism-familiyangizni yozing. /bekor — to'xtatish.")


async def text_message(update, context):
    text = update.effective_message.text.strip()
    actions = BUTTON_ACTIONS
    if text in actions:
        context.user_data.pop("input", None)
        context.user_data.pop("schedule", None)
        context.args = []
        await actions[text](update, context)
        return
    if context.user_data.get("schedule"):
        context.args = [text]
        await automatic_schedule(update, context)
        return
    if context.user_data.get("input"):
        await continue_input(update, context, text)
        return
    if context.user_data.get("join"):
        await reply(update, "Pastdagi 'Telefonimni ulashish' tugmasini bosing. Begona raqamni matn orqali yubormang.")
        return
    form = context.user_data.get("form")
    if not form:
        await reply(update, "Menyudan tanlang yoki /start yuboring. Murojaat uchun /murojaat.")
        return
    try:
        if form["step"] == "name":
            if not 1 <= len(text) <= 120:
                raise ValueError("Ism 1–120 belgi bo'lsin.")
            form.update(name=text, step="phone")
            await reply(update, "Telefon raqamingizni yozing: +998901234567")
        elif form["step"] == "phone":
            form.update(phone=phone_number(text), step="body")
            await reply(update, "Murojaatingizni yozing (1–1500 belgi).")
        else:
            if not 1 <= len(text) <= 1500:
                raise ValueError("Murojaat matni 1–1500 belgi bo'lsin.")
            appeal_id = "U-" + uuid.uuid4().hex[:12]
            context.bot_data["db"].import_appeals([dict(
                id=appeal_id, client_name=form["name"], phone=form["phone"], body=text,
                due_at=None, employee_ref="", row_number=1, source_user=update.effective_user.id,
            )])
            context.user_data.pop("form", None)
            scheduled = context.bot_data["db"].apply_default_schedule(context.bot_data["settings"].timezone)
            context.bot_data["service"].notice("new:" + appeal_id,
                f"Yangi murojaat #{appeal_id}.\n/batafsil {appeal_id}\n" +
                ("Umumiy jadval orqali avtomatik yuboriladi." if scheduled else "Avtomatik yuborish uchun /jadval ni sozlang."))
            await reply(update, f"Murojaatingiz qabul qilindi. Raqami: {appeal_id}. Admin ko'rib chiqadi.")
    except ValueError as exc:
        await reply(update, str(exc))


async def unknown_command(update, context):
    await reply(update, "Buyruq topilmadi. /menu orqali kerakli amalni tanlang.")


async def error_handler(update, context):
    # Exception strings and Update objects can contain tokens or personal information.
    logger.error("Handler failed error=%s", type(context.error).__name__)
    if isinstance(update, Update) and update.effective_message:
        try:
            await reply(update, "Amal bajarilmadi. Birozdan so'ng qayta urinib ko'ring.")
        except Exception:
            logger.warning("Could not send error response")


def register_handlers(application):
    from .ui import register_handlers as register_ui
    register_ui(application)


def command_actions():
    return {
        "start": start, "id": identity, "yordam": help_command,
        "yuklash_murojaat": upload_appeals, "yuklash_xodim": upload_employees,
        "holat": status, "bugun": today, "royxat": all_appeals, "xodimlar": employees,
        "vaqt": set_time, "biriktir": assign, "batafsil": details,
        "qayta": retry, "tasdiq": confirm, "shablon": templates,
        "bekor": cancel, "murojaat": new_appeal,
        "xodimlar_excel": export_employees, "murojaatlar_excel": export_all,
        "xodim_ulan": join_employee, "mening_murojaatlar": my_appeals,
        "tasdiqlar": pending_joins, "xodim_tasdiq": approve_join, "xodim_rad": reject_join,
        "hozir": send_now, "jadval": automatic_schedule,
    "xodim_qosh": add_employee,
    }


INPUT_PROMPTS = {
    "vaqt": ["Murojaat ID sini kiriting (/royxat dan olish mumkin).", "Yuborish vaqtini yozing: 2026-12-01 09:00"],
    "biriktir": ["Murojaat ID sini kiriting.", "Xodimning ro'yxatdagi to'liq ismini yoki Telegram ID sini kiriting (/xodimlar)."],
    "batafsil": ["Murojaat ID sini kiriting."],
    "qayta": ["Murojaat ID sini kiriting.", "Kanalni kiriting: telegram yoki sms"],
    "tasdiq": ["Murojaat ID sini kiriting.", "Kanalni kiriting: telegram yoki sms", "Qabul qiluvchidan tekshiring va yozing: yuborildi yoki yuborilmadi"],
    "xodim_tasdiq": ["Ulanish so'rovi raqamini kiriting (/tasdiqlar dan oling)."],
    "xodim_rad": ["Rad etiladigan ulanish so'rovi raqamini kiriting."],
    "xodim_qosh": ["Xodimning to'liq ismini kiriting.", "Xodimning telefon raqamini kiriting (+998901234567).",
                   "Telegram ID yoki username bilsangiz kiriting. Bilmasangiz 'o‘tkazish' deb yozing.",
                   "Smena vaqtini HH:MM-HH:MM shaklida kiriting. Shartsiz smena uchun 'o‘tkazish' deb yozing.",
                   "Qo'shimcha telefon raqamini kiriting. Raqam bo'lmasa yo'q yoki - yozing."],
}


async def begin_input(update, context, action):
    context.user_data.clear()
    context.user_data["input"] = {"action": action, "values": []}
    await reply(update, INPUT_PROMPTS[action][0] + "\n/bekor — bekor qilish.")


@admin_only
async def continue_input(update, context, text):
    state = context.user_data["input"]
    state["values"].append(text)
    prompts = INPUT_PROMPTS[state["action"]]
    if len(state["values"]) < len(prompts):
        await reply(update, prompts[len(state["values"])])
        return
    context.user_data.pop("input", None)
    context.args = state["values"] if state["action"] == "xodim_qosh" else " ".join(state["values"]).split()
    await command_actions()[state["action"]](update, context)


@admin_only
async def export_employees(update, context):
    rows = context.bot_data["db"].employee_list()
    values = [[r["name"], r["phone"], "" if r["telegram_ref"].startswith("phone:") else r["telegram_ref"],
               r["shift"], str(r["telegram_id"] or ""), "tasdiqlangan" if r["approved"] else "tasdiq kerak", r['phone_extra']] for r in rows]
    workbook = await asyncio.to_thread(build_workbook, ["ism", "telefon", "telegram_id", "smena_vaqti", "ulangan_telegram_id", "ulanish_holati", "qoshimcha_telefon"], values, "Xodimlar")
    await update.effective_message.reply_document(document=workbook, filename="xodimlar.xlsx",
        caption="Tahrirlab, Xodim yuklash orqali qayta yuboring. Ulanish holati Excel orqali tasdiqlanmaydi.")


@admin_only
async def add_employee(update, context):
    if not context.args:
        await begin_input(update, context, "xodim_qosh")
        return
    if len(context.args) not in (4, 5):
        await reply(update, "/xodim_qosh orqali ism, telefon, Telegram ID (ixtiyoriy) va smenani kiriting.")
        return
    name, phone_raw, reference, shift = context.args[:4]
    extra_raw = context.args[4] if len(context.args) == 5 else "-"
    try:
        phone = phone_number(phone_raw)
        skip = extra_raw.casefold().replace('‘', "'").replace('’', "'") in ("yo'q", "yoq", "-", "o'tkazish", "otkazish")
        extra = None if skip else phone_number(extra_raw)
        if extra == phone:
            extra = None
        reference = "" if reference.casefold() in ("o‘tkazish", "o'tkazish", "otkazish", "-") else reference
        shift = "" if shift.casefold() in ("o‘tkazish", "o'tkazish", "otkazish", "-") else shift
        from excel_utils import shift_value
        row = {"name": name.strip(), "phone": phone, "phone_extra": extra,
               "telegram_ref": reference.lower() or "phone:" + phone,
               "shift": shift_value(shift), "row_number": 1}
        if not row["name"] or len(row["name"]) > 120:
            raise ValueError("Ism 1–120 belgi bo'lsin.")
        if reference and not (reference.isdigit() or reference.startswith("@")):
            raise ValueError("Telegram ID musbat raqam yoki @username bo'lsin.")
        context.bot_data["db"].import_employees([row])
        await reply(update, f"{name.strip()} xodimlar ro'yxatiga qo'shildi. Asosiy: {phone}. Qo'shimcha: {extra or 'yo‘q'}. Xodim /start, so'ng 'Xodim sifatida ulanish' orqali o'z telefonini tasdiqlasin.", reply_markup=ADMIN_MENU)
    except (ValueError, ValidationError) as exc:
        await reply(update, str(exc))


@admin_only
async def export_all(update, context):
    rows = context.bot_data["db"].export_appeals()
    workbook = await asyncio.to_thread(appeals_workbook, rows, context.bot_data["settings"].timezone)
    await update.effective_message.reply_document(document=workbook, filename="murojaatlar.xlsx",
        caption="Yuborish boshlanmagan qatorlarni tahrirlab, Murojaat yuklash orqali qayta yuboring. Asl ustunlar asl: bilan saqlangan.")


async def my_appeals(update, context):
    rows = context.bot_data["db"].export_appeals(telegram_id=update.effective_user.id)
    if not rows:
        await reply(update, "Sizga biriktirilgan va vaqti kelgan murojaat yo'q yoki ulanishingiz hali tasdiqlanmagan.")
        return
    text = "\n\n".join(f"Murojaat #{row['id']}\nMijoz: {row['client_name']}\n{row['body']}" for row in rows)
    await reply(update, text)


def notify_join(context, request_id):
    context.bot_data["service"].notice(f"join:{request_id}",
        "Xodim ulanishi uchun yangi so'rov. /menu → Xodimlar → Ulanish so‘rovlari orqali ko‘ring.")


async def join_employee(update, context):
    user = update.effective_user
    context.bot_data["db"].register_user(user.id, user.username)
    context.user_data.clear()
    context.user_data["join"] = True
    keyboard = ReplyKeyboardMarkup([[KeyboardButton("Telefonimni ulashish", request_contact=True)], ["Bekor qilish"]], resize_keyboard=True)
    await reply(update, "Xodimlar ro'yxatidagi o'z telefon raqamingizni tugma orqali ulashing. Keyin admin tasdiqlaydi.", reply_markup=keyboard)


async def join_contact(update, context):
    if not context.user_data.get("join"):
        await reply(update, "Avval 'Xodim sifatida ulanish' tugmasini bosing.")
        return
    contact = update.effective_message.contact
    if contact.user_id != update.effective_user.id:
        await reply(update, "Faqat o'zingizning kontaktingizni 'Telefonimni ulashish' tugmasi orqali yuboring.")
        return
    try:
        raw = contact.phone_number
        normalized = phone_number(raw)
        request_id = context.bot_data["db"].join_by_phone(normalized, update.effective_user.id)
        context.user_data.clear()
        if request_id:
            notify_join(context, request_id)
        await reply(update, "Ulanish so'rovi yuborildi. Admin tasdig'ini kuting." if request_id else "Ulanishingiz allaqachon tasdiqlangan.", reply_markup=USER_MENU)
    except ValueError as exc:
        await reply(update, str(exc))


@admin_only
async def pending_joins(update, context):
    rows = context.bot_data["db"].pending_joins()
    if not rows:
        await reply(update, "Kutilayotgan ulanish so'rovlari yo'q.")
    for row in rows:
        keyboard = InlineKeyboardMarkup([[
            InlineKeyboardButton("Tasdiqlash", callback_data=f"join:yes:{row['id']}"),
            InlineKeyboardButton("Rad etish", callback_data=f"join:no:{row['id']}"),
        ]])
        await reply(update, f"So'rov #{row['id']}\nXodim: {row['name']}\nTelefon: {row['phone']}\nQo'shimcha: {row['phone_extra'] or 'yo‘q'}\nTelegram ID: {row['telegram_id']}", reply_markup=keyboard)


def decide_request(context, request_id, approve, admin_id):
    telegram_id = context.bot_data["db"].decide_join(request_id, approve, admin_id)
    text = "Xodim sifatida ulanishingiz tasdiqlandi. /menu → Murojaatlarim orqali murojaatlarni matn ko‘rinishida o‘qing." if approve else "Ulanish so'rovingiz rad etildi. Admin bilan bog'laning."
    context.bot_data["db"].enqueue_notice(f"join-result:{request_id}", [telegram_id], text)


@admin_only
async def approve_join(update, context):
    await decide_command(update, context, True)


@admin_only
async def reject_join(update, context):
    await decide_command(update, context, False)


async def decide_command(update, context, approve):
    if not context.args:
        await begin_input(update, context, "xodim_tasdiq" if approve else "xodim_rad")
        return
    if len(context.args) != 1 or not context.args[0].isdigit():
        await reply(update, "So'rov raqamini kiriting. /tasdiqlar orqali ro'yxatni oling.")
        return
    try:
        decide_request(context, int(context.args[0]), approve, update.effective_user.id)
        await reply(update, "Xodim tasdiqlandi." if approve else "So'rov rad etildi.")
    except ValueError as exc:
        await reply(update, str(exc))


async def join_callback(update, context):
    query = update.callback_query
    if update.effective_chat.type != "private" or update.effective_user.id not in context.bot_data["settings"].admin_ids:
        await query.answer("Faqat admin tasdiqlashi mumkin.", show_alert=True)
        return
    _, decision, request_id = query.data.split(":")
    try:
        decide_request(context, int(request_id), decision == "yes", update.effective_user.id)
        await query.answer("Saqlandi")
        await query.edit_message_text("Xodim tasdiqlandi." if decision == "yes" else "So'rov rad etildi.")
    except ValueError as exc:
        await query.answer(str(exc), show_alert=True)


@admin_only
async def send_now(update, context):
    if not context.args:
        context.args = ["darhol"]
        await automatic_schedule(update, context)
        return
    try:
        if len(context.args) != 1:
            raise ValueError("Bitta murojaat ID sini kiriting.")
        db = context.bot_data["db"]
        result = db.details(context.args[0])
        if not result:
            raise ValueError("Murojaat topilmadi.")
        employee = next((r for r in db.employee_list() if r["id"] == result[0]["employee_id"]), None)
        if not employee or not employee["approved"]:
            raise ValueError("Avval xodimni biriktiring va ulanishini tasdiqlang.")
        db.set_time(context.args[0], time.time())
        await reply(update, "Yuborish navbatga qo'yildi. 15 soniya ichida Excel xodimga yuboriladi; /holat orqali tekshiring.")
    except ValueError as exc:
        await reply(update, str(exc))


@admin_only
async def automatic_schedule(update, context):
    db, settings = context.bot_data["db"], context.bot_data["settings"]
    if not context.args:
        context.user_data.clear()
        context.user_data["schedule"] = True
        current = db.default_schedule()
        keyboard = ReplyKeyboardMarkup([["Darhol", "09:00", "18:00"], ["Bekor qilish"]], resize_keyboard=True)
        await reply(update,
            f"Joriy avtomatik yuborish: {current or 'sozlanmagan'}.\n"
            "Barcha yangi murojaatlar uchun vaqtni bir marta tanlang yoki HH:MM yozing.\n"
            "Darhol — yuklangach bot o'zi yuboradi. Soat tanlansa, eng yaqin shu soatga rejalashtiriladi.\n"
            "Excelda alohida vaqt bo'lsa, o'sha vaqt ishlatiladi. Xodim ulanishi tasdiqlangan bo'lishi kerak.\n"
            f"Vaqt mintaqasi: {settings.timezone.key}", reply_markup=keyboard)
        return
    try:
        count = db.set_default_schedule(" ".join(context.args), settings.timezone)
        db.assign_due(time.time(), settings.timezone)
        context.user_data.pop("schedule", None)
        mode = db.default_schedule()
        description = "yuklangach darhol" if mode == "darhol" else f"eng yaqin {mode} da"
        await reply(update,
            f"Avtomatik yuborish yoqildi: {description}.\n{count} ta kutayotgan murojaat jadvalga qo'yildi.\n"
            "Endi Excel yuklasangiz, bot tasdiqlangan xodimlarga o'zi yuboradi. Har bir murojaat ID sini kiritish shart emas.\n"
            "Xodimi topilmagan yoki tasdiqlanmagan murojaatlar kutadi; /holat orqali ko'ring.", reply_markup=ADMIN_MENU)
    except ValueError as exc:
        await reply(update, str(exc))


BUTTON_ACTIONS = {
    "Bosh menyu": start, "Bugungi yuborishlar": today, "Holat": status,
    "Murojaat yuklash": upload_appeals, "Xodim yuklash": upload_employees,
    "Murojaatlar": all_appeals, "Xodimlar": employees,
    "Murojaatlar Excel": export_all, "Xodimlar Excel": export_employees,
    "Xodim qo‘shish": add_employee,
    "Vaqt belgilash": set_time, "Xodim biriktirish": assign,
    "Hozir yuborish": send_now, "Batafsil": details,
    "Avtomatik yuborish": automatic_schedule,
    "Ulanish so'rovlari": pending_joins, "Xodimni tasdiqlash": approve_join,
    "Xodimni rad etish": reject_join, "Qayta yuborish": retry,
    "Yuborishni tasdiqlash": confirm, "Shablonlar": templates, "Yordam": help_command,
    "Murojaat yuborish": new_appeal, "Mening ID raqamim": identity, "Bekor qilish": cancel,
    "Xodim sifatida ulanish": join_employee, "Mening murojaatlarim": my_appeals,
}
