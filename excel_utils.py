"""Validate Excel input atomically and report row numbers."""
import hashlib
import json
import re
from datetime import datetime
from io import BytesIO
from zipfile import ZipFile
from zoneinfo import ZoneInfo

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from config import TIMEZONE

MAX_FILE_SIZE = 5 * 1024 * 1024
LEGACY_DEADLINE_WARNING = ("⚠️ Муддати — hujjat ijro muddati; yuborish vaqti sifatida olinmadi. "
                           "Yuborish bo'limidan vaqt tanlang yoki yuborish_vaqti ustunini to'ldiring.")
PORTAL_DEADLINE_HEADERS = ("муддат", "хозирги муддати")
PORTAL_DEADLINE_WARNING = ("⚠️ Муддат va Хозирги муддати — portal ijro muddatlari. "
                           "Xabar yuborish vaqtini Yuborish bo'limidan tanlang.")
APPEAL_UPLOAD_PROMPT = ("Murojaatlar uchun .xlsx fayl yuboring. Standart shablon, eski murojaatlar jadvali, "
                        "IIV portallari va PP100 formatlari qabul qilinadi. Bot birinchi varaqni o'qiydi. "
                        "Telefon orqali biriktirish uchun xodim_telefoni ustunini to'ldiring.")
EMPLOYEE_PHONE_REFERENCE = "phone:"
PHONE_EXPORT_FILENAME = re.compile(r"murojaatlar_bot_murojatlar asil_[0-9]{8}_[0-9]{6}\.xlsx", re.IGNORECASE)
PHONE_EXPORT_HEADERS = {"id", "mijoz_ismi", "telefon", "murojaat_matni", "yuborish_vaqti", "xodim"}


def import_notes(rows):
    notes = ["Bir xil ID va bir xil ma'lumot takroran yuborilmaydi. Tahrir faqat yuborish boshlanmaguncha mumkin; "
             "bo'sh vaqt/xodim oldingi qiymatni saqlaydi. Yangi murojaatga yangi ID bering."]
    duplicates = sum(row.get("pp100_duplicates", 0) for row in rows)
    if duplicates:
        notes.append(f"PP100: {duplicates} ta takroriy qator o'tkazib yuborildi. Har bir raqamning birinchi qatori olindi.")
    if any((row.get("employee_ref") or "").startswith(EMPLOYEE_PHONE_REFERENCE) for row in rows):
        notes.append("Telefon bo‘yicha moslik topilmagan yoki bir nechta xodim mos kelgan murojaatlar Holat bo‘limida ko‘rinadi.")
    if any('муддати' in row.get('source_data', {}) for row in rows):
        notes.append(LEGACY_DEADLINE_WARNING)
    if any(row.get("source_data", {}).keys() & PORTAL_DEADLINE_HEADERS for row in rows):
        notes.append(PORTAL_DEADLINE_WARNING)
    return "\n".join(notes)
HEADERS = {
    "appeals": ("id", "mijoz_ismi", "murojaat_matni", "yuborish_vaqti"),
    "employees": ("ism", "telefon", "telegram_id"),
}
LEGACY_HEADERS = {
    "кириш рақами": "id", "аризачи": "mijoz_ismi",
    "ҳужжат мазмуни": "murojaat_matni", "кимга юборилган": "xodim",
}
LEGACY_SENDER_HEADER = "юборган ижрочи"
PORTAL_HEADERS = {
    "ариза рақами": "id", "мурожаат рақами": "id",
    "мурожаат муаллифининг ф.и.ш.": "mijoz_ismi",
    "фамиля": "surname", "фамилия": "surname", "исми": "given_name",
    "мурожат таснифи": "murojaat_matni", "мурожаат таснифи": "murojaat_matni",
    "ижрочи": "xodim",
}
PORTAL_APPEAL_FIELDS = {"id", "mijoz_ismi", "surname", "given_name", "murojaat_matni"}
PORTAL_IMPORT_FIELDS = PORTAL_APPEAL_FIELDS | {"xodim", "xodim_telefoni", "telefon", "yuborish_vaqti"}


class ValidationError(ValueError):
    def __init__(self, errors):
        self.errors = errors
        super().__init__("\n".join(errors))


def clean(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalized_header(value):
    return " ".join(clean(value).casefold().split())


def appeal_contact(row):
    name = clean(row["client_name"])
    if name and name.casefold().replace("‘", "'").replace("’", "'") != "ko'rsatilmagan":
        return "Mijoz", name
    source = json.loads(row.get("source_json") or "{}")
    sender = clean(source.get(LEGACY_SENDER_HEADER))
    if sender:
        return "Юборган ижрочи", sender
    return "Mijoz", ""


def phone_number(value):
    value = re.sub(r"[\s()\-]", "", clean(value))
    if re.fullmatch(r"\d{9}", value):
        value = "+998" + value
    elif re.fullmatch(r"998\d{9}", value):
        value = "+" + value
    if not re.fullmatch(r"\+998[0-9]{9}", value):
        raise ValueError("telefon +998XXXXXXXXX formatida bo'lsin, masalan +998901234567")
    return value


def parse_datetime(value, timezone):
    timezone = TIMEZONE
    if isinstance(value, datetime):
        result = value
    else:
        result = None
        for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%d.%m.%Y %H:%M"):
            try:
                result = datetime.strptime(clean(value), fmt)
                break
            except ValueError:
                pass
        if result is None:
            raise ValueError("vaqt YYYY-MM-DD HH:MM yoki DD.MM.YYYY HH:MM formatida bo'lsin")
    return result.replace(tzinfo=timezone) if result.tzinfo is None else result.astimezone(timezone)


def shift_value(value):
    value = clean(value)
    if value and not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d-(?:[01]\d|2[0-3]):[0-5]\d", value):
        raise ValueError("smena_vaqti HH:MM-HH:MM formatida bo'lsin")
    return value


def in_shift(shift, moment):
    if not shift:
        return True
    start, end = shift.split("-")
    current = moment.strftime("%H:%M")
    if start == end:
        return True
    return start <= current < end if start < end else current >= start or current < end


def find_header(sheet, kind):
    for row_number, cells in enumerate(sheet.iter_rows(max_row=20, max_col=60), 1):
        source = [normalized_header(cell.value) for cell in cells]
        names = set(source)
        if kind == "employees" or (row_number == 1 and "id" in names):
            return row_number, source, source, ""
        if (set(LEGACY_HEADERS) - {"аризачи"}).issubset(names) and names & {"аризачи", LEGACY_SENDER_HEADER}:
            return row_number, source, [LEGACY_HEADERS.get(name, name) for name in source], "L"
        mapped = [PORTAL_HEADERS.get(name, name) for name in source]
        mapped_names = set(mapped)
        if names & {"ариза рақами", "мурожаат рақами"} and {"id", "murojaat_matni"}.issubset(mapped_names):
            if "mijoz_ismi" in mapped_names:
                return row_number, source, mapped, "IIV"
            if {"surname", "given_name"}.issubset(mapped_names):
                return row_number, source, mapped, "PP100"
    return None


def import_rows(workbook, kind, filename):
    sheet = next(iter(workbook.worksheets), None) if kind == "appeals" else workbook.active
    match = find_header(sheet, kind) if sheet is not None else None
    if match is None:
        required = ", ".join(HEADERS[kind])
        raise ValidationError([f"1-qator: shablon sarlavhasi topilmadi. Standart ustunlar — {required}"])

    header_row, source, header, layout = match
    portal = layout in ("IIV", "PP100")
    sheet_label = "1-varaq, " if portal else ""
    if sheet.max_row is None or sheet.max_column is None:
        # Some valid XLSX exporters omit the optional worksheet dimension.
        sheet.calculate_dimension(force=True)
    if sheet.max_row > 5020 or sheet.max_column > 60:
        raise ValidationError([f"{sheet_label}Ko'pi bilan 5000 qator va 60 ustun yuklang."])
    location = f"{sheet_label}{header_row}-qator"
    required = {"id", "murojaat_matni"} if layout else set(HEADERS[kind]) - {"telegram_id"}
    missing = sorted(required - set(header))
    if missing:
        raise ValidationError([f"{location}: ustunlar yetishmayapti: " + ", ".join(missing)])
    names = [name for name in header if name]
    if len(names) != len(set(names)):
        raise ValidationError([f"{location}: takroriy ustun nomi mavjud."])
    phone_export = (kind == "appeals" and not layout and set(names) == PHONE_EXPORT_HEADERS
                    and PHONE_EXPORT_FILENAME.fullmatch(filename) is not None)
    if phone_export:
        header = ["xodim_telefoni" if name == "telefon" else name for name in header]
    rows = sheet.iter_rows(min_row=header_row + 1)
    for row_number, cells in enumerate(rows, header_row + 1):
        if portal:
            populated = any(clean(cell.value) for name, cell in zip(header, cells) if name in PORTAL_APPEAL_FIELDS)
        else:
            populated = any(cell.value is not None for cell in cells)
        if populated:
            yield row_number, f"{sheet_label}{row_number}-qator", cells, source, header, layout, phone_export


def parse_excel(payload: bytes, kind: str, timezone: ZoneInfo, *, filename: str = "") -> list[dict]:
    timezone = TIMEZONE
    if len(payload) > MAX_FILE_SIZE:
        raise ValidationError(["Fayl 5 MB dan oshmasin."])
    try:
        with ZipFile(BytesIO(payload)) as archive:
            if sum(item.file_size for item in archive.infolist()) > 40 * 1024 * 1024:
                raise ValidationError(["Excel ochilgandagi hajmi juda katta."])
        workbook = load_workbook(BytesIO(payload), read_only=True, data_only=False)
    except ValidationError:
        raise
    except Exception:
        raise ValidationError(["Fayl buzilgan yoki haqiqiy .xlsx emas."]) from None
    errors, output, seen = [], [], {}
    try:
        rows = import_rows(workbook, kind, filename)
        for count, (row_number, location, cells, source_headers, header, layout, phone_export) in enumerate(rows, 1):
            if count > 5000:
                raise ValidationError(["Bir faylda jami 5000 tagacha ma'lumot qatori yuklang."])
            try:
                if any(cell.data_type == "f" for name, source, cell in zip(header, source_headers, cells)
                       if layout not in ("IIV", "PP100") or name in PORTAL_IMPORT_FIELDS
                       or source in PORTAL_DEADLINE_HEADERS):
                    raise ValueError("formula o'rniga oddiy qiymat kiriting")
                row = {key: cell.value for key, cell in zip(header, cells) if key}
                if layout:
                    source_data = {key: clean(cell.value) for key, cell in zip(source_headers, cells)
                                   if key and (layout == "L" or key in PORTAL_DEADLINE_HEADERS)}
                else:
                    source_data = {key.removeprefix("asl:"): clean(value)
                                   for key, value in row.items() if key.startswith("asl:")}
                if kind == "employees":
                    name = clean(row["ism"])
                    if not name or len(name) > 120:
                        raise ValueError("ism majburiy va 120 belgidan oshmasin")
                    contact = clean(row.get("telegram_id"))
                    if contact and not (re.fullmatch(r"[1-9]\d{0,18}", contact) or re.fullmatch(r"@[A-Za-z][A-Za-z0-9_]{4,31}", contact)):
                        raise ValueError("telegram_id musbat ID yoki @username bo'lsin")
                    phone = phone_number(row["telefon"])
                    extra = phone_number(row["qoshimcha_telefon"]) if clean(row.get("qoshimcha_telefon")) else None
                    if extra == phone:
                        extra = None
                    item = dict(name=name, phone=phone, phone_extra=extra,
                                telegram_ref=contact.lower() or "phone:" + phone, shift=shift_value(row.get("smena_vaqti")))
                    key = name.casefold()
                    if contact and any(other["telegram_ref"] == item["telegram_ref"] for other in output):
                        raise ValueError("Telegram identifikatori faylda takrorlangan")
                else:
                    appeal_id = clean(row["id"])
                    name, body = clean(row.get("mijoz_ismi")), clean(row["murojaat_matni"])
                    employee_ref = clean(row.get("xodim"))
                    employee_phone = clean(row.get("xodim_telefoni"))
                    if phone_export and not employee_phone:
                        raise ValueError("telefon ustuniga xodim raqamini kiriting")
                    if layout:
                        if not appeal_id:
                            raise ValueError("Murojaat raqami bo'sh" if layout != "L" else "Kirish raqami bo'sh")
                        number_label = "Kirish raqami" if layout == "L" else "Murojaat raqami"
                        body = body or "Hujjat mazmuni ko'rsatilmagan."
                        body = f"{number_label}: {appeal_id}\n{body}"
                        appeal_id = f"{layout}-" + hashlib.sha256(appeal_id.encode()).hexdigest()[:16]
                        if layout == "PP100":
                            surname, given_name = clean(row.get("surname")), clean(row.get("given_name"))
                            name = " ".join(filter(None, (surname, given_name)))
                            if not employee_ref and not employee_phone and surname and given_name:
                                employee_ref = name
                        name = name or "Ko'rsatilmagan"
                    if not re.fullmatch(r"[\w./-]{1,64}", appeal_id):
                        raise ValueError("id 1–64 belgidan iborat bo'lsin. Harf, raqam, _, -, . va / ishlatiladi")
                    if not name or len(name) > 32767:
                        raise ValueError("mijoz_ismi majburiy va Excel katagi chegarasidan (32767) oshmasin")
                    if not body or len(body) > 32767:
                        raise ValueError("murojaat_matni majburiy va Excel katagi chegarasidan (32767) oshmasin")
                    when = row.get("yuborish_vaqti")
                    legacy_row = bool(layout) or clean(row.get("manba_formati")) == "eski_jadval"
                    phone_raw = clean(row.get("telefon"))
                    phone = "Ko'rsatilmagan" if phone_raw in ("", "Ko'rsatilmagan") else phone_number(phone_raw)
                    if employee_phone:
                        try:
                            employee_phone = phone_number(employee_phone)
                        except ValueError as exc:
                            column = "telefon" if phone_export else "xodim_telefoni"
                            raise ValueError(f"{column}: {exc}") from None
                        qualifier = "|" + employee_ref if employee_ref else ""
                        employee_ref = EMPLOYEE_PHONE_REFERENCE + employee_phone + qualifier
                    elif employee_ref.startswith("+"):
                        try:
                            employee_ref = EMPLOYEE_PHONE_REFERENCE + phone_number(employee_ref)
                        except ValueError as exc:
                            raise ValueError(f"xodim: {exc}") from None
                    item = dict(id=appeal_id, client_name=name, phone=phone,
                                body=body, due_at=parse_datetime(when, timezone).timestamp() if clean(when) else None,
                                employee_ref=employee_ref, legacy=legacy_row, source_data=source_data)
                    key = appeal_id
                if key in seen:
                    if layout == "PP100":
                        first = seen[key]
                        first["pp100_duplicates"] = first.get("pp100_duplicates", 0) + 1
                        continue
                    raise ValueError(f"ID '{key}' faylda takrorlangan; har bir ID faqat bitta qatorda bo'lsin" if kind == 'appeals'
                                     else "ism faylda takrorlangan")
                seen[key] = item
                item["row_number"] = row_number
                if layout in ("IIV", "PP100"):
                    item["row_location"] = location
                output.append(item)
            except (ValueError, TypeError) as exc:
                errors.append(f"{location}: {exc}.")
                if len(errors) >= 50:
                    errors.append("Dastlabki 50 xato ko'rsatildi; tuzatib qayta yuklang.")
                    break
        if errors:
            raise ValidationError(errors)
        if not output:
            raise ValidationError(["Faylda ma'lumot qatorlari yo'q."])
        return output
    finally:
        workbook.close()


def build_workbook(columns, rows, title="Murojaatlar"):
    """Produce a download in memory; strings remain literal, never formulas."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = title
    sheet.append(columns)
    for values in rows:
        sheet.append(values)
        for cell in sheet[sheet.max_row]:
            if isinstance(cell.value, str):
                cell.data_type = "s"
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if isinstance(cell.value, datetime):
                cell.number_format = "yyyy-mm-dd hh:mm"
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="245A81")
        sheet.column_dimensions[cell.column_letter].width = 35
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    stream = BytesIO()
    workbook.save(stream)
    stream.seek(0)
    return stream


def appeals_workbook(rows, timezone):
    timezone = TIMEZONE
    sources = [json.loads(row.get("source_json") or "{}") for row in rows]
    keys = list(dict.fromkeys(key for source in sources for key in source))
    columns = ["id", "mijoz_ismi", "telefon", "murojaat_matni", "yuborish_vaqti", "xodim", "manba_formati", "xodim_telefoni"]
    values = []
    for row, source in zip(rows, sources):
        when = datetime.fromtimestamp(row["due_at"], timezone).replace(tzinfo=None) if row["due_at"] else None
        hint = row.get("employee_hint") or ""
        employee_phone = ""
        employee = row.get("employee") or hint
        if hint.startswith(EMPLOYEE_PHONE_REFERENCE):
            employee_phone, _, employee = hint.removeprefix(EMPLOYEE_PHONE_REFERENCE).partition("|")
        values.append([row["id"], row["client_name"], row["phone"], row["body"], when,
                       employee, "eski_jadval" if row.get("legacy") else "standart", employee_phone] +
                      [source.get(key, "") for key in keys])
    return build_workbook(columns + ["asl:" + key for key in keys], values)
