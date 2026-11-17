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

MAX_FILE_SIZE = 5 * 1024 * 1024
HEADERS = {
    "appeals": ("id", "mijoz_ismi", "telefon", "murojaat_matni", "yuborish_vaqti"),
    "employees": ("ism", "telefon", "telegram_id"),
}
LEGACY_HEADERS = {
    "кириш рақами": "id", "аризачи": "mijoz_ismi",
    "ҳужжат мазмуни": "murojaat_matni", "кимга юборилган": "xodim",
}


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


def phone_number(value):
    value = re.sub(r"[\s()\-]", "", clean(value))
    if re.fullmatch(r"\d{9}", value):
        value = "+998" + value
    elif re.fullmatch(r"998\d{9}", value):
        value = "+" + value
    if not re.fullmatch(r"\+998\d{9}", value):
        raise ValueError("telefon +998XXXXXXXXX formatida bo'lsin, masalan +998901234567")
    return value


def parse_datetime(value, timezone):
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


def parse_excel(payload: bytes, kind: str, timezone: ZoneInfo) -> list[dict]:
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
    errors, output, seen = [], [], set()
    try:
        sheet = workbook.active
        if sheet.max_row is None or sheet.max_column is None:
            # Some valid XLSX exporters omit the optional worksheet dimension.
            sheet.calculate_dimension(force=True)
        if sheet.max_row > 5020 or sheet.max_column > 60:
            raise ValidationError(["Ko'pi bilan 5000 qator va 60 ustun yuklang."])
        iterator = sheet.iter_rows()
        header = [normalized_header(cell.value) for cell in next(iterator, [])]
        header_row, legacy = 1, False
        if kind == "appeals" and "id" not in header:
            while True:
                if set(LEGACY_HEADERS).issubset(header):
                    legacy = True
                    break
                if header_row >= 20:
                    break
                cells = next(iterator, None)
                if cells is None:
                    break
                header = [normalized_header(cell.value) for cell in cells]
                header_row += 1
        source_headers = list(header)
        if legacy:
            header = [LEGACY_HEADERS.get(name, name) for name in header]
        missing = sorted(set(HEADERS[kind]) - set(header))
        if legacy:
            missing = [name for name in missing if name not in ("telefon", "yuborish_vaqti")]
        if kind == "employees":
            missing = [name for name in missing if name != "telegram_id"]
        if missing:
            raise ValidationError([f"{header_row}-qator: ustunlar yetishmayapti: " + ", ".join(missing)])
        names = [name for name in header if name]
        if len(names) != len(set(names)):
            raise ValidationError(["1-qator: takroriy ustun nomi mavjud."])
        for row_number, cells in enumerate(iterator, header_row + 1):
            if all(cell.value is None for cell in cells):
                continue
            try:
                if any(cell.data_type == "f" for cell in cells):
                    raise ValueError("formula o'rniga oddiy qiymat kiriting")
                row = {key: cell.value for key, cell in zip(header, cells) if key}
                source_data = {key: clean(cell.value) for key, cell in zip(source_headers, cells) if key} if legacy else {
                    key.removeprefix("asl:"): clean(value) for key, value in row.items() if key.startswith("asl:")}
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
                    for other in output:
                        if ({phone, extra} - {None}) & ({other["phone"], other.get("phone_extra")} - {None}):
                            raise ValueError(f"asosiy yoki qo'shimcha telefon {other['row_number']}-qatordagi boshqa xodimda takrorlangan")
                        if other["telegram_ref"] == item["telegram_ref"]:
                            raise ValueError("Telegram identifikatori faylda takrorlangan")
                else:
                    appeal_id = clean(row["id"])
                    name, body = clean(row["mijoz_ismi"]), clean(row["murojaat_matni"])
                    if legacy:
                        if not appeal_id:
                            raise ValueError("Kirish raqami bo'sh")
                        original_id = appeal_id
                        appeal_id = "L-" + hashlib.sha256(original_id.encode()).hexdigest()[:16]
                        body = f"Kirish raqami: {original_id}\n{body}" if body else ""
                        name = name or "Ko'rsatilmagan"
                        body = body or f"Kirish raqami: {original_id}\nHujjat mazmuni ko'rsatilmagan."
                    if not re.fullmatch(r"[\w.-]{1,64}", appeal_id):
                        raise ValueError("id 1–64 ta harf, raqam, nuqta yoki chiziqdan iborat bo'lsin")
                    if not name or len(name) > 32767:
                        raise ValueError("mijoz_ismi majburiy va Excel katagi chegarasidan (32767) oshmasin")
                    if not body or len(body) > 32767:
                        raise ValueError("murojaat_matni majburiy va Excel katagi chegarasidan (32767) oshmasin")
                    when = row.get("yuborish_vaqti")
                    legacy_row = legacy or clean(row.get("manba_formati")) == "eski_jadval"
                    phone_raw = clean(row.get("telefon"))
                    phone = "Ko'rsatilmagan" if legacy_row and phone_raw in ("", "Ko'rsatilmagan") else phone_number(phone_raw)
                    item = dict(id=appeal_id, client_name=name, phone=phone,
                                body=body, due_at=parse_datetime(when, timezone).timestamp() if clean(when) else None,
                                employee_ref=clean(row.get("xodim")), legacy=legacy_row, source_data=source_data)
                    key = appeal_id
                if key in seen:
                    raise ValueError("id yoki ism faylda takrorlangan")
                seen.add(key)
                item["row_number"] = row_number
                output.append(item)
            except (ValueError, TypeError) as exc:
                errors.append(f"{row_number}-qator: {exc}.")
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
    sources = [json.loads(row.get("source_json") or "{}") for row in rows]
    keys = list(dict.fromkeys(key for source in sources for key in source))
    columns = ["id", "mijoz_ismi", "telefon", "murojaat_matni", "yuborish_vaqti", "xodim", "manba_formati"]
    values = []
    for row, source in zip(rows, sources):
        when = datetime.fromtimestamp(row["due_at"], timezone).replace(tzinfo=None) if row["due_at"] else None
        values.append([row["id"], row["client_name"], row["phone"], row["body"], when,
                       row.get("employee") or row.get("employee_hint", ""), "eski_jadval" if row.get("legacy") else "standart"] +
                      [source.get(key, "") for key in keys])
    return build_workbook(columns + ["asl:" + key for key in keys], values)
