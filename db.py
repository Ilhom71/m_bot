"""SQLite outbox and atomic state transitions."""
import hashlib
import hmac
import json
import os
import re
import sqlite3
import time
from contextlib import contextmanager
from datetime import datetime, timedelta
from pathlib import Path
from contextlib import closing
from unicodedata import normalize

from excel_utils import EMPLOYEE_PHONE_REFERENCE, ValidationError, in_shift, phone_number
from config import TIMEZONE

STATUSES = {
    "pending": "kutmoqda", "retry": "qayta urinish", "sending": "yuborilmoqda",
    "sent": "yuborildi", "failed": "xato", "unknown": "natija noma'lum",
    "blocked": "SMS API kutilmoqda",
}
SMS_PASSWORD_ITERATIONS = 600_000
SMS_PASSWORD_ATTEMPTS = 5
SMS_PASSWORD_WINDOW = 300


def employee_name_key(name):
    return "".join(normalize("NFC", name).casefold().split())


def assignment_problem(row):
    if row.get('completed_at') is not None:
        return None
    if not row.get('employee_id'):
        return 'unassigned'
    if not row.get('employee_approved') or not row.get('employee_telegram_id'):
        return 'unconnected'
    return None


class Database:
    def __init__(self, path):
        self.path = str(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with self.connect() as conn:
            conn.executescript("""
                PRAGMA journal_mode=WAL;
                CREATE TABLE IF NOT EXISTS users (
                    telegram_id INTEGER PRIMARY KEY, username TEXT, started_at REAL NOT NULL
                );
                CREATE TABLE IF NOT EXISTS employees (
                    id INTEGER PRIMARY KEY, name TEXT NOT NULL, name_key TEXT UNIQUE NOT NULL,
                    phone TEXT NOT NULL, telegram_ref TEXT NOT NULL,
                    telegram_id INTEGER, shift TEXT NOT NULL DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS appeals (
                    id TEXT PRIMARY KEY, client_name TEXT NOT NULL, phone TEXT NOT NULL,
                    body TEXT NOT NULL, due_at REAL, employee_id INTEGER REFERENCES employees(id),
                    created_at REAL NOT NULL, source_user INTEGER
                );
                CREATE TABLE IF NOT EXISTS deliveries (
                    appeal_id TEXT REFERENCES appeals(id), channel TEXT NOT NULL,
                    status TEXT NOT NULL DEFAULT 'pending', attempts INTEGER NOT NULL DEFAULT 0,
                    next_attempt REAL NOT NULL DEFAULT 0, last_error TEXT,
                    message_id TEXT, updated_at REAL NOT NULL, part_index INTEGER NOT NULL DEFAULT 0,
                    PRIMARY KEY (appeal_id, channel)
                );
                CREATE TABLE IF NOT EXISTS notices (
                    notice_key TEXT PRIMARY KEY, admin_id INTEGER, text TEXT NOT NULL,
                    status TEXT NOT NULL DEFAULT 'pending', attempts INTEGER NOT NULL DEFAULT 0,
                    next_attempt REAL NOT NULL DEFAULT 0
                );
                CREATE INDEX IF NOT EXISTS due_appeals ON appeals(due_at);
                CREATE INDEX IF NOT EXISTS ready_deliveries ON deliveries(status, next_attempt);
                CREATE TABLE IF NOT EXISTS join_requests (
                    id INTEGER PRIMARY KEY, employee_id INTEGER NOT NULL REFERENCES employees(id),
                    telegram_id INTEGER NOT NULL, status TEXT NOT NULL DEFAULT 'pending',
                    created_at REAL NOT NULL, decided_by INTEGER, decided_at REAL
                );
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY, value TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS sms_deliveries (
                    appeal_id TEXT NOT NULL REFERENCES appeals(id), phone TEXT NOT NULL,
                    status TEXT NOT NULL DEFAULT 'pending', attempts INTEGER NOT NULL DEFAULT 0,
                    next_attempt REAL NOT NULL DEFAULT 0, last_error TEXT, message_id TEXT,
                    updated_at REAL NOT NULL, PRIMARY KEY(appeal_id,phone)
                );
            """)
            # Rebuilding a parent table requires disabling foreign key enforcement before the transaction.
            conn.execute("PRAGMA foreign_keys=OFF")
            conn.execute("BEGIN IMMEDIATE")
            # Additive migration: existing contacts also require explicit approval.
            for table, columns in {
                "employees": {"approved": "INTEGER NOT NULL DEFAULT 0", "phone_extra": "TEXT"},
                "deliveries": {"part_index": "INTEGER NOT NULL DEFAULT 0"},
                "appeals": {"employee_hint": "TEXT NOT NULL DEFAULT ''", "source_json": "TEXT NOT NULL DEFAULT '{}'",
                            "legacy": "INTEGER NOT NULL DEFAULT 0", "schedule_source": "TEXT NOT NULL DEFAULT 'explicit'",
                            "completed_at": "REAL", "completed_by": "INTEGER",
                            "sms_held": "INTEGER NOT NULL DEFAULT 0"},
                "notices": {"updated_at": "REAL NOT NULL DEFAULT 0"},
            }.items():
                present = {row[1] for row in conn.execute(f"PRAGMA table_info({table})")}
                for name, definition in columns.items():
                    if name not in present:
                        conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {definition}")
                        if table == 'appeals' and name == 'sms_held':
                            # Protect the old backlog if SMS had already been explicitly enabled.
                            self._hold_overdue_sms(conn, previous_setting='1')
            self._allow_shared_employee_phones(conn)

    @staticmethod
    def _allow_shared_employee_phones(conn):
        old_constraints = conn.execute("""
            SELECT 1 FROM pragma_index_list('employees') AS indexes
            JOIN pragma_index_info(indexes.name) AS columns
            WHERE indexes.origin='u' AND columns.name IN ('phone', 'telegram_ref')
        """).fetchone()
        if old_constraints:
            conn.execute("""CREATE TABLE employees_new (
                id INTEGER PRIMARY KEY, name TEXT NOT NULL, name_key TEXT UNIQUE NOT NULL,
                phone TEXT NOT NULL, telegram_ref TEXT NOT NULL, telegram_id INTEGER,
                shift TEXT NOT NULL DEFAULT '', approved INTEGER NOT NULL DEFAULT 0, phone_extra TEXT
            )""")
            conn.execute("""INSERT INTO employees_new (id,name,name_key,phone,telegram_ref,telegram_id,shift,approved,phone_extra)
                SELECT id,name,name_key,phone,telegram_ref,telegram_id,shift,approved,phone_extra FROM employees""")
            conn.execute("DROP TABLE employees")
            conn.execute("ALTER TABLE employees_new RENAME TO employees")
            if conn.execute("PRAGMA foreign_key_check").fetchone():
                raise sqlite3.IntegrityError("Foreign key check failed during employee migration")
        conn.execute("""CREATE UNIQUE INDEX IF NOT EXISTS unique_employee_telegram_ref
            ON employees(telegram_ref) WHERE telegram_ref NOT GLOB 'phone:*'""")

    @contextmanager
    def connect(self):
        conn = sqlite3.connect(self.path, timeout=15)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys=ON")
        conn.create_function("employee_name_key", 1, employee_name_key, deterministic=True)
        try:
            with conn:
                yield conn
        finally:
            conn.close()

    def register_user(self, telegram_id, username):
        ref = "@" + username.lower() if username else None
        with self.connect() as conn:
            # A username can change ownership. Keep previously verified employee IDs pinned.
            conn.execute("UPDATE users SET username=NULL WHERE username=? AND telegram_id<>?", (ref, telegram_id))
            conn.execute("INSERT INTO users VALUES (?,?,?) ON CONFLICT(telegram_id) DO UPDATE SET username=excluded.username",
                         (telegram_id, ref, time.time()))
        # Knowing a username or numeric ID is never proof of employment.
        # A real Telegram contact share must match the Excel phone first.
        return []

    def request_join(self, employee_id, telegram_id):
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            return self._request_join(conn, employee_id, telegram_id)

    @staticmethod
    def _request_join(conn, employee_id, telegram_id):
        if conn.execute("SELECT 1 FROM employees WHERE telegram_id=? AND approved=1 AND id<>?",
                        (telegram_id, employee_id)).fetchone():
            raise ValueError("Bu Telegram hisobi boshqa xodimga ulangan.")
        employee = conn.execute("SELECT approved,telegram_id FROM employees WHERE id=?", (employee_id,)).fetchone()
        if not employee:
            raise ValueError("Xodim topilmadi. Admin avval xodimlar Excel faylini yuklasin.")
        if employee["approved"]:
            if employee["telegram_id"] == telegram_id:
                return None
            raise ValueError("Bu xodim boshqa hisobga tasdiqlangan. Admin bilan bog'laning.")
        request = conn.execute("SELECT id,status FROM join_requests WHERE employee_id=? AND telegram_id=? ORDER BY id DESC LIMIT 1",
                               (employee_id, telegram_id)).fetchone()
        if request and request["status"] == "pending":
            return request["id"]
        if request and request["status"] == "rejected":
            raise ValueError("Oldingi so'rov ko'rib chiqilgan. Admin bilan bog'laning.")
        return conn.execute("INSERT INTO join_requests(employee_id,telegram_id,created_at) VALUES (?,?,?)",
                            (employee_id, telegram_id, time.time())).lastrowid

    def join_by_phone(self, phone, telegram_id, username=None):
        phone = phone_number(phone)
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            rows = conn.execute("""SELECT id,telegram_ref,telegram_id,approved FROM employees
                WHERE phone=? OR phone_extra=?""", (phone, phone)).fetchall()
            if not rows:
                raise ValueError("Bu telefon xodimlar ro'yxatida topilmadi. Admin ro'yxatni tekshirsin.")
            if len(rows) > 1:
                linked = [row for row in rows if row["approved"] and row["telegram_id"] == telegram_id]
                refs = {str(telegram_id)}
                if username:
                    refs.add("@" + username.lower())
                rows = linked or [row for row in rows if not row["approved"] and row["telegram_ref"] in refs]
                if len(rows) != 1:
                    raise ValueError("Bu telefon bir nechta xodimda bor. Admin ro‘yxatga Telegram ID yoki @username kiritsin.")
            return self._request_join(conn, rows[0]["id"], telegram_id)

    def pending_joins(self):
        with self.connect() as conn:
            return [dict(row) for row in conn.execute("""SELECT r.id,r.telegram_id,e.name,e.phone,e.phone_extra
                FROM join_requests r JOIN employees e ON e.id=r.employee_id WHERE r.status='pending' ORDER BY r.id""")]

    def decide_join(self, request_id, approve, admin_id):
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            request = conn.execute("SELECT * FROM join_requests WHERE id=? AND status='pending'", (request_id,)).fetchone()
            if not request:
                raise ValueError("So'rov topilmadi yoki allaqachon ko'rib chiqilgan.")
            if approve:
                conflict = conn.execute("SELECT id FROM employees WHERE telegram_id=? AND id<>?",
                                        (request["telegram_id"], request["employee_id"])).fetchone()
                if conflict:
                    raise ValueError("Bu Telegram hisobi boshqa xodimga tasdiqlangan.")
                current = conn.execute("SELECT approved,telegram_id FROM employees WHERE id=?", (request["employee_id"],)).fetchone()
                if current["approved"] and current["telegram_id"] != request["telegram_id"]:
                    raise ValueError("Xodim boshqa hisobga tasdiqlangan.")
                conn.execute("UPDATE employees SET telegram_id=?,approved=1 WHERE id=?", (request["telegram_id"], request["employee_id"]))
                conn.execute("UPDATE join_requests SET status='rejected',decided_by=?,decided_at=? WHERE employee_id=? AND id<>? AND status='pending'",
                             (admin_id, time.time(), request["employee_id"], request_id))
                conn.execute("UPDATE deliveries SET next_attempt=0 WHERE status='retry' AND attempts=0 AND appeal_id IN (SELECT id FROM appeals WHERE employee_id=?)",
                             (request["employee_id"],))
            conn.execute("UPDATE join_requests SET status=?,decided_by=?,decided_at=? WHERE id=?",
                         ("approved" if approve else "rejected", admin_id, time.time(), request_id))
            return request["telegram_id"]

    def import_employees(self, rows):
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            for row in rows:
                row = dict(row)
                try:
                    row['phone'] = phone_number(row['phone'])
                    row['phone_extra'] = phone_number(row['phone_extra']) if row.get('phone_extra') else None
                except ValueError as exc:
                    raise ValidationError([f"{row['row_number']}-qator: {exc}"]) from None
                if row['phone_extra'] == row['phone']:
                    row['phone_extra'] = None
                ref = row["telegram_ref"]
                user = conn.execute("SELECT telegram_id FROM users WHERE username=?", (ref,)).fetchone()
                tid = int(ref) if ref.isdigit() else user[0] if user else None
                if tid and conn.execute("SELECT 1 FROM employees WHERE telegram_id=? AND name_key<>?", (tid, row['name'].casefold())).fetchone():
                    raise ValidationError([f"{row['row_number']}-qator: Telegram akkaunt boshqa xodimda mavjud."])
                try:
                    previous = conn.execute("SELECT * FROM employees WHERE name_key=?", (row["name"].casefold(),)).fetchone()
                    if previous and (previous["telegram_ref"] != ref or previous["phone"] != row["phone"] or previous['phone_extra'] != row['phone_extra']):
                        conn.execute("UPDATE join_requests SET status='superseded' WHERE employee_id=?", (previous["id"],))
                    conn.execute("""INSERT INTO employees(name,name_key,phone,telegram_ref,telegram_id,shift,phone_extra)
                        VALUES (?,?,?,?,?,?,?) ON CONFLICT(name_key) DO UPDATE SET
                        name=excluded.name, phone=excluded.phone, phone_extra=excluded.phone_extra, shift=excluded.shift,
                        approved=CASE WHEN employees.telegram_ref=excluded.telegram_ref AND employees.phone=excluded.phone AND employees.phone_extra IS excluded.phone_extra
                            THEN employees.approved ELSE 0 END,
                        telegram_id=CASE WHEN employees.telegram_ref=excluded.telegram_ref
                            THEN COALESCE(employees.telegram_id,excluded.telegram_id) ELSE excluded.telegram_id END,
                        telegram_ref=excluded.telegram_ref""",
                        (row["name"], row["name"].casefold(), row["phone"], ref, tid, row["shift"], row['phone_extra']))
                except sqlite3.IntegrityError:
                    raise ValidationError([f"{row['row_number']}-qator: Telegram identifikatori boshqa xodimga tegishli."]) from None

    @staticmethod
    def resolve_employee(conn, reference):
        reference = (reference or "").strip()
        if not reference:
            return None
        if reference.startswith(EMPLOYEE_PHONE_REFERENCE):
            phone, _, qualifier = reference.removeprefix(EMPLOYEE_PHONE_REFERENCE).partition("|")
            phone = phone_number(phone)
            candidates = conn.execute("""SELECT id,name,telegram_ref FROM employees
                WHERE phone=? OR phone_extra=?""", (phone, phone)).fetchall()
            if qualifier:
                key = employee_name_key(qualifier)
                candidates = [row for row in candidates
                              if employee_name_key(row["name"]) == key or row["telegram_ref"] == qualifier.casefold()]
            if not candidates:
                raise ValueError("xodim telefoni bo‘yicha moslik topilmadi")
            if len(candidates) > 1:
                raise ValueError("umumiy telefon uchun xodim ustunida ism yoki Telegram ID ko‘rsating")
            return candidates[0]["id"]
        rows = conn.execute("SELECT id FROM employees WHERE employee_name_key(name)=? OR telegram_ref=? LIMIT 2",
                            (employee_name_key(reference), reference.lower())).fetchall()
        if not rows:
            raise ValueError("xodim topilmadi; oldin xodimlar ro'yxatini yuklang")
        if len(rows) > 1:
            raise ValueError("bir nechta xodim mos keldi; Telegram ID yoki @username orqali xodimni aniq tanlang")
        return rows[0][0]

    def import_appeals(self, rows):
        added = skipped = 0
        seen = set()
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            for row in rows:
                location = row.get("row_location") or f"{row.get('row_number', 1)}-qator"
                if row['id'] in seen:
                    raise ValidationError([f"{location}: ID '{row['id']}' faylda takrorlangan; har bir ID yagona bo'lsin."])
                seen.add(row['id'])
                reference = row.get("employee_ref") or ""
                try:
                    employee_id = self.resolve_employee(conn, reference)
                except ValueError as exc:
                    if row.get("legacy") or reference.startswith(EMPLOYEE_PHONE_REFERENCE):
                        employee_id = None
                    else:
                        raise ValidationError([f"{location}: {exc}."]) from None
                source_json = json.dumps(row.get("source_data", {}), ensure_ascii=False)
                previous = conn.execute("SELECT * FROM appeals WHERE id=?", (row["id"],)).fetchone()
                if previous:
                    due_at = row["due_at"] if row["due_at"] is not None else previous["due_at"]
                    employee_id = employee_id if reference else previous["employee_id"]
                    hint = reference or previous["employee_hint"]
                    if not row.get("source_data"):
                        source_json = previous["source_json"]
                    changed = (any(previous[key] != row[key] for key in ("client_name", "phone", "body")) or
                               previous["due_at"] != due_at or previous["employee_id"] != employee_id or
                               previous["employee_hint"] != hint or previous["source_json"] != source_json)
                    if changed:
                        if previous['completed_at'] is not None or conn.execute("SELECT 1 FROM deliveries WHERE appeal_id=? AND (attempts>0 OR status IN ('sent','sending','unknown'))", (row["id"],)).fetchone():
                            raise ValidationError([f"{location}: ID '{row['id']}' mavjud; yuborish boshlangan yoki bajarilgan murojaatni o'zgartirib bo'lmaydi. Yangi murojaatga yangi ID bering. Butun import bekor qilindi."])
                        conn.execute("""UPDATE appeals SET client_name=?,phone=?,body=?,due_at=?,employee_id=?,employee_hint=?,source_json=?,
                            schedule_source=CASE WHEN ? THEN 'explicit' ELSE schedule_source END WHERE id=?""",
                                     (row["client_name"], row["phone"], row["body"], due_at, employee_id, hint, source_json,
                                      row["due_at"] is not None and previous["due_at"] != due_at, row["id"]))
                        added += 1
                        continue
                    skipped += 1
                    continue
                conn.execute("""INSERT INTO appeals(id,client_name,phone,body,due_at,employee_id,created_at,source_user,employee_hint,source_json,legacy)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                             (row["id"], row["client_name"], row["phone"], row["body"], row["due_at"], employee_id,
                              time.time(), row.get("source_user"), reference, source_json, int(row.get("legacy", False))))
                conn.executemany("INSERT INTO deliveries(appeal_id,channel,updated_at) VALUES (?,?,?)",
                                 [(row["id"], channel, time.time()) for channel in ("telegram", "sms")])
                added += 1
        return added, skipped

    def recover(self):
        with self.connect() as conn:
            conn.execute("UPDATE deliveries SET status='unknown', last_error='Jarayon yuborish paytida uzilgan' WHERE status='sending'")
            conn.execute("UPDATE notices SET status='unknown' WHERE status='sending'")
            conn.execute("UPDATE notices SET status='superseded' WHERE notice_key GLOB 'assignment:*' AND status IN ('pending','retry')")
            conn.execute("UPDATE sms_deliveries SET status='unknown',last_error='Jarayon yuborish paytida uzilgan' WHERE status='sending'")
            for row in conn.execute("SELECT DISTINCT appeal_id FROM sms_deliveries").fetchall():
                self._aggregate_sms(conn, row[0])

    def default_schedule(self):
        with self.connect() as conn:
            row = conn.execute("SELECT value FROM settings WHERE key='default_send_time'").fetchone()
            return row[0] if row else None

    def sms_enabled(self):
        with self.connect() as conn:
            row = conn.execute("SELECT value FROM settings WHERE key='sms_enabled'").fetchone()
            return row is None or row[0] != '0'

    def set_sms_enabled(self, enabled):
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            self._set_sms_enabled(conn, enabled)

    def _set_sms_enabled(self, conn, enabled):
        if enabled:
            self._hold_overdue_sms(conn, previous_setting='0')
        conn.execute("INSERT INTO settings(key,value) VALUES ('sms_enabled',?) "
                     "ON CONFLICT(key) DO UPDATE SET value=excluded.value", ('1' if enabled else '0',))

    @staticmethod
    def _sms_password_hash(password, salt):
        return hashlib.pbkdf2_hmac("sha256", password.encode(), salt, SMS_PASSWORD_ITERATIONS)

    def set_sms_password(self, password):
        if len(password) < 12 or len(password) > 128 or not password.strip():
            raise ValueError("SMS paroli 12–128 belgidan iborat bo‘lsin.")
        if password.startswith("/"):
            raise ValueError("SMS parolini harf, raqam yoki boshqa belgidan boshlang. Telegram / bilan boshlangan matnni buyruq deb qabul qiladi.")
        salt = os.urandom(16)
        digest = self._sms_password_hash(password, salt)
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute("INSERT INTO settings VALUES ('sms_password_hash',?) "
                         "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (f"{salt.hex()}:{digest.hex()}",))
            conn.execute("DELETE FROM settings WHERE key GLOB 'sms_password_attempts:*'")

    def sms_password_configured(self):
        with self.connect() as conn:
            return conn.execute("SELECT 1 FROM settings WHERE key='sms_password_hash'").fetchone() is not None

    def change_sms_with_password(self, enabled, password, admin_id):
        attempt_key = f"sms_password_attempts:{admin_id}"
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            stored = conn.execute("SELECT value FROM settings WHERE key='sms_password_hash'").fetchone()
            if stored is None:
                return "unconfigured"
            attempts = conn.execute("SELECT value FROM settings WHERE key=?", (attempt_key,)).fetchone()
            count, deadline = map(int, attempts[0].split(":")) if attempts else (0, 0)
            now = int(time.time())
            if deadline <= now:
                count, deadline = 0, now + SMS_PASSWORD_WINDOW
            if count >= SMS_PASSWORD_ATTEMPTS:
                return "locked"
            salt, expected = map(bytes.fromhex, stored[0].split(":"))
            if 12 <= len(password) <= 128 and hmac.compare_digest(self._sms_password_hash(password, salt), expected):
                self._set_sms_enabled(conn, enabled)
                conn.execute("DELETE FROM settings WHERE key=?", (attempt_key,))
                return "changed"
            count += 1
            conn.execute("INSERT INTO settings VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                         (attempt_key, f"{count}:{deadline}"))
            return "locked" if count >= SMS_PASSWORD_ATTEMPTS else "invalid"

    @staticmethod
    def _hold_overdue_sms(conn, previous_setting):
        conn.execute("""UPDATE appeals SET sms_held=1
            WHERE due_at<=? AND completed_at IS NULL
            AND EXISTS (SELECT 1 FROM settings WHERE key='sms_enabled' AND value=?)
            AND EXISTS (SELECT 1 FROM deliveries WHERE appeal_id=appeals.id AND channel='sms' AND status!='sent')""",
            (time.time(), previous_setting))

    @staticmethod
    def next_send_time(mode, timezone, now):
        timezone = TIMEZONE
        if mode == "darhol":
            return now
        hour, minute = map(int, mode.split(":"))
        local = datetime.fromtimestamp(now, timezone)
        target = local.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if target < local:
            target += timedelta(days=1)
        return target.timestamp()

    def set_default_schedule(self, mode, timezone, now=None):
        mode = mode.strip().lower()
        if mode != "darhol" and not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", mode):
            raise ValueError("Yuborish vaqtini HH:MM formatida kiriting (masalan 09:00) yoki darhol deb yozing.")
        now = time.time() if now is None else now
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute("INSERT INTO settings VALUES ('default_send_time',?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (mode,))
            return self._schedule_waiting(conn, mode, timezone, now, change_future=True)

    @classmethod
    def _schedule_waiting(cls, conn, mode, timezone, now, change_future=False):
        due_at = cls.next_send_time(mode, timezone, now)
        result = conn.execute("""UPDATE appeals SET due_at=?,schedule_source='automatic'
            WHERE completed_at IS NULL AND (legacy=0 OR (employee_id IS NOT NULL AND employee_hint<>''))
            AND (due_at IS NULL OR (? AND schedule_source='automatic' AND due_at>?))
            AND NOT EXISTS (SELECT 1 FROM deliveries d WHERE d.appeal_id=appeals.id
                AND (d.attempts>0 OR d.status IN ('sent','sending','unknown')))""", (due_at, int(change_future), now))
        return result.rowcount

    def apply_default_schedule(self, timezone, now=None):
        now = time.time() if now is None else now
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            row = conn.execute("SELECT value FROM settings WHERE key='default_send_time'").fetchone()
            return self._schedule_waiting(conn, row[0], timezone, now) if row else 0

    def assign_due(self, now, timezone):
        timezone = TIMEZONE
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            rows = conn.execute("""SELECT id,due_at,employee_hint,legacy FROM appeals
                WHERE employee_id IS NULL AND completed_at IS NULL ORDER BY due_at,id""").fetchall()
            for row in rows:
                if row["employee_hint"]:
                    try:
                        employee_id = self.resolve_employee(conn, row["employee_hint"])
                        conn.execute("UPDATE appeals SET employee_id=? WHERE id=?", (employee_id, row["id"]))
                    except ValueError:
                        pass
                    continue
                if row["legacy"] or row["due_at"] is None:
                    continue
                employees = conn.execute("""SELECT e.*, COUNT(a.id) AS workload FROM employees e
                    LEFT JOIN appeals a ON e.id=a.employee_id WHERE e.approved=1 GROUP BY e.id ORDER BY workload,e.id""").fetchall()
                moment = datetime.fromtimestamp(row["due_at"], timezone)
                selected = next((e for e in employees if in_shift(e["shift"], moment)), None)
                if selected:
                    conn.execute("UPDATE appeals SET employee_id=? WHERE id=?", (selected["id"], row["id"]))

    def ready_deliveries(self, now, sms_target="both"):
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            # Freeze recipients when a due, approved delivery first enters the worker.
            for row in conn.execute("""SELECT d.*,e.phone,e.phone_extra FROM deliveries d
                JOIN appeals a ON a.id=d.appeal_id JOIN employees e ON e.id=a.employee_id
                WHERE d.channel='sms' AND a.due_at<=? AND a.completed_at IS NULL AND a.sms_held=0 AND e.approved=1 AND e.telegram_id IS NOT NULL
                AND (a.legacy=0 OR a.employee_hint<>'')
                AND NOT EXISTS (SELECT 1 FROM settings WHERE key='sms_enabled' AND value='0')""", (now,)).fetchall():
                existing = conn.execute('SELECT * FROM sms_deliveries WHERE appeal_id=?', (row['appeal_id'],)).fetchall()
                if existing and any(s['attempts'] > 0 or s['status'] not in ('pending','blocked') for s in existing):
                    continue
                phones = [row['phone']]
                # Historic completed/uncertain deliveries must never be expanded silently.
                if sms_target == 'both' and row['phone_extra'] and row['status'] not in ('sent','unknown','sending'):
                    phones.append(row['phone_extra'])
                for target in existing:
                    if target['phone'] not in phones:
                        conn.execute('DELETE FROM sms_deliveries WHERE appeal_id=? AND phone=?', (row['appeal_id'], target['phone']))
                for index, phone in enumerate(phones):
                    conn.execute("""INSERT OR IGNORE INTO sms_deliveries VALUES (?,?,?,?,?,?,?,?)""",
                        (row['appeal_id'], phone, row['status'] if index == 0 else 'pending',
                         row['attempts'] if index == 0 else 0, row['next_attempt'] if index == 0 else 0,
                         row['last_error'] if index == 0 else None, row['message_id'] if index == 0 else None, now))
                self._aggregate_sms(conn, row['appeal_id'])
            telegram = [dict(row) for row in conn.execute("""SELECT d.*,a.client_name,a.phone AS client_phone,a.body,a.source_json,
                e.name AS employee_name,e.phone AS employee_phone,e.telegram_id
                FROM deliveries d JOIN appeals a ON a.id=d.appeal_id JOIN employees e ON e.id=a.employee_id
                WHERE d.channel='telegram' AND a.due_at<=? AND a.completed_at IS NULL AND e.approved=1 AND d.status IN ('pending','retry','blocked')
                AND (a.legacy=0 OR a.employee_hint<>'')
                AND d.next_attempt<=? ORDER BY a.due_at,d.channel LIMIT 100""", (now, now))]
            sms_rows = [dict(row) for row in conn.execute("""SELECT s.*, 'sms' AS channel, s.phone AS target_phone,
                a.rowid AS appeal_number,a.client_name,a.phone AS client_phone,a.body,
                e.name AS employee_name,e.phone AS employee_phone,e.telegram_id
                FROM sms_deliveries s JOIN appeals a ON a.id=s.appeal_id JOIN employees e ON e.id=a.employee_id
                WHERE a.due_at<=? AND a.completed_at IS NULL AND a.sms_held=0 AND e.approved=1 AND s.status IN ('pending','retry','blocked') AND s.next_attempt<=?
                AND (a.legacy=0 OR a.employee_hint<>'')
                AND NOT EXISTS (SELECT 1 FROM settings WHERE key='sms_enabled' AND value='0')
                ORDER BY a.due_at LIMIT 100""", (now, now))]
            return telegram + sms_rows

    @staticmethod
    def _aggregate_sms(conn, appeal_id):
        rows = conn.execute("SELECT * FROM sms_deliveries WHERE appeal_id=?", (appeal_id,)).fetchall()
        if not rows:
            return
        priority = ('unknown','failed','sending','retry','blocked','pending','sent')
        status = next(value for value in priority if any(row['status'] == value for row in rows))
        attempts = max(row['attempts'] for row in rows)
        next_attempt = min(row['next_attempt'] for row in rows)
        error = next(row['last_error'] for row in rows if row['status'] == status)
        conn.execute("""UPDATE deliveries SET status=?,attempts=?,next_attempt=?,last_error=?,updated_at=?
            WHERE appeal_id=? AND channel='sms'
            AND (status IS NOT ? OR attempts IS NOT ? OR next_attempt IS NOT ? OR last_error IS NOT ?)""",
            (status, attempts, next_attempt, error, time.time(), appeal_id, status, attempts, next_attempt, error))

    def claim_sms(self, appeal_id, phone, now):
        with self.connect() as conn:
            result = conn.execute("""UPDATE sms_deliveries SET status='sending',attempts=attempts+1,updated_at=?
                WHERE appeal_id=? AND phone=? AND status IN ('pending','retry','blocked') AND next_attempt<=?
                AND NOT EXISTS (SELECT 1 FROM settings WHERE key='sms_enabled' AND value='0')
                AND EXISTS (SELECT 1 FROM appeals a JOIN employees e ON e.id=a.employee_id
                    WHERE a.id=sms_deliveries.appeal_id AND a.due_at<=? AND a.completed_at IS NULL AND a.sms_held=0
                    AND (a.legacy=0 OR a.employee_hint<>'') AND e.approved=1 AND e.telegram_id IS NOT NULL
                    AND sms_deliveries.phone IN (e.phone,e.phone_extra))""",
                (now, appeal_id, phone, now, now))
            self._aggregate_sms(conn, appeal_id)
            return result.rowcount == 1

    def finish_sms(self, appeal_id, phone, status, error=None, message_id=None, delay=0, refund=False):
        with self.connect() as conn:
            conn.execute("""UPDATE sms_deliveries SET status=?,last_error=?,message_id=?,next_attempt=?,updated_at=?,
                attempts=MAX(0,attempts-?) WHERE appeal_id=? AND phone=? AND status='sending'""",
                (status,error,str(message_id) if message_id is not None else None,time.time()+delay,time.time(),int(refund),appeal_id,phone))
            self._aggregate_sms(conn, appeal_id)

    def sms_details(self, appeal_id):
        with self.connect() as conn:
            return [dict(row) for row in conn.execute("SELECT * FROM sms_deliveries WHERE appeal_id=? ORDER BY phone", (appeal_id,))]

    def claim(self, appeal_id, channel, now, *, delivery=None):
        with self.connect() as conn:
            expected = ""
            values = ()
            if delivery is not None:
                expected = "AND e.telegram_id IS ? AND a.client_name=? AND a.phone=? AND a.body=? AND a.source_json=?"
                values = (delivery["telegram_id"], delivery["client_name"], delivery["client_phone"], delivery["body"], delivery["source_json"])
            result = conn.execute(f"""UPDATE deliveries SET status='sending',attempts=attempts+1,updated_at=?
                WHERE appeal_id=? AND channel=? AND status IN ('pending','retry','blocked') AND next_attempt<=?
                AND (channel!='sms' OR NOT EXISTS (SELECT 1 FROM settings WHERE key='sms_enabled' AND value='0'))
                AND EXISTS (SELECT 1 FROM appeals a JOIN employees e ON e.id=a.employee_id
                    WHERE a.id=deliveries.appeal_id AND a.due_at<=? AND a.completed_at IS NULL AND e.approved=1
                    AND (a.legacy=0 OR a.employee_hint<>'')
                    AND (channel!='sms' OR a.sms_held=0) {expected})""",
                (now, appeal_id, channel, now, now, *values))
            return result.rowcount == 1

    def finish(self, appeal_id, channel, status, error=None, message_id=None, delay=0, refund=False):
        with self.connect() as conn:
            conn.execute("""UPDATE deliveries SET status=?,last_error=?,message_id=?,next_attempt=?,updated_at=?,
                attempts=MAX(0,attempts-?) WHERE appeal_id=? AND channel=? AND status='sending'""",
                (status, error, str(message_id) if message_id is not None else None, time.time()+delay,
                 time.time(), int(refund), appeal_id, channel))

    def advance_part(self, appeal_id, part_index, message_id):
        with self.connect() as conn:
            result = conn.execute("""UPDATE deliveries SET part_index=?,message_id=?,updated_at=?
                WHERE appeal_id=? AND channel='telegram' AND status='sending' AND part_index=?""",
                (part_index, str(message_id), time.time(), appeal_id, part_index - 1))
            if result.rowcount != 1:
                raise ValueError("Xabar qismi holati o'zgargan; takror yuborish to'xtatildi.")

    def set_time(self, appeal_id, due_at):
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            if not conn.execute("SELECT 1 FROM appeals WHERE id=?", (appeal_id,)).fetchone():
                raise ValueError("Murojaat topilmadi.")
            if conn.execute("SELECT 1 FROM appeals WHERE id=? AND completed_at IS NOT NULL", (appeal_id,)).fetchone() or conn.execute("SELECT 1 FROM deliveries WHERE appeal_id=? AND (attempts>0 OR status IN ('sent','sending','unknown'))", (appeal_id,)).fetchone():
                raise ValueError("Yuborish boshlangan; vaqtni o'zgartirib bo'lmaydi.")
            conn.execute("UPDATE appeals SET due_at=?,schedule_source='explicit' WHERE id=?", (due_at, appeal_id))

    def assign(self, appeal_id, reference):
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            employee_id = self.resolve_employee(conn, reference)
            if conn.execute("SELECT 1 FROM appeals WHERE id=? AND completed_at IS NOT NULL", (appeal_id,)).fetchone() or conn.execute("SELECT 1 FROM deliveries WHERE appeal_id=? AND (attempts>0 OR status IN ('sent','sending','unknown'))", (appeal_id,)).fetchone():
                raise ValueError("Yuborish boshlangan; xodimni almashtirib bo'lmaydi.")
            if not conn.execute("UPDATE appeals SET employee_id=?,employee_hint=? WHERE id=?", (employee_id, reference, appeal_id)).rowcount:
                raise ValueError("Murojaat topilmadi.")

    def retry(self, appeal_id, channel):
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            if conn.execute("SELECT 1 FROM appeals WHERE id=? AND completed_at IS NOT NULL", (appeal_id,)).fetchone():
                raise ValueError("Bajarilgan murojaat qayta yuborilmaydi.")
            if channel == 'sms':
                held = conn.execute("SELECT sms_held FROM appeals WHERE id=?", (appeal_id,)).fetchone()
                if held and held[0]:
                    recipients = conn.execute("SELECT 1 FROM sms_deliveries WHERE appeal_id=?", (appeal_id,)).fetchone()
                    table = 'sms_deliveries' if recipients else 'deliveries'
                    channel_filter = "" if recipients else "AND channel='sms'"
                    changed = conn.execute(f"""UPDATE {table}
                        SET attempts=CASE WHEN status IN ('failed','blocked') THEN 0 ELSE attempts END,
                            next_attempt=0,last_error=NULL,status='pending'
                        WHERE appeal_id=? {channel_filter} AND status IN ('pending','retry','failed','blocked')""",
                        (appeal_id,)).rowcount
                    if not changed:
                        raise ValueError("Navbatga qo‘yish mumkin bo‘lgan SMS topilmadi.")
                    conn.execute("UPDATE appeals SET sms_held=0 WHERE id=?", (appeal_id,))
                    self._aggregate_sms(conn, appeal_id)
                    return
            if channel == 'sms' and conn.execute("SELECT 1 FROM sms_deliveries WHERE appeal_id=?", (appeal_id,)).fetchone():
                changed = conn.execute("""UPDATE sms_deliveries SET status='pending',attempts=0,next_attempt=0,last_error=NULL
                    WHERE appeal_id=? AND status IN ('failed','blocked')""", (appeal_id,)).rowcount
                if not changed:
                    raise ValueError("Qayta urinish uchun xato yoki API kutilayotgan SMS topilmadi.")
                self._aggregate_sms(conn, appeal_id)
                return
            if not conn.execute("""UPDATE deliveries SET status='pending',attempts=0,next_attempt=0,last_error=NULL
                WHERE appeal_id=? AND channel=? AND status IN ('failed','blocked')""", (appeal_id, channel)).rowcount:
                raise ValueError("Faqat xato yoki SMS API kutilayotgan kanalni qayta yuborish mumkin.")

    def resolve_unknown(self, appeal_id, channel, sent):
        with self.connect() as conn:
            if channel == 'sms' and conn.execute("SELECT 1 FROM sms_deliveries WHERE appeal_id=?", (appeal_id,)).fetchone():
                changed = conn.execute("""UPDATE sms_deliveries SET status=?,attempts=0,next_attempt=0,last_error=NULL
                    WHERE appeal_id=? AND status='unknown'""", ('sent' if sent else 'pending', appeal_id)).rowcount
                if not changed:
                    raise ValueError("Noma'lum SMS topilmadi.")
                self._aggregate_sms(conn, appeal_id)
                return
            result = conn.execute("""UPDATE deliveries SET status=?,attempts=0,next_attempt=0,last_error=NULL
                WHERE appeal_id=? AND channel=? AND status='unknown'""", ("sent" if sent else "pending", appeal_id, channel))
            if not result.rowcount:
                raise ValueError("Noma'lum holatdagi yuborish topilmadi.")

    def status_text(self, include_sms=True):
        with self.connect() as conn:
            total = conn.execute("SELECT COUNT(*) FROM appeals").fetchone()[0]
            unscheduled = conn.execute("SELECT COUNT(*) FROM appeals WHERE due_at IS NULL").fetchone()[0]
            unassigned = conn.execute("SELECT COUNT(*) FROM appeals WHERE employee_id IS NULL").fetchone()[0]
            unapproved = conn.execute("SELECT COUNT(*) FROM employees WHERE approved=0").fetchone()[0]
            rows = conn.execute("SELECT channel,status,COUNT(*) AS n FROM deliveries GROUP BY channel,status").fetchall()
        lines = [f"Jami murojaat: {total}", f"Vaqt belgilanmagan: {unscheduled}", f"Xodim biriktirilmagan: {unassigned}",
                 f"Tasdiqlanmagan xodimlar: {unapproved}"]
        appeals = self.list_appeals(None)
        if not include_sms:
            sent = sum(row['telegram'] == 'sent' for row in appeals)
            unsent = sum(row['telegram'] != 'sent' and row['completed_at'] is None for row in appeals)
            completed = sum(row['completed_at'] is not None for row in appeals)
            return (f'Jami murojaat: {total}\n✅ Yuborilgan: {sent}\n⏳ Yuborilmagan: {unsent}'
                    f'\nBajarildi: {completed}\nIjro kutilmoqda: {total-completed}')
        lines += [f"Mos xodim topilmagan murojaatlar: {sum(assignment_problem(r) == 'unassigned' for r in appeals)}",
                  f"Ulanmagan xodimga biriktirilgan murojaatlar: {sum(assignment_problem(r) == 'unconnected' for r in appeals)}",
                  f"Bajarildi: {sum(r['completed_at'] is not None for r in appeals)}",
                  f"Ijro kutilmoqda: {sum(r['completed_at'] is None for r in appeals)}"]
        mode = self.default_schedule()
        lines.insert(0, 'SMS yuborish: ' + ('yoqilgan' if self.sms_enabled() else 'o‘chirilgan'))
        lines.insert(0, "Avtomatik yuborish: " + ("yuklangach darhol" if mode == "darhol" else f"har kuni {mode}" if mode else "sozlanmagan — /jadval ni tanlang"))
        for channel in ("telegram", "sms"):
            lines.append("\n" + channel.upper() + ":")
            counts = {r["status"]: r["n"] for r in rows if r["channel"] == channel}
            lines.extend(f"{label}: {counts.get(key, 0)}" for key, label in STATUSES.items())
        return "\n".join(lines)

    def list_appeals(self, timezone, today=False, *, sent_today=False):
        timezone = TIMEZONE
        now = datetime.now(timezone)
        start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        with self.connect() as conn:
            query = """SELECT a.id,a.due_at,a.employee_id,a.completed_at,a.completed_by,a.sms_held,COALESCE(e.name,NULLIF(a.employee_hint,'')) AS employee,
                e.approved AS employee_approved,e.telegram_id AS employee_telegram_id,
                t.status AS telegram,t.updated_at AS telegram_updated_at,s.status AS sms FROM appeals a
                LEFT JOIN employees e ON a.employee_id=e.id
                JOIN deliveries t ON t.appeal_id=a.id AND t.channel='telegram'
                JOIN deliveries s ON s.appeal_id=a.id AND s.channel='sms'"""
            params = ()
            if sent_today:
                query += " WHERE t.status='sent' AND t.updated_at>=? AND t.updated_at<?"
                params = (start.timestamp(), (start + timedelta(days=1)).timestamp())
            elif today:
                query += " WHERE a.due_at>=? AND a.due_at<?"
                params = (start.timestamp(), (start + timedelta(days=1)).timestamp())
            query += " ORDER BY a.due_at,a.id"
            return [dict(row) for row in conn.execute(query, params)]

    def employee_list(self):
        with self.connect() as conn:
            return [dict(row) for row in conn.execute("SELECT * FROM employees ORDER BY name_key")]

    def export_appeals(self, telegram_id=None, appeal_id=None):
        with self.connect() as conn:
            query = "SELECT a.*,e.name AS employee FROM appeals a LEFT JOIN employees e ON e.id=a.employee_id WHERE 1=1"
            params = []
            if telegram_id is not None:
                query += " AND e.telegram_id=? AND e.approved=1 AND a.due_at<=?"
                params.extend([telegram_id, time.time()])
            if appeal_id is not None:
                query += " AND a.id=?"
                params.append(appeal_id)
            return [dict(row) for row in conn.execute(query + " ORDER BY a.due_at,a.id", params)]

    def details(self, appeal_id):
        with self.connect() as conn:
            row = conn.execute("SELECT * FROM appeals WHERE id=?", (appeal_id,)).fetchone()
            states = conn.execute("SELECT * FROM deliveries WHERE appeal_id=?", (appeal_id,)).fetchall()
            return (dict(row), [dict(s) for s in states]) if row else None

    def enqueue_notice(self, key, admin_ids, text):
        with self.connect() as conn:
            conn.executemany("INSERT OR IGNORE INTO notices(notice_key,admin_id,text) VALUES (?,?,?)",
                             [(f"{key}:{admin}", admin, text) for admin in admin_ids])

    def ready_notices(self, now):
        with self.connect() as conn:
            return [dict(row) for row in conn.execute("SELECT * FROM notices WHERE status IN ('pending','retry') AND next_attempt<=? LIMIT 50", (now,))]

    def claim_notice(self, key):
        with self.connect() as conn:
            return conn.execute("UPDATE notices SET status='sending',attempts=attempts+1,updated_at=? WHERE notice_key=? AND status IN ('pending','retry')", (time.time(), key)).rowcount == 1

    def finish_notice(self, key, status, delay=0):
        with self.connect() as conn:
            conn.execute("UPDATE notices SET status=?,next_attempt=?,updated_at=? WHERE notice_key=? AND status='sending'", (status, time.time()+delay, time.time(), key))

    def expire_sending(self, now, timeout=120):
        with self.connect() as conn:
            cutoff = now - timeout
            error = "Yuborish timeouti; natijani qabul qiluvchidan tekshiring."
            conn.execute("UPDATE deliveries SET status='unknown',last_error=? WHERE status='sending' AND updated_at<=? AND channel='telegram'", (error, cutoff))
            conn.execute("""UPDATE deliveries SET status='unknown',last_error=? WHERE status='sending' AND updated_at<=?
                AND channel='sms' AND NOT EXISTS (SELECT 1 FROM sms_deliveries s WHERE s.appeal_id=deliveries.appeal_id)""", (error, cutoff))
            conn.execute("UPDATE sms_deliveries SET status='unknown',last_error=? WHERE status='sending' AND updated_at<=?", (error, cutoff))
            conn.execute("UPDATE notices SET status='unknown' WHERE status='sending' AND updated_at<=?", (cutoff,))
            for row in conn.execute("SELECT DISTINCT appeal_id FROM sms_deliveries WHERE status='unknown'").fetchall():
                self._aggregate_sms(conn, row[0])

    def disconnect_employee(self, employee_id):
        with self.connect() as conn:
            conn.execute('BEGIN IMMEDIATE')
            employee = conn.execute('SELECT * FROM employees WHERE id=?', (employee_id,)).fetchone()
            if not employee:
                raise ValueError('Xodim topilmadi.')
            conn.execute("UPDATE employees SET approved=0,telegram_id=NULL,telegram_ref='phone:'||phone WHERE id=?", (employee_id,))
            conn.execute("UPDATE join_requests SET status='superseded' WHERE employee_id=? AND status='pending'", (employee_id,))
            identifiers = [r[0] for r in conn.execute('SELECT id FROM appeals WHERE employee_id=?', (employee_id,))]
            for identifier in identifiers:
                conn.execute("UPDATE deliveries SET status='unknown',last_error='Xodim yuborish paytida uzildi' WHERE appeal_id=? AND status='sending'", (identifier,))
                conn.execute("UPDATE sms_deliveries SET status='unknown',last_error='Xodim yuborish paytida uzildi' WHERE appeal_id=? AND status='sending'", (identifier,))
            if employee['telegram_id']:
                conn.execute('INSERT INTO notices(notice_key,admin_id,text) VALUES (?,?,?)',
                             (f"disconnect:{employee_id}:{time.time_ns()}", employee['telegram_id'],
                              "Bot bilan ulanishingiz admin tomonidan uzildi. Qayta ulanish uchun Ulanish tugmasini bosing."))

    def complete_appeal(self, identifier, telegram_id, admin_ids):
        with self.connect() as conn:
            conn.execute('BEGIN IMMEDIATE')
            row = conn.execute("""SELECT a.* FROM appeals a JOIN employees e ON e.id=a.employee_id
                WHERE a.id=? AND e.telegram_id=? AND e.approved=1 AND a.due_at<=?""", (identifier, telegram_id, time.time())).fetchone()
            if not row:
                raise ValueError('Bu murojaat sizga biriktirilmagan yoki vaqti kelmagan.')
            if row['completed_at'] is not None:
                return False
            conn.execute('UPDATE appeals SET completed_at=?,completed_by=? WHERE id=?', (time.time(), telegram_id, identifier))
            conn.executemany('INSERT OR IGNORE INTO notices(notice_key,admin_id,text) VALUES (?,?,?)',
                             [(f'completed:{identifier}:{admin_id}', admin_id, f'Murojaat #{identifier}: ✅ Bajarildi.') for admin_id in admin_ids])
            return True

    @staticmethod
    def create_backup(source, target):
        source, target = Path(source), Path(target)
        target.parent.mkdir(parents=True, exist_ok=True)
        partial = target.with_suffix('.partial')
        started = time.monotonic()
        def progress(status, remaining, total):
            if time.monotonic() - started > 30:
                raise TimeoutError('SQLite backup timeout')
        with closing(sqlite3.connect(source.resolve().as_uri() + '?mode=ro', uri=True, timeout=5)) as origin:
            with closing(sqlite3.connect(partial, timeout=5)) as destination:
                origin.backup(destination, pages=128, progress=progress)
                if destination.execute('PRAGMA quick_check').fetchone()[0] != 'ok':
                    raise RuntimeError('SQLite backup verification failed')
        os.replace(partial, target)
