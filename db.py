"""SQLite outbox and atomic state transitions."""
import json
import re
import sqlite3
import time
from contextlib import contextmanager
from datetime import datetime, timedelta

from excel_utils import ValidationError, in_shift, phone_number

STATUSES = {
    "pending": "kutmoqda", "retry": "qayta urinish", "sending": "yuborilmoqda",
    "sent": "yuborildi", "failed": "xato", "unknown": "natija noma'lum",
    "blocked": "SMS API kutilmoqda",
}


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
                    phone TEXT UNIQUE NOT NULL, telegram_ref TEXT UNIQUE NOT NULL,
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
            # Additive migration: existing contacts also require explicit approval.
            for table, columns in {
                "employees": {"approved": "INTEGER NOT NULL DEFAULT 0", "phone_extra": "TEXT"},
                "deliveries": {"part_index": "INTEGER NOT NULL DEFAULT 0"},
                "appeals": {"employee_hint": "TEXT NOT NULL DEFAULT ''", "source_json": "TEXT NOT NULL DEFAULT '{}'",
                            "legacy": "INTEGER NOT NULL DEFAULT 0", "schedule_source": "TEXT NOT NULL DEFAULT 'explicit'"},
            }.items():
                present = {row[1] for row in conn.execute(f"PRAGMA table_info({table})")}
                for name, definition in columns.items():
                    if name not in present:
                        conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {definition}")

    @contextmanager
    def connect(self):
        conn = sqlite3.connect(self.path, timeout=15)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys=ON")
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
            employee = conn.execute("SELECT * FROM employees WHERE id=?", (employee_id,)).fetchone()
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

    def join_by_phone(self, phone, telegram_id):
        with self.connect() as conn:
            phone = phone_number(phone)
            row = conn.execute("SELECT id FROM employees WHERE phone=? OR phone_extra=?", (phone, phone)).fetchone()
        if not row:
            raise ValueError("Bu telefon xodimlar ro'yxatida topilmadi. Admin ro'yxatni tekshirsin.")
        return self.request_join(row[0], telegram_id)

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
                conflict = conn.execute("SELECT id FROM employees WHERE approved=1 AND telegram_id=? AND id<>?",
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
                conflict = conn.execute("""SELECT id FROM employees WHERE name_key<>? AND
                    (phone IN (?,?) OR phone_extra IN (?,?))""", (row['name'].casefold(), row['phone'], row['phone_extra'],
                    row['phone'], row['phone_extra'])).fetchone()
                if conflict:
                    raise ValidationError([f"{row['row_number']}-qator: asosiy yoki qo'shimcha telefon boshqa xodimda mavjud."])
                ref = row["telegram_ref"]
                user = conn.execute("SELECT telegram_id FROM users WHERE username=?", (ref,)).fetchone()
                tid = int(ref) if ref.isdigit() else user[0] if user else None
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
                    raise ValidationError([f"{row['row_number']}-qator: telefon yoki Telegram boshqa xodimga tegishli."]) from None

    @staticmethod
    def resolve_employee(conn, reference):
        if not reference:
            return None
        row = conn.execute("SELECT id FROM employees WHERE name_key=? OR telegram_ref=?",
                           (reference.casefold(), reference.lower())).fetchone()
        if not row:
            raise ValueError("xodim topilmadi; oldin xodimlar ro'yxatini yuklang")
        return row[0]

    def import_appeals(self, rows):
        added = skipped = 0
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            for row in rows:
                try:
                    employee_id = self.resolve_employee(conn, row.get("employee_ref", ""))
                except ValueError as exc:
                    if row.get("legacy"):
                        employee_id = None
                    else:
                        raise ValidationError([f"{row['row_number']}-qator: {exc}."]) from None
                source_json = json.dumps(row.get("source_data", {}), ensure_ascii=False)
                previous = conn.execute("SELECT * FROM appeals WHERE id=?", (row["id"],)).fetchone()
                if previous:
                    due_at = row["due_at"] if row["due_at"] is not None else previous["due_at"]
                    employee_id = employee_id if row.get("employee_ref") else previous["employee_id"]
                    hint = row.get("employee_ref") or previous["employee_hint"]
                    if not row.get("source_data"):
                        source_json = previous["source_json"]
                    changed = (any(previous[key] != row[key] for key in ("client_name", "phone", "body")) or
                               previous["due_at"] != due_at or previous["employee_id"] != employee_id or
                               previous["employee_hint"] != hint or previous["source_json"] != source_json)
                    if changed:
                        if conn.execute("SELECT 1 FROM deliveries WHERE appeal_id=? AND (attempts>0 OR status IN ('sent','sending','unknown'))", (row["id"],)).fetchone():
                            raise ValidationError([f"{row['row_number']}-qator: yuborish boshlangan murojaatni Excel orqali o'zgartirib bo'lmaydi."])
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
                              time.time(), row.get("source_user"), row.get("employee_ref", ""), source_json, int(row.get("legacy", False))))
                conn.executemany("INSERT INTO deliveries(appeal_id,channel,updated_at) VALUES (?,?,?)",
                                 [(row["id"], channel, time.time()) for channel in ("telegram", "sms")])
                added += 1
        return added, skipped

    def recover(self):
        with self.connect() as conn:
            conn.execute("UPDATE deliveries SET status='unknown', last_error='Jarayon yuborish paytida uzilgan' WHERE status='sending'")
            conn.execute("UPDATE notices SET status='unknown' WHERE status='sending'")
            conn.execute("UPDATE sms_deliveries SET status='unknown',last_error='Jarayon yuborish paytida uzilgan' WHERE status='sending'")
            for row in conn.execute("SELECT DISTINCT appeal_id FROM sms_deliveries").fetchall():
                self._aggregate_sms(conn, row[0])

    def default_schedule(self):
        with self.connect() as conn:
            row = conn.execute("SELECT value FROM settings WHERE key='default_send_time'").fetchone()
            return row[0] if row else None

    @staticmethod
    def next_send_time(mode, timezone, now):
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
            WHERE (due_at IS NULL OR (? AND schedule_source='automatic' AND due_at>?))
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
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            rows = conn.execute("SELECT id,due_at,employee_hint FROM appeals WHERE employee_id IS NULL ORDER BY due_at,id").fetchall()
            for row in rows:
                if row["employee_hint"]:
                    try:
                        employee_id = self.resolve_employee(conn, row["employee_hint"])
                        conn.execute("UPDATE appeals SET employee_id=? WHERE id=?", (employee_id, row["id"]))
                    except ValueError:
                        pass
                    continue
                if row["due_at"] is None:
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
                WHERE d.channel='sms' AND a.due_at<=? AND e.approved=1""", (now,)).fetchall():
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
            telegram = [dict(row) for row in conn.execute("""SELECT d.*,a.client_name,a.phone AS client_phone,a.body,
                e.name AS employee_name,e.phone AS employee_phone,e.telegram_id
                FROM deliveries d JOIN appeals a ON a.id=d.appeal_id JOIN employees e ON e.id=a.employee_id
                WHERE d.channel='telegram' AND a.due_at<=? AND e.approved=1 AND d.status IN ('pending','retry','blocked')
                AND d.next_attempt<=? ORDER BY a.due_at,d.channel LIMIT 100""", (now, now))]
            sms_rows = [dict(row) for row in conn.execute("""SELECT s.*, 'sms' AS channel, s.phone AS target_phone,
                a.client_name,a.phone AS client_phone,a.body,e.name AS employee_name,e.phone AS employee_phone,e.telegram_id
                FROM sms_deliveries s JOIN appeals a ON a.id=s.appeal_id JOIN employees e ON e.id=a.employee_id
                WHERE a.due_at<=? AND e.approved=1 AND s.status IN ('pending','retry','blocked') AND s.next_attempt<=?
                ORDER BY a.due_at LIMIT 100""", (now, now))]
            return telegram + sms_rows

    @staticmethod
    def _aggregate_sms(conn, appeal_id):
        rows = conn.execute("SELECT * FROM sms_deliveries WHERE appeal_id=?", (appeal_id,)).fetchall()
        if not rows:
            return
        priority = ('unknown','failed','sending','retry','blocked','pending','sent')
        status = next(value for value in priority if any(row['status'] == value for row in rows))
        conn.execute("""UPDATE deliveries SET status=?,attempts=?,next_attempt=?,last_error=?,updated_at=?
            WHERE appeal_id=? AND channel='sms'""", (status, max(row['attempts'] for row in rows),
            min(row['next_attempt'] for row in rows), next((row['last_error'] for row in rows if row['status']==status), None),
            time.time(), appeal_id))

    def claim_sms(self, appeal_id, phone, now):
        with self.connect() as conn:
            result = conn.execute("""UPDATE sms_deliveries SET status='sending',attempts=attempts+1,updated_at=?
                WHERE appeal_id=? AND phone=? AND status IN ('pending','retry','blocked') AND next_attempt<=?
                AND EXISTS (SELECT 1 FROM appeals a JOIN employees e ON e.id=a.employee_id
                    WHERE a.id=sms_deliveries.appeal_id AND e.approved=1 AND sms_deliveries.phone IN (e.phone,e.phone_extra))""",
                (now, appeal_id, phone, now))
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

    def claim(self, appeal_id, channel, now):
        with self.connect() as conn:
            result = conn.execute("""UPDATE deliveries SET status='sending',attempts=attempts+1,updated_at=?
                WHERE appeal_id=? AND channel=? AND status IN ('pending','retry','blocked') AND next_attempt<=?""",
                (now, appeal_id, channel, now))
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
            if conn.execute("SELECT 1 FROM deliveries WHERE appeal_id=? AND (attempts>0 OR status IN ('sent','sending','unknown'))", (appeal_id,)).fetchone():
                raise ValueError("Yuborish boshlangan; vaqtni o'zgartirib bo'lmaydi.")
            conn.execute("UPDATE appeals SET due_at=?,schedule_source='explicit' WHERE id=?", (due_at, appeal_id))

    def assign(self, appeal_id, reference):
        with self.connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            employee_id = self.resolve_employee(conn, reference)
            if conn.execute("SELECT 1 FROM deliveries WHERE appeal_id=? AND (attempts>0 OR status IN ('sent','sending','unknown'))", (appeal_id,)).fetchone():
                raise ValueError("Yuborish boshlangan; xodimni almashtirib bo'lmaydi.")
            if not conn.execute("UPDATE appeals SET employee_id=?,employee_hint=? WHERE id=?", (employee_id, reference, appeal_id)).rowcount:
                raise ValueError("Murojaat topilmadi.")

    def retry(self, appeal_id, channel):
        with self.connect() as conn:
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

    def status_text(self):
        with self.connect() as conn:
            total = conn.execute("SELECT COUNT(*) FROM appeals").fetchone()[0]
            unscheduled = conn.execute("SELECT COUNT(*) FROM appeals WHERE due_at IS NULL").fetchone()[0]
            unassigned = conn.execute("SELECT COUNT(*) FROM appeals WHERE employee_id IS NULL").fetchone()[0]
            unapproved = conn.execute("SELECT COUNT(*) FROM employees WHERE approved=0").fetchone()[0]
            rows = conn.execute("SELECT channel,status,COUNT(*) AS n FROM deliveries GROUP BY channel,status").fetchall()
        lines = [f"Jami murojaat: {total}", f"Vaqt belgilanmagan: {unscheduled}", f"Xodim biriktirilmagan: {unassigned}",
                 f"Tasdiqlanmagan xodimlar: {unapproved}"]
        mode = self.default_schedule()
        lines.insert(0, "Avtomatik yuborish: " + ("yuklangach darhol" if mode == "darhol" else f"har kuni {mode}" if mode else "sozlanmagan — /jadval ni tanlang"))
        for channel in ("telegram", "sms"):
            lines.append("\n" + channel.upper() + ":")
            counts = {r["status"]: r["n"] for r in rows if r["channel"] == channel}
            lines.extend(f"{label}: {counts.get(key, 0)}" for key, label in STATUSES.items())
        return "\n".join(lines)

    def list_appeals(self, timezone, today=False):
        now = datetime.now(timezone)
        start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        with self.connect() as conn:
            query = """SELECT a.id,a.due_at,COALESCE(e.name,NULLIF(a.employee_hint,'')) AS employee,
                e.approved AS employee_approved,
                t.status AS telegram,s.status AS sms FROM appeals a
                LEFT JOIN employees e ON a.employee_id=e.id
                JOIN deliveries t ON t.appeal_id=a.id AND t.channel='telegram'
                JOIN deliveries s ON s.appeal_id=a.id AND s.channel='sms'"""
            params = ()
            if today:
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
            return conn.execute("UPDATE notices SET status='sending',attempts=attempts+1 WHERE notice_key=? AND status IN ('pending','retry')", (key,)).rowcount == 1

    def finish_notice(self, key, status, delay=0):
        with self.connect() as conn:
            conn.execute("UPDATE notices SET status=?,next_attempt=? WHERE notice_key=?", (status, time.time()+delay, key))
