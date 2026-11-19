"""Persistent APScheduler worker backed by a transactional SQLite outbox."""
import asyncio
import logging
import time
from functools import partial
from datetime import datetime, timedelta
from pathlib import Path

from apscheduler.jobstores.sqlalchemy import SQLAlchemyJobStore
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from telegram.error import BadRequest, Forbidden, NetworkError, RetryAfter, TelegramError

from db import STATUSES, Database, assignment_problem
from excel_utils import appeal_contact, appeals_workbook
import sms

logger = logging.getLogger(__name__)
_service = None
SEND_TIMEOUT = 60


def pages(text, size=3500):
    while len(text) > size:
        end = text.rfind("\n", 0, size)
        if end < 1:
            end = size
        yield text[:end]
        text = text[end:].lstrip("\n")
    if text:
        yield text


def appeal_list_text(rows, timezone, title):
    lines = [title]
    for row in rows:
        when = datetime.fromtimestamp(row["due_at"], timezone).strftime("%d.%m %H:%M") if row["due_at"] else "vaqt belgilanmagan"
        lines.append(f"{row['id']} | {when} | {row['employee'] or 'xodim kerak'}\n"
                     f"Xodim: {'tasdiqlangan' if row.get('employee_approved') else 'biriktirish yoki tasdiq kerak'}\n"
                     f"Holat: {STATUSES[row['telegram']]}")
    if not rows:
        lines.append("Murojaatlar yo'q.")
    return "\n\n".join(lines)


def telegram_text_parts(appeal_id, client_name, client_phone, body, limit=3500, *, contact_label="Mijoz"):
    contact_line = f"{contact_label}: {client_name}\n" if client_name else ""
    value = f"Murojaat #{appeal_id}\n{contact_line}Telefon: {client_phone}\n\n{body}"
    parts, remaining = [], value
    while remaining:
        size = units = 0
        for index, char in enumerate(remaining):
            count = len(char.encode("utf-16-le")) // 2
            if units + count > limit:
                break
            units += count
            size = index + 1
        if not size:
            raise ValueError("Xabarni Telegram matniga ajratib bo'lmadi.")
        if size < len(remaining):
            boundary = remaining.rfind("\n", size // 2, size)
            if boundary > 0:
                size = boundary + 1
        parts.append(remaining[:size].rstrip())
        remaining = remaining[size:].lstrip()
    if len(parts) == 1:
        return parts
    total = len(parts)
    return [f"Murojaat #{appeal_id} ({index}/{total})\n{part}" for index, part in enumerate(parts, 1)]


async def dispatch_tick():
    # Stored jobs only contain this importable function, never a Bot instance.
    if _service:
        try:
            await _service.tick()
        except Exception as exc:
            logger.error("Worker tick failed error=%s", type(exc).__name__)


class DeliveryService:
    def __init__(self, bot, database, settings, sms_sender=None):
        self.bot = bot
        self.db = database
        self.settings = settings
        self.sms_sender = sms_sender or sms.send_sms
        self.lock = asyncio.Lock()
        self.scheduler = None
        self._backup_key = None

    def start(self):
        global _service
        _service = self
        self.db.recover()
        self.scheduler = AsyncIOScheduler(
            jobstores={"default": SQLAlchemyJobStore(url="sqlite:///" + self.settings.jobs_database.as_posix())},
            timezone=self.settings.timezone,
            job_defaults={"coalesce": True, "max_instances": 1, "misfire_grace_time": None},
        )
        self.scheduler.start(paused=True)
        self.scheduler.add_job(dispatch_tick, "interval", seconds=15, id="delivery-worker", replace_existing=True,
                               next_run_time=datetime.now(self.settings.timezone))
        self.scheduler.resume()

    async def stop(self):
        if self.scheduler:
            self.scheduler.pause()
            async with self.lock:
                self.scheduler.shutdown(wait=False)
            await asyncio.sleep(0)

    def notice(self, key, text):
        for index, part in enumerate(pages(text)):
            self.db.enqueue_notice(f"{key}:{index}", self.settings.admin_ids, part)

    async def tick(self):
        async with self.lock:
            now = time.time()
            self.db.expire_sending(now)
            self.db.apply_default_schedule(self.settings.timezone, now)
            self.db.assign_due(now, self.settings.timezone)
            for delivery in self.db.ready_deliveries(now, self.settings.sms_target):
                await self.deliver(delivery)
            self.prepare_reminders()
            await self.send_notices()
            try:
                await asyncio.to_thread(self.backup_databases, now)
            except Exception as exc:
                logger.error('Database backup failed error=%s', type(exc).__name__)
                self.notice('backup-failed:' + datetime.fromtimestamp(now, self.settings.timezone).strftime('%Y-%m-%d-%H'),
                            "Avtomatik backup bajarilmadi. Disk joyi va yozish ruxsatlarini tekshiring.")

    def backup_databases(self, now):
        day = datetime.fromtimestamp(now, self.settings.timezone).strftime('%Y-%m-%d')
        sources = [Path(path) for path in (self.settings.database, self.settings.jobs_database) if Path(path).exists()]
        key = (day, tuple(str(path) for path in sources))
        if self._backup_key == key:
            return
        directory = self.settings.database.parent / 'backups'
        for index, source in enumerate(sources):
            target = directory / f'{index}-{source.stem}-{day}.sqlite3'
            if not target.exists():
                Database.create_backup(source, target)
        self._backup_key = key

    async def deliver(self, item):
        appeal_id, channel = item["appeal_id"], item["channel"]
        destination = '/menu → Sozlamalar → SMS sozlamalari → SMSda bormaganlar' if channel == 'sms' else '/menu → Holat → E’tibor kerak'
        phone = item.get('target_phone')
        now = time.time()
        if phone:
            claim = self.db.claim_sms(appeal_id, phone, now)
        else:
            claim = self.db.claim(appeal_id, channel, now, delivery=item if channel == "telegram" else None)
        if not claim:
            return
        finish = partial(self.db.finish_sms, appeal_id, phone) if phone else partial(self.db.finish, appeal_id, channel)
        attempt = item["attempts"] + 1
        try:
            if channel == "telegram":
                if not item["telegram_id"]:
                    # No network call was made, so waiting consumes no retry budget.
                    finish("retry", "Xodim botga /start yuborishi kerak", delay=60, refund=True)
                    self.notice(f"unregistered:{appeal_id}", f"#{appeal_id}: xodim botga /start yubormagan. /menu → Xodimlar bo‘limini tekshiring.")
                    return
                label, name = appeal_contact(item)
                parts = telegram_text_parts(appeal_id, name, item["client_phone"], item["body"], contact_label=label)
                message = None
                for index in range(item["part_index"], len(parts)):
                    message = await asyncio.wait_for(self.bot.send_message(chat_id=item["telegram_id"], text=parts[index]), SEND_TIMEOUT)
                    self.db.advance_part(appeal_id, index + 1, message.message_id)
                if message is None:
                    raise ValueError("Yuborilmagan xabar qismi qolmadi.")
                message_id = message.message_id
            else:
                text = sms.task_notification(self.bot.username, item["appeal_number"])
                message_id = await asyncio.wait_for(self.sms_sender(phone, text), SEND_TIMEOUT)
            finish("sent", message_id=message_id)
        except sms.SMSNotConfigured as exc:
            finish("blocked", str(exc), delay=3600, refund=True)
        except sms.SMSRejectedError as exc:
            finish("failed", str(exc))
            self.notice(f"rejected:{appeal_id}:sms", f"#{appeal_id}: SMS rad etildi. {destination} bo‘limini tekshiring.")
        except asyncio.CancelledError:
            finish('unknown', 'Yuborish bekor qilindi; natijani tekshiring.')
            raise
        except (TimeoutError, NetworkError, sms.SMSDeliveryUnknown) as exc:
            if isinstance(exc, BadRequest):
                self.record_retry(appeal_id, channel, attempt, exc, finish)
                return
            # Telegram has no idempotency keys: retrying an uncertain HTTP outcome can duplicate delivery.
            finish("unknown", type(exc).__name__)
            logger.warning("Delivery outcome unknown channel=%s error=%s", channel, type(exc).__name__)
            self.notice(f"unknown:{appeal_id}:{channel}",
                        f"#{appeal_id}: {channel} natijasi noma'lum. Qabul qiluvchidan tekshirib, {destination} bo‘limida natijani belgilang.")
        except (TelegramError, sms.SMSRetryableError) as exc:
            self.record_retry(appeal_id, channel, attempt, exc, finish)
        except Exception as exc:
            finish("unknown", type(exc).__name__)
            logger.error("Unexpected delivery failure channel=%s error=%s", channel, type(exc).__name__)
            self.notice(f"unknown:{appeal_id}:{channel}", f"#{appeal_id}: {channel} yuborish natijasi noma'lum; {destination} bo‘limida tekshiring.")

    def record_retry(self, appeal_id, channel, attempt, exc, finish):
        delay = 30 * (2 ** (attempt - 1))
        if isinstance(exc, RetryAfter):
            value = exc.retry_after
            delay = max(delay, value.total_seconds() if isinstance(value, timedelta) else value)
        status = "failed" if attempt >= 4 else "retry"
        finish(status, type(exc).__name__, delay=delay)
        logger.warning("Delivery failed channel=%s attempt=%d error=%s", channel, attempt, type(exc).__name__)
        if status == "failed":
            destination = '/menu → Sozlamalar → SMS sozlamalari → SMSda bormaganlar' if channel == 'sms' else '/menu → Holat → E’tibor kerak'
            self.notice(f"failed:{appeal_id}:{channel}", f"#{appeal_id}: {channel} 1 urinish va 3 qayta urinishdan keyin yuborilmadi. {destination} bo‘limida ko‘ring.")

    def prepare_reminders(self):
        now = datetime.now(self.settings.timezone)
        rows = self.db.list_appeals(self.settings.timezone, today=True)
        if now.hour >= self.settings.reminder_hour:
            self.notice("daily:" + now.date().isoformat(), appeal_list_text(rows, self.settings.timezone, "Bugungi yuborishlar"))
        all_rows = self.db.list_appeals(self.settings.timezone)
        reasons = {'unassigned': 'Mos xodim topilmadi. Xodim biriktiring yoki smenani tekshiring.',
                   'unconnected': 'Biriktirilgan xodim botga ulanmagan. Ulanishini tasdiqlang.'}
        for row in all_rows:
            if row['telegram'] == 'sent':
                continue
            problem = assignment_problem(row)
            if problem:
                self.notice(f"assignment:{row['id']}:{problem}:" + now.strftime('%Y-%m-%d-%H'),
                            f"⚠️ Murojaat #{row['id']}: {reasons[problem]} /menu → Holat.")
        pending = [row for row in all_rows if row['completed_at'] is None and row['telegram'] != 'sent'
                   and not assignment_problem(row) and (row['due_at'] is None or row['due_at'] <= now.timestamp())]
        if pending:
            self.notice("attention:" + now.strftime("%Y-%m-%d-%H"),
                        appeal_list_text(pending, self.settings.timezone, "Eslatma: vaqt kutilayotgan yoki yakunlanmagan murojaatlar"))

    async def send_notices(self):
        for notice in self.db.ready_notices(time.time()):
            key = notice["notice_key"]
            if not self.db.claim_notice(key):
                continue
            try:
                await asyncio.wait_for(self.bot.send_message(chat_id=notice["admin_id"], text=notice["text"]), SEND_TIMEOUT)
                self.db.finish_notice(key, "sent")
            except (TimeoutError, NetworkError) as exc:
                if isinstance(exc, BadRequest):
                    self.db.finish_notice(key, "failed" if notice["attempts"] >= 3 else "retry", 60)
                else:
                    self.db.finish_notice(key, "unknown")
                    logger.warning("Admin reminder outcome unknown")
            except Exception as exc:
                status = "failed" if notice["attempts"] >= 3 else "retry"
                delay = 60 * 2 ** notice["attempts"]
                if isinstance(exc, RetryAfter):
                    value = exc.retry_after
                    delay = max(delay, value.total_seconds() if isinstance(value, timedelta) else value)
                self.db.finish_notice(key, status, delay)
                logger.warning("Admin reminder failed error=%s", type(exc).__name__)
