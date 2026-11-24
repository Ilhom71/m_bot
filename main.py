"""Telegram bot entry point. Run one polling process per token."""
import argparse
import asyncio
import errno
import logging
import os
import re
import sys
import warnings
from datetime import datetime
from getpass import GetPassWarning, getpass
from logging.handlers import RotatingFileHandler

from telegram import Bot, BotCommand, BotCommandScopeChat
from telegram.ext import Application

from config import ROOT, TIMEZONE, Settings, load_demo_settings, load_settings
from db import Database
from handlers.ui import register_handlers
from scheduler import DeliveryService
from sms import SMSNotConfigured

logger = logging.getLogger(__name__)


class RedactingFormatter(logging.Formatter):
    def formatTime(self, record, datefmt=None):
        moment = datetime.fromtimestamp(record.created, TIMEZONE)
        return moment.strftime(datefmt or '%Y-%m-%d %H:%M:%S%z')

    def format(self, record):
        value = super().format(record)
        for key, secret in os.environ.items():
            if secret and any(word in key.upper() for word in ('TOKEN', 'PASSWORD', 'API_KEY', 'SECRET')):
                value = value.replace(secret, '[REDACTED]')
        value = re.sub(r'(?i)(/bot|\b)([0-9]{5,}:[A-Za-z0-9_-]{10,})', r'\1[REDACTED]', value)
        value = re.sub(r'(?i)(Bearer\s+)[^\s\"\',;]+', r'\1[REDACTED]', value)
        value = re.sub(r'\beyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+', '[REDACTED]', value)
        value = re.sub(r'''(?ix)((?:password|accessToken|refreshToken|api[_-]?key|bot[_-]?token)["']?\s*[:=]\s*)
                          (?:"[^"\n]*"|'[^'\n]*'|[^\s,;}]+)''', r'\1[REDACTED]', value)
        return value


def configure_logging():
    directory = ROOT / "logs"
    directory.mkdir(exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s", handlers=[
        RotatingFileHandler(directory / "bot.log", maxBytes=2_000_000, backupCount=3, encoding="utf-8"),
        logging.StreamHandler(),
    ])
    for handler in logging.getLogger().handlers:
        handler.setFormatter(RedactingFormatter('%(asctime)s %(levelname)s %(name)s %(message)s'))
    # HTTP URLs contain bot tokens, so never enable request logging.
    for name in ("httpx", "httpcore", "telegram", "apscheduler"):
        logging.getLogger(name).setLevel(logging.CRITICAL)


class InstanceLock:
    """OS lock released automatically on crash; prevents duplicate workers."""
    def __init__(self, path=None):
        self.path = path if path is not None else ROOT / "data" / "bot.lock"

    def __enter__(self):
        path = self.path
        path.parent.mkdir(parents=True, exist_ok=True)
        self.handle = path.open("a+b")
        # Windows byte locks also deny reads. Acquire before any file I/O;
        # locking an empty file is supported, so no sentinel byte is needed.
        self.handle.seek(0)
        try:
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(self.handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            self.handle.close()
            if exc.errno in (errno.EACCES, errno.EAGAIN, errno.EDEADLK):
                raise RuntimeError("Botning boshqa nusxasi allaqachon ishlayapti. Avval o'sha nusxani to'xtating.") from None
            raise
        return self

    def __exit__(self, *args):
        self.handle.close()


async def post_init(application):
    settings = application.bot_data["settings"]
    public = [BotCommand("start", "Botni boshlash"), BotCommand("menu", "Bosh menyu"),
              BotCommand("bekor", "Kiritishni bekor qilish")]
    await application.bot.set_my_commands(public)
    for admin_id in settings.admin_ids:
        try:
            await application.bot.set_my_commands(public, scope=BotCommandScopeChat(admin_id))
        except Exception as exc:
            logger.warning("Admin menu unavailable error=%s", type(exc).__name__)
    application.bot_data["service"].start()
    logger.info("Bot started; configured_admins=%d", len(settings.admin_ids))
    if not settings.admin_ids:
        logger.warning("ADMIN_IDS is empty; admin features are disabled")


async def post_stop(application):
    await application.bot_data["service"].stop()


async def demo_sms_sender(phone, message):
    raise SMSNotConfigured("Namoyish rejimida haqiqiy SMS yuborilmaydi.")


def build_application(settings):
    application = (Application.builder().token(settings.token).concurrent_updates(False)
                   .post_init(post_init).post_stop(post_stop).build())
    database = Database(settings.database)
    application.bot_data.update(settings=settings, db=database,
                                service=DeliveryService(application.bot, database, settings,
                                                        sms_sender=demo_sms_sender if settings.demo else None))
    register_handlers(application)
    return application


async def check_connection(settings):
    async with Bot(settings.token) as bot:
        me = await bot.get_me()
        logger.info("Telegram connection OK; bot=@%s; configured_admins=%d", me.username, len(settings.admin_ids))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true", help="Check Telegram authentication without polling or sending messages")
    parser.add_argument("--demo", action="store_true", help="Run with a separate demo bot and database")
    parser.add_argument("--set-sms-password", action="store_true", help="Set the SMS control password in a local terminal")
    args = parser.parse_args()
    if args.set_sms_password:
        if args.demo or args.check or not sys.stdin.isatty():
            parser.error("SMS parolini alohida buyruq bilan interaktiv terminalda belgilang.")
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("error", GetPassWarning)
                password = getpass("SMS boshqaruvi paroli: ")
                confirmation = getpass("Parolni takrorlang: ")
            if password != confirmation:
                parser.error("Parollar bir xil bo‘lishi kerak.")
            Database(Settings.database).set_sms_password(password)
        except (EOFError, GetPassWarning, KeyboardInterrupt):
            parser.error("Parolni yashirin kiritish yakunlanmadi.")
        except ValueError as exc:
            parser.error(str(exc))
        print("SMS boshqaruvi paroli saqlandi.")
        return
    configure_logging()
    try:
        settings = load_demo_settings() if args.demo else load_settings()
        if args.check:
            asyncio.run(check_connection(settings))
        else:
            with InstanceLock(settings.database.parent / "bot.lock"):
                application = build_application(settings)
                application.run_polling(allowed_updates=["message", "callback_query"], drop_pending_updates=False, bootstrap_retries=3)
    except PermissionError as exc:
        logger.error("Faylga yoki tizim amaliga ruxsat yo'q: %s (errno=%s)", exc.filename or "yo'l ko'rsatilmagan", exc.errno)
        raise SystemExit(1) from None
    except (ValueError, RuntimeError) as exc:
        logger.error("%s", exc)
        raise SystemExit(1) from None
    except Exception as exc:
        logger.error("Bot stopped error=%s", type(exc).__name__)
        raise SystemExit(1) from None


if __name__ == "__main__":
    main()
