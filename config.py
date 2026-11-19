"""Environment configuration; secrets are never logged."""
import os
from dataclasses import dataclass, field
from pathlib import Path
from zoneinfo import ZoneInfo

from dotenv import dotenv_values, load_dotenv

ROOT = Path(__file__).resolve().parent
TIMEZONE = ZoneInfo("Asia/Tashkent")


@dataclass(frozen=True)
class Settings:
    token: str = field(repr=False)
    admin_ids: frozenset[int]
    timezone: ZoneInfo
    reminder_hour: int = 9
    database: Path = ROOT / "data" / "bot.sqlite3"
    jobs_database: Path = ROOT / "data" / "jobs.sqlite3"
    sms_target: str = "both"
    demo: bool = False

    def __post_init__(self):
        object.__setattr__(self, "timezone", TIMEZONE)


def load_settings() -> Settings:
    file_settings = dotenv_values(ROOT / ".env", interpolate=False)
    load_dotenv(ROOT / ".env", override=False)
    token = os.getenv("BOT_TOKEN", "").strip()
    if not token:
        raise ValueError(".env faylida BOT_TOKEN belgilanmagan.")
    try:
        admins = frozenset(int(value.strip()) for value in (file_settings.get("ADMIN_IDS") or "").split(",") if value.strip())
        if any(value <= 0 for value in admins):
            raise ValueError
    except ValueError:
        raise ValueError("ADMIN_IDS musbat Telegram ID laridan iborat bo'lsin; vergul bilan ajrating.") from None
    try:
        hour = int(os.getenv("ADMIN_REMINDER_HOUR", "9"))
        if not 0 <= hour <= 23:
            raise ValueError
    except Exception:
        raise ValueError("ADMIN_REMINDER_HOUR 0–23 orasida bo'lsin.") from None
    sms_target = os.getenv("SMS_TARGET", "both").strip().lower()
    if sms_target not in ("primary", "both"):
        raise ValueError("SMS_TARGET primary yoki both bo'lsin.")
    return Settings(token, admins, TIMEZONE, hour, sms_target=sms_target)


def load_demo_settings() -> Settings:
    token = os.getenv("DEMO_BOT_TOKEN", "").strip()
    if not token or token == os.getenv("BOT_TOKEN", "").strip():
        raise ValueError("DEMO_BOT_TOKEN uchun alohida Telegram bot tokenini kiriting.")
    try:
        admin_id = int(os.getenv("DEMO_ADMIN_ID", ""))
        if admin_id <= 0:
            raise ValueError
    except ValueError:
        raise ValueError("DEMO_ADMIN_ID musbat Telegram ID bo‘lsin.") from None
    directory = ROOT / "data" / "demo"
    return Settings(token, frozenset({admin_id}), TIMEZONE, database=directory / "bot.sqlite3",
                    jobs_database=directory / "jobs.sqlite3", demo=True)
