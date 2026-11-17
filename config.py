"""Environment configuration; secrets are never logged."""
import os
from dataclasses import dataclass
from pathlib import Path
from zoneinfo import ZoneInfo

from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent


@dataclass(frozen=True)
class Settings:
    token: str
    admin_ids: frozenset[int]
    timezone: ZoneInfo
    reminder_hour: int = 9
    database: Path = ROOT / "data" / "bot.sqlite3"
    jobs_database: Path = ROOT / "data" / "jobs.sqlite3"
    sms_target: str = "both"


def load_settings() -> Settings:
    load_dotenv(ROOT / ".env", override=False)
    token = os.getenv("BOT_TOKEN", "").strip()
    if not token:
        raise ValueError(".env faylida BOT_TOKEN belgilanmagan.")
    try:
        admins = frozenset(int(value.strip()) for value in os.getenv("ADMIN_IDS", "").split(",") if value.strip())
        if any(value <= 0 for value in admins):
            raise ValueError
    except ValueError:
        raise ValueError("ADMIN_IDS musbat Telegram ID laridan iborat bo'lsin; vergul bilan ajrating.") from None
    try:
        tz = ZoneInfo(os.getenv("TIMEZONE", "Asia/Tashkent"))
        hour = int(os.getenv("ADMIN_REMINDER_HOUR", "9"))
        if not 0 <= hour <= 23:
            raise ValueError
    except Exception:
        raise ValueError("TIMEZONE yoki ADMIN_REMINDER_HOUR noto'g'ri.") from None
    sms_target = os.getenv("SMS_TARGET", "both").strip().lower()
    if sms_target not in ("primary", "both"):
        raise ValueError("SMS_TARGET primary yoki both bo'lsin.")
    return Settings(token, admins, tz, hour, sms_target=sms_target)
