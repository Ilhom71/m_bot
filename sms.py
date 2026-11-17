"""Replace this adapter once the provider contract and templates are supplied."""
import os

from dotenv import load_dotenv
from config import ROOT


class SMSNotConfigured(Exception):
    """No live SMS integration is installed."""


class SMSDeliveryUnknown(Exception):
    """The provider may have accepted the message; do not retry blindly."""


class SMSRetryableError(Exception):
    """The provider explicitly rejected the attempt without accepting delivery."""


async def send_sms(phone: str, text: str) -> str:
    """Return a provider message ID only after confirmed API acceptance.

    Never report success from a placeholder. A future provider should support
    idempotency keys or reconciliation for requests with uncertain outcomes.
    """
    load_dotenv(ROOT / ".env", override=False)
    api_key = os.getenv("SMS_API_KEY", "")
    api_url = os.getenv("SMS_API_URL", "")
    if not api_key or not api_url:
        raise SMSNotConfigured("SMS API sozlanmagan.")
    raise SMSNotConfigured("SMS provayder adapteri hali ulanmagan.")
