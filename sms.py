"""Short task notifications and authenticated TextUP SMS delivery."""
import asyncio
import os

import httpx
from telegram.helpers import create_deep_linked_url

from excel_utils import phone_number


TASK_NOTIFICATION_TEMPLATE = "Sizga yangi vazifa biriktirildi. Batafsil: {TELEGRAM_HAVOLA}"
LOGIN_URL = "https://api-auth.textup.uz/v1/login"
SEND_URL = "https://sms-api.textup.uz/v1/send"
_textup_client = None


def task_notification(bot_username: str, appeal_number: int) -> str:
    link = create_deep_linked_url(bot_username, f"task_{appeal_number}")
    return TASK_NOTIFICATION_TEMPLATE.format(TELEGRAM_HAVOLA=link)


class SMSNotConfigured(Exception):
    """Credentials are missing or TextUP authentication was rejected."""


class SMSDeliveryUnknown(Exception):
    """The provider may have accepted the message; do not retry blindly."""


class SMSRetryableError(Exception):
    """No delivery was accepted, so a later attempt is safe."""


class SMSRejectedError(Exception):
    """The provider rejected the SMS; automatic retries will not fix it."""


class TextUPClient:
    def __init__(self, email, password, template_id="", nickname_id="", *, transport=None):
        self.configuration = (email, password, template_id, nickname_id)
        self._transport = transport
        self._access_token = None
        self._user_id = None
        self._auth_lock = asyncio.Lock()

    async def _authenticate(self, client, rejected_token=None):
        async with self._auth_lock:
            if self._access_token and self._access_token != rejected_token:
                return self._access_token, self._user_id
            self._access_token = self._user_id = None
            email, password, _, _ = self.configuration
            try:
                response = await client.post(LOGIN_URL, json={"email": email, "password": password})
            except httpx.RequestError:
                raise SMSRetryableError("TextUP bilan kirish aloqasi o'rnatilmadi.") from None
            if 400 <= response.status_code < 500 and response.status_code not in (408, 429):
                raise SMSNotConfigured("TextUP kirishni rad etdi. TEXTUP_EMAIL va TEXTUP_PASSWORD ni tekshiring.")
            if not 200 <= response.status_code < 300:
                raise SMSRetryableError("TextUP kirish xizmati vaqtincha javob bermadi.")
            try:
                data = response.json()
                token, user_id = data["accessToken"], data["user"]["id"]
                if not isinstance(token, str) or not token.strip() or not isinstance(user_id, str) or not user_id.strip():
                    raise ValueError
            except (ValueError, KeyError, TypeError):
                raise SMSRetryableError("TextUP kirish javobi noto'g'ri.") from None
            self._access_token, self._user_id = token, user_id
            return token, user_id

    async def send(self, phone: str, text: str) -> str:
        try:
            recipient = phone_number(phone)
        except ValueError:
            raise SMSRejectedError("SMS qabul qiluvchi telefoni noto'g'ri.") from None
        if not isinstance(text, str) or not text.strip():
            raise SMSRejectedError("SMS matni bo'sh bo'lmasin.")
        async with httpx.AsyncClient(timeout=httpx.Timeout(20, connect=5), follow_redirects=False,
                                     transport=self._transport) as client:
            token, user_id = await self._authenticate(client)
            for attempt in range(2):
                payload = {"message": text, "userId": user_id, "name": "Task notification",
                           "recipients": [recipient]}
                _, _, template_id, nickname_id = self.configuration
                if template_id:
                    payload["templateId"] = template_id
                if nickname_id:
                    payload["nicknameId"] = nickname_id
                try:
                    response = await client.post(SEND_URL, json=payload,
                                                 headers={"Authorization": "Bearer " + token})
                except (httpx.ConnectError, httpx.ConnectTimeout, httpx.PoolTimeout):
                    raise SMSRetryableError("TextUP SMS xizmatiga ulanib bo'lmadi.") from None
                except httpx.RequestError:
                    # A write/read failure can happen after TextUP accepts the SMS.
                    raise SMSDeliveryUnknown("TextUP SMS yuborish natijasi noma'lum.") from None
                if response.status_code == 401:
                    if attempt == 0:
                        # The supplied API has no refresh endpoint; log in again once.
                        token, user_id = await self._authenticate(client, rejected_token=token)
                        continue
                    raise SMSNotConfigured("TextUP yangi tokenni ham rad etdi. Hisob ruxsatlarini tekshiring.")
                if response.status_code == 429:
                    raise SMSRetryableError("TextUP so'rovlar chegarasi. Keyinroq qayta uriniladi.")
                if 400 <= response.status_code < 500 and response.status_code != 408:
                    raise SMSRejectedError(f"TextUP SMS so'rovini rad etdi (HTTP {response.status_code}).")
                if not 200 <= response.status_code < 300:
                    raise SMSDeliveryUnknown(f"TextUP SMS natijasi tasdiqlanmadi (HTTP {response.status_code}).")
                try:
                    sms_id = response.json()["smsId"]
                    if not isinstance(sms_id, str) or not sms_id.strip():
                        raise ValueError
                except (ValueError, KeyError, TypeError):
                    raise SMSDeliveryUnknown("TextUP javobida smsId topilmadi.") from None
                return sms_id


async def send_sms(phone: str, text: str) -> str:
    """Use settings loaded by main; return smsId only on confirmed API acceptance."""
    global _textup_client
    configuration = (os.getenv("TEXTUP_EMAIL", "").strip(), os.getenv("TEXTUP_PASSWORD", ""),
                     os.getenv("TEXTUP_TEMPLATE_ID", "").strip(), os.getenv("TEXTUP_NICKNAME_ID", "").strip())
    if not configuration[0] or not configuration[1]:
        raise SMSNotConfigured("TEXTUP_EMAIL yoki TEXTUP_PASSWORD sozlanmagan.")
    if _textup_client is None or _textup_client.configuration != configuration:
        _textup_client = TextUPClient(*configuration)
    return await _textup_client.send(phone, text)
