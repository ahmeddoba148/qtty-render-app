import os
import json
import re
import html
import imaplib
import email
import datetime
import unicodedata
import requests

from zoneinfo import ZoneInfo
from email.header import decode_header

import gspread
from google.oauth2.service_account import Credentials


# =========================================================
# CONFIG
# =========================================================

GMAIL_EMAIL = os.environ["GMAIL_EMAIL"]
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]

GOOGLE_CREDS_JSON = os.environ["GOOGLE_CREDS_JSON"]

GOOGLE_SHEET_NAME = (
    os.environ.get("GOOGLE_SHEET_NAME", "mail_tracker").strip()
    or "mail_tracker"
)

TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
TELEGRAM_CHAT_ID = os.environ["TELEGRAM_CHAT_ID"]

RENDER_SITE_URL = os.environ.get(
    "RENDER_SITE_URL",
    ""
).strip()


# عدد الأيام التي نرجع للخلف فيها كشبكة أمان.
# لا يسبب تكرار التنبيه لأن Message-ID يتم حفظه.
LOOKBACK_DAYS = 3


# =========================================================
# TIME
# =========================================================

def cairo_now():
    """
    الوقت الحالي في القاهرة.
    """
    return datetime.datetime.now(
        ZoneInfo("Africa/Cairo")
    )


# =========================================================
# EMAIL TEXT DECODING
# =========================================================

def decode_mime_text(value):
    """
    يفك ترميز Subject وغيره حتى لو كان MIME encoded
    أو استخدم UTF-8 / Latin / Windows encoding.
    """

    if not value:
        return ""

    parts = decode_header(value)

    result = []

    for part, encoding in parts:

        if isinstance(part, bytes):

            candidates = [
                encoding,
                "utf-8",
                "windows-1256",
                "latin-1",
            ]

            decoded = None

            for enc in candidates:

                if not enc:
                    continue

                try:
                    decoded = part.decode(
                        enc,
                        errors="strict"
                    )
                    break

                except (
                    LookupError,
                    UnicodeDecodeError
                ):
                    pass

            if decoded is None:
                decoded = part.decode(
                    "utf-8",
                    errors="replace"
                )

            result.append(decoded)

        else:
            result.append(str(part))

    return "".join(result).strip()


# =========================================================
# SUBJECT MATCHING
# =========================================================

def normalize_subject(subject):
    """
    توحيد شكل العنوان قبل المقارنة.

    مثال:
        QTTY RECAP
        qtty recap
        Qtty Recap

    كلهم يتحولوا لشكل موحد.
    """

    text = unicodedata.normalize(
        "NFKC",
        subject or ""
    )

    text = text.casefold()

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def is_qtty_recap_subject(subject):
    """
    التعرف المرن على رسائل Qtty Recap.

    يغطي مثلًا:

    Qtty Recap

    QTTY RECAP

    qtty recap

    Qty Recap

    Qtty-Recap

    Qtty_Recap

    QttyRecap

    Q T T Y Recap

    Q.T.T.Y Recap

    Qtty Re-Cap

    RE: Qtty Recap

    Re: Qtty Recap for Nabil Factory

    FW: Qtty Recap

    FWD: Qtty Recap

    [External] Qtty Recap

    Qtty Recap - Nabil Factory

    Qtty Recap for Nabil Factory

    Recap Qtty

    Quantity Recap

    وبالتالي لا نعتمد على الاسم الكامل حرفيًا.
    """

    normalized = normalize_subject(
        subject
    )

    if not normalized:
        return False

    # ---------------------------------------------
    # إزالة كل المسافات والعلامات للمقارنة المرنة
    #
    # مثال:
    #
    # Qtty-Recap
    # Qtty_Recap
    # Qtty Recap
    # QttyRecap
    #
    # كلهم يتحولوا تقريبًا إلى:
    #
    # qttyrecap
    # ---------------------------------------------

    compact = re.sub(
        r"[^a-z0-9]+",
        "",
        normalized
    )

    compact_patterns = (
        "qttyrecap",
        "qtyrecap",
        "quantityrecap",

        # لو الترتيب انعكس لأي سبب
        "recapqtty",
        "recapqty",
        "recapquantity",
    )

    if any(
        pattern in compact
        for pattern in compact_patterns
    ):
        return True

    # ---------------------------------------------
    # تغطية وجود علامات بين الأحرف نفسها
    #
    # مثال:
    #
    # Q.T.T.Y Recap
    # Q T T Y Recap
    # Q-T-T-Y Re-Cap
    # ---------------------------------------------

    qtty_pattern = (
        r"q[\W_]*"
        r"t(?:[\W_]*t)?"
        r"[\W_]*y"
    )

    recap_pattern = (
        r"r[\W_]*"
        r"e[\W_]*"
        r"c[\W_]*"
        r"a[\W_]*"
        r"p"
    )

    if (
        re.search(
            qtty_pattern,
            normalized
        )
        and
        re.search(
            recap_pattern,
            normalized
        )
    ):
        return True

    # ---------------------------------------------
    # Quantity Recap
    # ---------------------------------------------

    quantity_pattern = (
        r"q[\W_]*"
        r"u[\W_]*"
        r"a[\W_]*"
        r"n[\W_]*"
        r"t[\W_]*"
        r"i[\W_]*"
        r"t[\W_]*"
        r"y"
    )

    if (
        re.search(
            quantity_pattern,
            normalized
        )
        and
        re.search(
            recap_pattern,
            normalized
        )
    ):
        return True

    return False


# =========================================================
# GOOGLE SHEETS
# =========================================================

def get_google_book():

    creds_dict = json.loads(
        GOOGLE_CREDS_JSON
    )

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = (
        Credentials
        .from_service_account_info(
            creds_dict,
            scopes=scopes
        )
    )

    client = gspread.authorize(
        creds
    )

    return client.open(
        GOOGLE_SHEET_NAME
    )


def get_or_create_notifications_sheet(book):
    """
    إنشاء شيت notifications إذا لم يكن موجودًا.
    """

    try:

        ws = book.worksheet(
            "notifications"
        )

    except gspread.WorksheetNotFound:

        ws = book.add_worksheet(
            title="notifications",
            rows=1000,
            cols=4
        )

        ws.update(
            range_name="A1:D1",
            values=[
                [
                    "message_id",
                    "subject",
                    "notified_at",
                    "date_key"
                ]
            ]
        )

    return ws


def get_notified_message_ids(ws):
    """
    قراءة كل Message-ID التي سبق إرسال Telegram عنها.
    """

    values = ws.col_values(1)

    return {
        str(value).strip()

        for value in values[1:]

        if str(value).strip()
    }


# =========================================================
# TELEGRAM
# =========================================================

def send_telegram(text):

    url = (
        "https://api.telegram.org/"
        f"bot{TELEGRAM_BOT_TOKEN}/"
        "sendMessage"
    )

    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": True,
    }

    response = requests.post(
        url,
        json=payload,
        timeout=30
    )

    if not response.ok:

        print(
            "Telegram HTTP Error:"
        )

        print(
            response.status_code
        )

        print(
            response.text
        )

    response.raise_for_status()

    print(
        "Telegram notification sent successfully."
    )


# =========================================================
# EMAIL MESSAGE ID
# =========================================================

def get_message_real_id(
    msg,
    fallback_id
):
    """
    Message-ID هو أفضل طريقة لمنع تكرار نفس التنبيه.

    في حالة نادرة جدًا إذا الرسالة ليس بها Message-ID
    نستخدم رقم IMAP كبديل.
    """

    real_id = (
        msg.get("Message-ID")
        or ""
    ).strip()

    if real_id:
        return real_id

    return str(
        fallback_id
    )


# =========================================================
# MAIN
# =========================================================

def main():

    now = cairo_now()

    date_key = now.strftime(
        "%Y-%m-%d"
    )

    print("=" * 70)

    print(
        "Qtty Recap Alert"
    )

    print(
        "Cairo time:",
        now.strftime(
            "%Y-%m-%d %H:%M:%S %Z"
        )
    )

    print("=" * 70)

    # =====================================================
    # GOOGLE SHEETS
    # =====================================================

    book = get_google_book()

    notify_sheet = (
        get_or_create_notifications_sheet(
            book
        )
    )

    notified_ids = (
        get_notified_message_ids(
            notify_sheet
        )
    )

    print(
        "Previously notified message IDs:",
        len(notified_ids)
    )

    # =====================================================
    # GMAIL
    # =====================================================

    mail = None

    try:

        mail = imaplib.IMAP4_SSL(
            "imap.gmail.com"
        )

        print(
            "Connecting to Gmail..."
        )

        mail.login(
            GMAIL_EMAIL,
            GMAIL_APP_PASSWORD
        )

        print(
            "Gmail login successful."
        )

        # readonly=True
        #
        # البرنامج لن يغيّر حالة الرسالة
        # ولن يجعلها مقروءة.
        status, _ = mail.select(
            "INBOX",
            readonly=True
        )

        if status != "OK":
            raise RuntimeError(
                "Could not select Gmail INBOX."
            )

        # =================================================
        # SEARCH
        # =================================================

        since_date = (
            now
            - datetime.timedelta(
                days=LOOKBACK_DAYS
            )
        ).strftime(
            "%d-%b-%Y"
        )

        print(
            "Searching inbox emails since:",
            since_date
        )

        # =================================================
        # مهم جدًا
        #
        # لا نكتب:
        #
        # SUBJECT "Qtty Recap"
        #
        # داخل Gmail Search.
        #
        # لأننا نريد تغطية أي تغيير معقول
        # في كتابة الاسم.
        #
        # لذلك نأخذ رسائل الأيام الأخيرة
        # ثم Python نفسه يقرر هل Subject
        # هو Qtty Recap أم لا.
        # =================================================

        status, messages = mail.search(
            None,
            f'(SINCE "{since_date}")'
        )

        if status != "OK":

            raise RuntimeError(
                "Gmail IMAP search failed."
            )

        if (
            not messages
            or
            not messages[0]
        ):

            print(
                "No recent inbox emails found."
            )

            return

        message_numbers = (
            messages[0].split()
        )

        print(
            "Recent inbox messages found:",
            len(message_numbers)
        )

        # =================================================
        # CHECK EMAILS
        # =================================================

        new_items = []

        # reversed =
        # نبدأ بأحدث رسالة أولًا.
        for num in reversed(
            message_numbers
        ):

            status, msg_data = mail.fetch(

                num,

                (
                    "BODY.PEEK[HEADER.FIELDS "
                    "(MESSAGE-ID SUBJECT DATE FROM)]"
                )
            )

            if (
                status != "OK"
                or
                not msg_data
            ):

                print(
                    "Could not fetch header for:",
                    num
                )

                continue

            raw_header = None

            for item in msg_data:

                if (
                    isinstance(
                        item,
                        tuple
                    )
                    and
                    len(item) >= 2
                ):

                    raw_header = item[1]

                    break

            if not raw_header:
                continue

            msg = email.message_from_bytes(
                raw_header
            )

            # =============================================
            # SUBJECT
            # =============================================

            subject = decode_mime_text(
                msg.get(
                    "Subject",
                    ""
                )
            )

            # =============================================
            # هل هي Qtty Recap؟
            # =============================================

            if not is_qtty_recap_subject(
                subject
            ):

                continue

            print(
                "MATCHED SUBJECT:",
                subject
                or
                "(no subject)"
            )

            # =============================================
            # MESSAGE ID
            # =============================================

            message_id = (
                get_message_real_id(
                    msg,
                    num.decode(
                        errors="replace"
                    )
                )
            )

            # =============================================
            # منع تكرار Telegram
            # =============================================

            if message_id in notified_ids:

                print(
                    "Already notified -> skipped."
                )

                continue

            print(
                "NEW Qtty Recap -> "
                "queued for Telegram."
            )

            new_items.append(
                (
                    message_id,
                    subject
                    or
                    "(بدون عنوان)"
                )
            )

        # =================================================
        # NO NEW EMAIL
        # =================================================

        if not new_items:

            print(
                "No new Qtty Recap emails "
                "requiring notification."
            )

            return

        # =================================================
        # CREATE TELEGRAM MESSAGE
        # =================================================

        lines = [

            "📩 <b>Qtty Recap جديد وصل</b>",

            "",

            (
                "عدد الرسائل الجديدة: "
                f"<b>{len(new_items)}</b>"
            ),
        ]

        # عرض أسماء الرسائل
        # بحد أقصى 10 في نفس Telegram.
        for _, subject in new_items[:10]:

            safe_subject = html.escape(
                subject
            )

            lines.append(
                f"• {safe_subject}"
            )

        if len(new_items) > 10:

            lines.append(

                "• ... و "

                f"{len(new_items) - 10}"

                " رسالة أخرى"
            )

        # =================================================
        # RENDER SITE
        # =================================================

        if RENDER_SITE_URL:

            safe_url = html.escape(
                RENDER_SITE_URL
            )

            lines.extend(
                [
                    "",
                    "⬇️ <b>افتح الموقع للتحميل:</b>",
                    safe_url,
                ]
            )

        telegram_text = "\n".join(
            lines
        )

        # =================================================
        # SEND TELEGRAM
        # =================================================

        send_telegram(
            telegram_text
        )

        # =================================================
        # SAVE MESSAGE IDs
        # =================================================

        notified_at = now.strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        rows = [

            [
                message_id,
                subject,
                notified_at,
                date_key
            ]

            for (
                message_id,
                subject
            )
            in new_items
        ]

        notify_sheet.append_rows(
            rows,
            value_input_option="RAW"
        )

        print(
            "Saved notification rows:",
            len(rows)
        )

        print(
            "DONE - Telegram alert sent for",
            len(new_items),
            "new email(s)."
        )

    # =====================================================
    # CLOSE GMAIL
    # =====================================================

    finally:

        if mail is not None:

            try:
                mail.logout()

            except Exception:
                pass


# =========================================================
# START
# =========================================================

if __name__ == "__main__":
    main()
