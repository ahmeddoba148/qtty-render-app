import os
import json
import imaplib
import email
import datetime
import requests
from zoneinfo import ZoneInfo
from email.header import decode_header

import gspread
from google.oauth2.service_account import Credentials


GMAIL_EMAIL = os.environ["GMAIL_EMAIL"]
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]
GOOGLE_CREDS_JSON = os.environ["GOOGLE_CREDS_JSON"]
GOOGLE_SHEET_NAME = os.environ.get("GOOGLE_SHEET_NAME", "mail_tracker")

TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
TELEGRAM_CHAT_ID = os.environ["TELEGRAM_CHAT_ID"]
RENDER_SITE_URL = os.environ.get("RENDER_SITE_URL", "").strip()


def cairo_now():
    return datetime.datetime.now(ZoneInfo("Africa/Cairo"))


def decode_mime_text(value):
    if not value:
        return ""
    parts = decode_header(value)
    result = []
    for part, enc in parts:
        if isinstance(part, bytes):
            result.append(part.decode(enc or "utf-8", errors="replace"))
        else:
            result.append(str(part))
    return "".join(result)


def get_google_book():
    creds_dict = json.loads(GOOGLE_CREDS_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open(GOOGLE_SHEET_NAME)


def get_main_sheet(book):
    sheet = book.sheet1
    return sheet


def get_or_create_notifications_sheet(book):
    try:
        ws = book.worksheet("notifications")
    except gspread.WorksheetNotFound:
        ws = book.add_worksheet(title="notifications", rows=1000, cols=4)
        ws.update("A1:D1", [["message_id", "subject", "notified_at", "date_key"]])
    return ws


def get_processed_message_ids(sheet):
    values = sheet.col_values(4)  # D = message_id
    return set(str(v).strip() for v in values[1:] if str(v).strip())


def get_notified_message_ids(ws):
    values = ws.col_values(1)
    return set(str(v).strip() for v in values[1:] if str(v).strip())


def already_ran_today(ws, date_key):
    values = ws.col_values(4)
    return date_key in set(str(v).strip() for v in values[1:] if str(v).strip())


def send_telegram(text):
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": True,
    }
    r = requests.post(url, json=payload, timeout=30)
    r.raise_for_status()


def get_message_real_id(msg, fallback_id):
    real_id = msg.get("Message-ID", "")
    return real_id.strip() if real_id else str(fallback_id)


def main():
    now = cairo_now()

    # ضمان إنه يشتغل فعليًا الساعة 2 ظهرًا بتوقيت القاهرة فقط
    if now.hour != 14:
        print(f"Skipped. Cairo time is {now.strftime('%H:%M')}")
        return

    date_key = now.strftime("%Y-%m-%d")

    book = get_google_book()
    main_sheet = get_main_sheet(book)
    notify_sheet = get_or_create_notifications_sheet(book)

    if already_ran_today(notify_sheet, date_key):
        print("Already checked today.")
        return

    processed_ids = get_processed_message_ids(main_sheet)
    notified_ids = get_notified_message_ids(notify_sheet)

    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)
    mail.select("INBOX")

    since_date = (now - datetime.timedelta(days=2)).strftime("%d-%b-%Y")
    status, messages = mail.search(None, f'(SINCE "{since_date}" SUBJECT "Qtty Recap")')

    if status != "OK" or not messages or not messages[0]:
        notify_sheet.append_row(["NO_NEW_MAIL", "No Qtty Recap", now.strftime("%Y-%m-%d %H:%M:%S"), date_key])
        mail.logout()
        print("No Qtty Recap emails.")
        return

    new_items = []

    for num in messages[0].split():
        status, msg_data = mail.fetch(num, "(BODY.PEEK[HEADER])")  # لا يعمل Seen
        if status != "OK" or not msg_data or not msg_data[0]:
            continue

        msg = email.message_from_bytes(msg_data[0][1])
        message_id = get_message_real_id(msg, num.decode())
        subject = decode_mime_text(msg.get("Subject", ""))

        if message_id in processed_ids:
            continue

        if message_id in notified_ids:
            continue

        new_items.append((message_id, subject))

    if not new_items:
        notify_sheet.append_row(["NO_NEW_UNPROCESSED", "No new unprocessed Qtty Recap", now.strftime("%Y-%m-%d %H:%M:%S"), date_key])
        mail.logout()
        print("No new unprocessed emails.")
        return

    lines = [
        "📩 <b>Qtty Recap جديد وصل</b>",
        "",
        f"العدد: {len(new_items)}",
    ]

    if RENDER_SITE_URL:
        lines += ["", f"افتح الموقع للتحميل:\n{RENDER_SITE_URL}"]

    send_telegram("\n".join(lines))

    notified_at = now.strftime("%Y-%m-%d %H:%M:%S")
    for message_id, subject in new_items:
        notify_sheet.append_row([message_id, subject, notified_at, date_key])

    mail.logout()
    print(f"Alert sent for {len(new_items)} email(s).")


if __name__ == "__main__":
    main()
