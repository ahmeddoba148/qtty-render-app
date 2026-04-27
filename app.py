import os
import re
import imaplib
import email
import datetime
import traceback
import pandas as pd

from email.header import decode_header
from flask import Flask, jsonify, render_template_string

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


app = Flask(__name__)

GMAIL_EMAIL = os.environ.get("GMAIL_EMAIL", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")

CURRENT_YEAR = datetime.datetime.now().year
BASE_DIR = os.getcwd()

DOWNLOAD_PATH = os.path.join(BASE_DIR, "QTTY_RECAPS")
EDIT_PATH = os.path.join(DOWNLOAD_PATH, "Edit")
MAIL_TRACKER_FILE = os.path.join(BASE_DIR, "mail_tracker.csv")

os.makedirs(DOWNLOAD_PATH, exist_ok=True)
os.makedirs(EDIT_PATH, exist_ok=True)

BLACK_BORDER = Border(
    top=Side(border_style="medium", color="000000"),
    bottom=Side(border_style="thin", color="000000"),
    left=Side(border_style="thin", color="000000"),
    right=Side(border_style="thin", color="000000")
)


HOME_HTML = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>QTTY Recap</title>
<style>
body{
    margin:0;
    height:100vh;
    font-family:Arial,Tahoma,sans-serif;
    background:linear-gradient(135deg,#0f172a,#1e3a8a);
    display:flex;
    justify-content:center;
    align-items:center;
    color:white;
}
.card{
    width:90%;
    max-width:420px;
    background:rgba(255,255,255,0.12);
    border:1px solid rgba(255,255,255,0.2);
    border-radius:28px;
    padding:28px;
    text-align:center;
    box-shadow:0 20px 60px rgba(0,0,0,0.35);
}
h1{font-size:26px;margin-bottom:10px}
p{opacity:.85}
button{
    width:100%;
    padding:18px;
    border:0;
    border-radius:18px;
    background:#22c55e;
    color:white;
    font-size:22px;
    font-weight:bold;
    cursor:pointer;
    margin-top:20px;
}
#result{
    margin-top:20px;
    background:rgba(0,0,0,0.25);
    border-radius:16px;
    padding:15px;
    min-height:40px;
    white-space:pre-line;
}
</style>
</head>
<body>
<div class="card">
    <h1>QTTY Recap Downloader</h1>
    <p>هيقرأ فقط الإيميلات غير المقروءة ويحوّل ملفات Edit فقط</p>
    <button onclick="runProcess()">تشغيل الآن</button>
    <div id="result">جاهز للتشغيل ✅</div>
</div>

<script>
async function runProcess(){
    const result = document.getElementById("result");
    result.innerText = "جاري التشغيل... انتظر";
    try{
        const res = await fetch("/run");
        const data = await res.json();
        result.innerText = data.message;
    }catch(e){
        result.innerText = "حدث خطأ: " + e;
    }
}
</script>
</body>
</html>
"""


def log(msg):
    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}", flush=True)


def decode_mime_text(value):
    if not value:
        return ""

    decoded_parts = decode_header(value)
    result = []

    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            try:
                result.append(part.decode(encoding or "utf-8", errors="replace"))
            except Exception:
                result.append(part.decode("latin-1", errors="replace"))
        else:
            result.append(str(part))

    return "".join(result)


def clean_filename(filename: str) -> str:
    if isinstance(filename, bytes):
        try:
            filename = filename.decode("utf-8")
        except UnicodeDecodeError:
            filename = filename.decode("latin-1", errors="replace")

    cleaned = re.sub(r'[\\/*?:"<>|]', "_", filename)
    return cleaned.strip()


def get_mail_tracker_file() -> str:
    if not os.path.exists(MAIL_TRACKER_FILE):
        df = pd.DataFrame([{
            "message_id": "startfile",
            "file_name": "startfile",
            "year": str(CURRENT_YEAR),
            "page": "0"
        }])
        df.to_csv(MAIL_TRACKER_FILE, index=False, encoding="utf-8")

    return MAIL_TRACKER_FILE


def download_attachments() -> list:
    if not GMAIL_EMAIL or not GMAIL_APP_PASSWORD:
        raise Exception("GMAIL_EMAIL أو GMAIL_APP_PASSWORD غير موجودين في Environment Variables")

    downloaded_files = []
    new_rows = []

    log("Connecting to Gmail...")

    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)
    mail.select("INBOX")

    # مهم: يقرأ فقط الإيميلات غير المقروءة
    status, messages = mail.search(None, 'UNSEEN SUBJECT "Qtty Recap"')

    if status != "OK" or not messages or not messages[0]:
        mail.logout()
        return []

    all_message_ids = messages[0].split()

    df = pd.read_csv(get_mail_tracker_file())
    existing_ids = set(df["message_id"].astype(str).dropna())

    message_ids_to_process = [
        msg_id.decode()
        for msg_id in all_message_ids
        if msg_id.decode() not in existing_ids
    ]

    if not message_ids_to_process:
        mail.logout()
        return []

    for num_str in message_ids_to_process:
        message_processed_successfully = False

        try:
            status, msg_data = mail.fetch(num_str.encode(), "(RFC822)")

            if status != "OK" or not msg_data or not msg_data[0]:
                continue

            msg = email.message_from_bytes(msg_data[0][1])
            subject = decode_mime_text(msg.get("Subject", ""))
            log(f"Processing email: {subject}")

            date_str = msg.get("Date", "")
            file_year = CURRENT_YEAR
            local_date_prefix = datetime.datetime.now().strftime("%Y%m%d")

            if date_str:
                try:
                    date_tuple = email.utils.parsedate_tz(date_str)
                    if date_tuple:
                        local_timestamp = email.utils.mktime_tz(date_tuple)
                        local_dt = datetime.datetime.fromtimestamp(local_timestamp)
                        file_year = local_dt.year
                        local_date_prefix = local_dt.strftime("%Y%m%d")
                except Exception:
                    pass

            file_year_str = str(file_year)
            files_saved_from_this_email = 0

            for part in msg.walk():
                if part.get_content_maintype() == "multipart":
                    continue

                if part.get("Content-Disposition") is None:
                    continue

                filename = part.get_filename()

                if not filename:
                    continue

                original_filename = decode_mime_text(filename)

                if not original_filename.lower().endswith((".xlsx", ".xls")):
                    continue

                clean_name = clean_filename(f"{local_date_prefix}_{original_filename}")
                filepath = os.path.join(DOWNLOAD_PATH, clean_name)

                payload = part.get_payload(decode=True)

                if not payload:
                    continue

                with open(filepath, "wb") as f:
                    f.write(payload)

                downloaded_files.append(filepath)
                files_saved_from_this_email += 1
                log(f"Downloaded: {clean_name}")

                try:
                    wb_temp = load_workbook(filepath, read_only=True, data_only=True)
                    ws_temp = wb_temp.active

                    target_col_idx = None

                    for cell in ws_temp[1]:
                        if cell.value and "delivery date" in str(cell.value).lower():
                            target_col_idx = cell.column
                            break

                    if target_col_idx:
                        cell_val = ws_temp.cell(row=2, column=target_col_idx).value

                        if cell_val:
                            extracted_suffix = None

                            if isinstance(cell_val, (datetime.datetime, datetime.date)):
                                extracted_suffix = str(cell_val.year)[-2:]
                            else:
                                digits = re.findall(r"\d", str(cell_val))
                                if len(digits) >= 2:
                                    extracted_suffix = "".join(digits[-2:])

                            if extracted_suffix:
                                file_year_str = "20" + extracted_suffix

                    wb_temp.close()

                except Exception as e:
                    log(f"Year extraction warning: {e}")

                new_rows.append({
                    "message_id": num_str,
                    "file_name": clean_name,
                    "year": file_year_str,
                    "page": "0"
                })

            # لو الإيميل اتفحص واتعاملنا معاه، نعلمه مقروء عشان مايتكررش
            if files_saved_from_this_email > 0:
                message_processed_successfully = True
                mail.store(num_str.encode(), '+FLAGS', '\\Seen')
                log(f"Marked email as read: {num_str}")

        except Exception as e:
            log(f"Error processing email {num_str}: {e}")
            log(traceback.format_exc())

    if new_rows:
        new_df = pd.DataFrame(new_rows)
        updated_df = pd.concat([df, new_df], ignore_index=True)
        updated_df.to_csv(MAIL_TRACKER_FILE, index=False, encoding="utf-8")

    mail.close()
    mail.logout()

    return downloaded_files


def get_page_title(file_name: str) -> dict:
    df = pd.read_csv(get_mail_tracker_file())

    df["message_id"] = df["message_id"].astype(str)
    df["year"] = df["year"].astype(str)
    df["file_name"] = df["file_name"].astype(str)

    file_row = df[df["file_name"] == file_name]

    if not file_row.empty:
        file_year_str = str(file_row.iloc[0]["year"])
    else:
        file_year_str = str(CURRENT_YEAR)

    page_title = "PAGE 00/00"
    output_base_name = file_name.replace(".xlsx", "").replace(".xls", "")
    is_update = "updated" in file_name.lower()

    if is_update:
        page_title = "PAGE 00/00"
        output_base_name = f"{output_base_name}_updated"
    else:
        try:
            year_pages = pd.to_numeric(
                df[df["year"] == file_year_str]["page"],
                errors="coerce"
            )
            max_page = year_pages.max()
            next_page = int(max_page + 1) if pd.notna(max_page) else 1
        except Exception:
            next_page = 1

        file_year_suffix = file_year_str[-2:]
        page_title = f"PAGE {next_page}/{file_year_suffix}"
        output_base_name = f"PAGE {next_page}-{file_year_suffix}"

        match_index = df[
            (df["file_name"] == file_name) &
            (df["year"] == file_year_str)
        ].index

        if not match_index.empty:
            df.loc[match_index[0], "page"] = str(next_page)
            df.to_csv(MAIL_TRACKER_FILE, index=False, encoding="utf-8")

    return {
        "page_title": page_title,
        "file_name": output_base_name
    }


def transform_excel_file(input_file: str, page_title: str, output_file_name: str) -> str | None:
    try:
        output_file_path = os.path.join(EDIT_PATH, f"{output_file_name}_edit.xlsx")

        wb = load_workbook(input_file)
        ws = wb.active

        df = pd.read_excel(input_file, engine="openpyxl")

        header_row_index = 1
        headers = [cell.value for cell in ws[header_row_index] if cell.value is not None]

        keywords_to_delete = ["customer", "price", "factory"]
        cols_to_delete_indices = []

        for idx, header in enumerate(headers):
            if header and any(kw.lower() in str(header).lower() for kw in keywords_to_delete):
                cols_to_delete_indices.append(idx + 1)

        for col_idx in sorted(cols_to_delete_indices, reverse=True):
            ws.delete_cols(col_idx)

        current_num_columns = ws.max_column
        ws.insert_rows(1)

        title_cell = ws.cell(row=1, column=1)
        title_cell.value = page_title
        title_cell.font = Font(name="Arial", color="FF0000", bold=True, size=36)
        title_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=current_num_columns
        )

        ws["A2"].fill = PatternFill(fill_type=None)

        if "Style # " not in df.columns or "Quantity Ordered" not in df.columns:
            raise Exception("الأعمدة المطلوبة غير موجودة: Style # و Quantity Ordered")

        df["Style_number"] = df["Style # "].astype(str)
        df["last4"] = df["Style_number"].str.slice(-4)
        df["Quantity Ordered"] = pd.to_numeric(
            df["Quantity Ordered"],
            errors="coerce"
        ).fillna(0)

        grouped = df.groupby("last4")["Quantity Ordered"].sum().reset_index()

        insert_col_index = ws.max_column
        ws.insert_cols(insert_col_index + 1, 2)

        total_col_letter = get_column_letter(insert_col_index + 1)
        model_col_letter = get_column_letter(insert_col_index + 2)

        header_row_num = 2

        ws[f"{total_col_letter}{header_row_num}"].value = "الإجمالي"
        ws[f"{model_col_letter}{header_row_num}"].value = "الموديل"

        header_font = Font(color="FFFFFF", bold=True)
        header_fill = PatternFill(
            start_color="4a90e2",
            end_color="4a90e2",
            fill_type="solid"
        )
        header_alignment = Alignment(horizontal="center", vertical="center")

        for col_letter in [total_col_letter, model_col_letter]:
            cell = ws[f"{col_letter}{header_row_num}"]
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = BLACK_BORDER

        cell_font = Font(name="Arial", size=20, color="FF0000", bold=True)
        cell_alignment = Alignment(horizontal="center", vertical="center")

        visited_last4 = set()
        base_row_offset = 3

        for last4_val in df["last4"]:
            if last4_val in visited_last4:
                continue

            group_df_indices = df[df["last4"] == last4_val].index.tolist()

            if not group_df_indices:
                continue

            start_excel_row = group_df_indices[0] + base_row_offset
            end_excel_row = group_df_indices[-1] + base_row_offset

            total_qty_series = grouped[grouped["last4"] == last4_val]["Quantity Ordered"]
            total_qty = total_qty_series.values[0] if not total_qty_series.empty else 0

            total_cell = ws[f"{total_col_letter}{start_excel_row}"]
            total_cell.value = total_qty
            total_cell.font = cell_font
            total_cell.alignment = cell_alignment

            if end_excel_row >= start_excel_row:
                ws.merge_cells(
                    start_row=start_excel_row,
                    start_column=insert_col_index + 1,
                    end_row=end_excel_row,
                    end_column=insert_col_index + 1
                )

            for row_num in range(start_excel_row, end_excel_row + 1):
                ws.cell(row=row_num, column=insert_col_index + 1).border = BLACK_BORDER

            model_cell = ws[f"{model_col_letter}{start_excel_row}"]
            model_cell.value = last4_val
            model_cell.font = cell_font
            model_cell.alignment = cell_alignment

            if end_excel_row >= start_excel_row:
                ws.merge_cells(
                    start_row=start_excel_row,
                    start_column=insert_col_index + 2,
                    end_row=end_excel_row,
                    end_column=insert_col_index + 2
                )

            for row_num in range(start_excel_row, end_excel_row + 1):
                ws.cell(row=row_num, column=insert_col_index + 2).border = BLACK_BORDER

            visited_last4.add(last4_val)

        current_headers = [cell.value for cell in ws[2]]

        column_widths = {
            "po #": 20,
            "style #": 18,
            "size": 15,
            "date": 12,
            "breakdown": 12,
            "quantity ordered": 12,
            "poly bag": 10,
            "الإجمالي": 12,
            "الموديل": 12
        }

        for col_idx, header_val in enumerate(current_headers):
            if not header_val:
                continue

            col_letter = get_column_letter(col_idx + 1)
            header_lower = str(header_val).lower()

            for key, width in column_widths.items():
                if key.lower() in header_lower:
                    ws.column_dimensions[col_letter].width = width
                    break

        wb.save(output_file_path)
        log(f"Saved transformed file in Edit folder: {output_file_path}")

        return output_file_path

    except Exception as e:
        log(f"Transform error: {e}")
        log(traceback.format_exc())
        return None


def process_all():
    downloaded_files = download_attachments()

    edit_files_only = []

    for file_path in downloaded_files:
        file_name = os.path.basename(file_path)

        page_data = get_page_title(file_name)

        transformed_file = transform_excel_file(
            input_file=file_path,
            page_title=page_data["page_title"],
            output_file_name=page_data["file_name"]
        )

        if transformed_file:
            edit_files_only.append(transformed_file)

    return {
        "downloaded": len(downloaded_files),
        "edit_files_count": len(edit_files_only),
        "edit_files": edit_files_only
    }


@app.route("/")
def home():
    return render_template_string(HOME_HTML)


@app.route("/run")
def run():
    try:
        result = process_all()

        msg = (
            f"تم التشغيل بنجاح ✅\n"
            f"تم تحميل: {result['downloaded']} ملف\n"
            f"ملفات Edit الجاهزة: {result['edit_files_count']} ملف\n\n"
            f"ملاحظة: أي إرسال لاحق سيتم من ملفات Edit فقط."
        )

        return jsonify({
            "success": True,
            "message": msg,
            "edit_files_only": result["edit_files"]
        })

    except Exception as e:
        log(traceback.format_exc())

        return jsonify({
            "success": False,
            "message": f"حدث خطأ ❌\n{str(e)}"
        }), 500


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
