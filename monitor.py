import os
import smtplib
import time
from datetime import datetime, timezone
from email.message import EmailMessage
import openpyxl
from openpyxl.styles import Font, PatternFill
import requests

TARGET_URL = os.environ.get("TARGET_URL", "https://example.com")
APP_NAME = os.environ.get("APP_NAME", "ROBO-AI-AGENT")
EXCEL_FILE = "Assessment_Results.xlsx"

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")
RECIPIENT_EMAIL = os.environ.get("RECIPIENT_EMAIL", "")


def check_runtime_health(url):
    start_time = time.time()
    try:
        response = requests.get(url, timeout=20)
        elapsed_ms = round((time.time() - start_time) * 1000, 2)
        status = "PASSED" if response.status_code == 200 else "FAILED"
        return status, response.status_code, elapsed_ms
    except requests.RequestException:
        elapsed_ms = round((time.time() - start_time) * 1000, 2)
        return "UNREACHABLE", 0, elapsed_ms


def update_excel_report(status, code, latency_ms):
    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Runtime_Assessments"
        headers = ["Timestamp (UTC)", "Application", "Status", "HTTP Code", "Response Time (ms)"]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    ws.append([timestamp, APP_NAME, status, code, latency_ms])

    # Highlight Status cell green for PASSED or red for FAILED
    status_cell = ws.cell(row=ws.max_row, column=3)
    if status == "PASSED":
        status_cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    else:
        status_cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    # Auto-adjust column widths for clean viewing in Microsoft Excel
    for column_cells in ws.columns:
        max_length = max(len(str(cell.value or "")) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = max_length + 4

    wb.save(EXCEL_FILE)
    print(f"Successfully updated {EXCEL_FILE} with status: {status}")


def send_outlook_report(status, code, latency_ms):
    if not SENDER_EMAIL or not EMAIL_PASSWORD or not RECIPIENT_EMAIL:
        print("Email credentials not set; skipping email delivery.")
        return

    msg = EmailMessage()
    msg["Subject"] = f"[{APP_NAME}] CI/CD Runtime Assessment: {status}"
    msg["From"] = SENDER_EMAIL
    msg["To"] = RECIPIENT_EMAIL
    msg.set_content(
        f"Automated GitHub Actions CI/CD Runtime Assessment\n\n"
        f"Application: {APP_NAME}\n"
        f"Target URL: {TARGET_URL}\n"
        f"Runtime Status: {status}\n"
        f"HTTP Code: {code}\n"
        f"Response Time: {latency_ms} ms\n\n"
        f"Click the attached {EXCEL_FILE} file to open it directly in Microsoft Excel!"
    )

    with open(EXCEL_FILE, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=EXCEL_FILE,
        )

    with smtplib.SMTP_SSL("://gmail.com", 465) as server:
        server.login(SENDER_EMAIL, EMAIL_PASSWORD)
        server.send_message(msg)
    print(f"Successfully delivered {EXCEL_FILE} to {RECIPIENT_EMAIL}")


if __name__ == "__main__":
    status, code, latency_ms = check_runtime_health(TARGET_URL)
    update_excel_report(status, code, latency_ms)
    send_outlook_report(status, code, latency_ms)
