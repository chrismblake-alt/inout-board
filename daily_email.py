"""
Daily Email Digest — In/Out Board
Sends a morning summary email to all staff showing who's where today.

Schedule this script to run at 7:00 AM on weekdays using:
  - GitHub Actions (free, recommended)
  - Streamlit Cloud cron (if available)
  - Any task scheduler (Windows Task Scheduler, cron, etc.)

Requires a Gmail account with an App Password for sending.
See README.md for setup instructions.
"""

import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import json
import os

# ── Configuration ──────────────────────────────────────────────────────────────
TEAM = [
    "Adelaide", "Brittany", "Chris", "Daniel",
    "Erica", "Ginny", "Karen", "Katie", "Paula"
]
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

# Load from environment variables (set in GitHub Actions secrets or .env)
GCP_CREDS = os.environ.get("GCP_SERVICE_ACCOUNT")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")
SMTP_EMAIL = os.environ.get("SMTP_EMAIL")       # e.g., chris@kidsincrisis.org
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD")  # App Password
SMTP_SERVER = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))

# Email recipients — add your team's email addresses here
RECIPIENTS = [
    # "chris@kidsincrisis.org",
    # "katie@kidsincrisis.org",
    # etc.
]

# ── Google Sheets Connection ───────────────────────────────────────────────────
def get_spreadsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_dict = json.loads(GCP_CREDS)
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)

def get_current_week_monday():
    today = datetime.now().date()
    return today - timedelta(days=today.weekday())

def load_weekly_status(ws, week_of):
    records = ws.get_all_records()
    week_str = week_of.strftime("%Y-%m-%d")
    week_data = {}
    for row in records:
        if str(row.get("Week_Of", "")) == week_str:
            name = row.get("Your Name", "")
            if name in TEAM:
                week_data[name] = {
                    "Monday": row.get("Monday", "—") or "—",
                    "Tuesday": row.get("Tuesday", "—") or "—",
                    "Wednesday": row.get("Wednesday", "—") or "—",
                    "Thursday": row.get("Thursday", "—") or "—",
                    "Friday": row.get("Friday", "—") or "—",
                }
    for name in TEAM:
        if name not in week_data:
            week_data[name] = {day: "—" for day in DAYS}
    return week_data

def load_desk_bookings(ws, date):
    records = ws.get_all_records()
    date_str = date.strftime("%Y-%m-%d")
    bookings = {}
    for row in records:
        if row.get("Date") == date_str:
            bookings[row.get("Name")] = row.get("Desk")
    return bookings

# ── Build Email ────────────────────────────────────────────────────────────────
def build_email_html(day_name, date, week_data, desk_bookings):
    status_colors = {
        "In Office": "#22c55e",
        "In Office - AM Only": "#f59e0b",
        "Remote": "#3b82f6",
        "Out": "#ef4444",
        "—": "#9ca3af",
    }

    in_office = [n for n in TEAM if week_data.get(n, {}).get(day_name, "—") == "In Office"]
    am_only = [n for n in TEAM if week_data.get(n, {}).get(day_name, "—") == "In Office - AM Only"]
    remote = [n for n in TEAM if week_data.get(n, {}).get(day_name, "—") == "Remote"]
    out = [n for n in TEAM if week_data.get(n, {}).get(day_name, "—") == "Out"]
    no_entry = [n for n in TEAM if week_data.get(n, {}).get(day_name, "—") == "—"]

    # Build rows
    rows_html = ""
    for name in TEAM:
        status = week_data.get(name, {}).get(day_name, "—")
        color = status_colors.get(status, "#9ca3af")
        desk = desk_bookings.get(name, "")
        desk_text = f" — {desk}" if desk else ""

        rows_html += f"""
        <tr>
            <td style="padding:8px 12px; border-bottom:1px solid #f3f4f6; font-size:14px;">{name}</td>
            <td style="padding:8px 12px; border-bottom:1px solid #f3f4f6; text-align:center;">
                <span style="background-color:{color}; color:white; padding:3px 10px; border-radius:4px; font-size:13px; font-weight:600;">
                    {status}
                </span>
            </td>
            <td style="padding:8px 12px; border-bottom:1px solid #f3f4f6; color:#6b7280; font-size:13px;">{desk_text}</td>
        </tr>
        """

    html = f"""
    <html>
    <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; max-width:600px; margin:0 auto;">
        <div style="background:linear-gradient(135deg, #3b82f6, #1d4ed8); color:white; padding:20px 24px; border-radius:10px 10px 0 0;">
            <h2 style="margin:0;">🏢 {day_name}'s In/Out Board</h2>
            <p style="margin:4px 0 0; opacity:0.9;">{date.strftime('%B %d, %Y')} — Development Team</p>
        </div>

        <div style="background:white; padding:16px 24px; border:1px solid #e5e7eb;">
            <div style="display:flex; gap:16px; margin-bottom:16px; text-align:center;">
                <div style="flex:1; padding:8px; background:#f0fdf4; border-radius:6px;">
                    <div style="font-size:24px; font-weight:700; color:#22c55e;">{len(in_office)}</div>
                    <div style="font-size:12px; color:#6b7280;">In Office</div>
                </div>
                <div style="flex:1; padding:8px; background:#fffbeb; border-radius:6px;">
                    <div style="font-size:24px; font-weight:700; color:#f59e0b;">{len(am_only)}</div>
                    <div style="font-size:12px; color:#6b7280;">AM Only</div>
                </div>
                <div style="flex:1; padding:8px; background:#eff6ff; border-radius:6px;">
                    <div style="font-size:24px; font-weight:700; color:#3b82f6;">{len(remote)}</div>
                    <div style="font-size:12px; color:#6b7280;">Remote</div>
                </div>
                <div style="flex:1; padding:8px; background:#fef2f2; border-radius:6px;">
                    <div style="font-size:24px; font-weight:700; color:#ef4444;">{len(out)}</div>
                    <div style="font-size:12px; color:#6b7280;">Out</div>
                </div>
            </div>

            <table style="width:100%; border-collapse:collapse;">
                <thead>
                    <tr style="border-bottom:2px solid #e5e7eb;">
                        <th style="padding:8px 12px; text-align:left; font-size:13px; color:#6b7280;">Name</th>
                        <th style="padding:8px 12px; text-align:center; font-size:13px; color:#6b7280;">Status</th>
                        <th style="padding:8px 12px; text-align:left; font-size:13px; color:#6b7280;">Desk</th>
                    </tr>
                </thead>
                <tbody>{rows_html}</tbody>
            </table>
        </div>

        <div style="background:#f9fafb; padding:12px 24px; border:1px solid #e5e7eb; border-top:none; border-radius:0 0 10px 10px;">
            <p style="margin:0; font-size:12px; color:#9ca3af; text-align:center;">
                Need to update your status? Visit the <a href="YOUR_STREAMLIT_APP_URL">In/Out Board Dashboard</a>
            </p>
        </div>
    </body>
    </html>
    """
    return html

# ── Send Email ─────────────────────────────────────────────────────────────────
def send_digest():
    today = datetime.now().date()
    day_name = today.strftime("%A")

    if day_name not in DAYS:
        print("Weekend — no digest sent.")
        return

    spreadsheet = get_spreadsheet()
    status_ws = spreadsheet.worksheet("Weekly Status")
    desk_ws = spreadsheet.worksheet("Desk Bookings")

    week_monday = get_current_week_monday()
    week_data = load_weekly_status(status_ws, week_monday)
    desk_bookings = load_desk_bookings(desk_ws, today)

    html = build_email_html(day_name, today, week_data, desk_bookings)

    # Build email
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"🏢 {day_name}'s In/Out Board — {today.strftime('%b %d')}"
    msg["From"] = SMTP_EMAIL
    msg["To"] = ", ".join(RECIPIENTS)
    msg.attach(MIMEText(html, "html"))

    # Send
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_EMAIL, SMTP_PASSWORD)
        server.sendmail(SMTP_EMAIL, RECIPIENTS, msg.as_string())

    print(f"✅ Digest sent for {day_name}, {today} to {len(RECIPIENTS)} recipients.")

if __name__ == "__main__":
    send_digest()
