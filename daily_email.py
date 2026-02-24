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
    "cblake@kidsincrisis.org",
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
    no_entry = [n for n in TEAM if week_data.get(n, {}).get(day_name, "—") == "—"]

    # Build grouped rows
    status_groups = [
        ("🟢 In Office", "#22c55e", "#f0fdf4", in_office),
        ("🟡 AM Only", "#f59e0b", "#fffbeb", am_only),
        ("🔵 Remote", "#3b82f6", "#eff6ff", remote),
        ("🔴 Out", "#ef4444", "#fef2f2", out),
        ("⚪ No Entry", "#9ca3af", "#f9fafb", no_entry),
    ]

    grouped_html = ""
    for label, color, bg, members in status_groups:
        if not members:
            continue
        grouped_html += f"""
        <tr>
            <td colspan="2" style="padding:10px 12px; background:{bg}; font-weight:600; font-size:14px; border-bottom:1px solid #e5e7eb;">
                {label} ({len(members)})
            </td>
        </tr>
        """
        for name in members:
            desk = desk_bookings.get(name, "")
            desk_text = f"📍 {desk}" if desk else ""
            grouped_html += f"""
            <tr>
                <td style="padding:6px 12px 6px 24px; border-bottom:1px solid #f3f4f6; font-size:14px;">{name}</td>
                <td style="padding:6px 12px; border-bottom:1px solid #f3f4f6; color:#6b7280; font-size:13px; text-align:right;">{desk_text}</td>
            </tr>
            """

    # Build weekly view
    status_short = {
        "In Office": ("In", "#22c55e"),
        "In Office - AM Only": ("AM", "#f59e0b"),
        "Remote": ("Rem", "#3b82f6"),
        "Out": ("Out", "#ef4444"),
        "—": ("—", "#9ca3af"),
    }
    day_abbrevs = ["Mon", "Tue", "Wed", "Thu", "Fri"]

    # Determine today's column for highlighting
    today_day = date.strftime("%A")

    week_header = ""
    for i, abbr in enumerate(day_abbrevs):
        full_day = DAYS[i]
        highlight = "background-color:#eff6ff;" if full_day == today_day else ""
        bold = "<strong>" if full_day == today_day else ""
        bold_end = "</strong>" if full_day == today_day else ""
        week_header += f'<th style="padding:6px 4px; text-align:center; font-size:12px; color:#6b7280; border-bottom:2px solid #e5e7eb; {highlight}">{bold}{abbr}{bold_end}</th>'

    week_rows = ""
    for name in sorted(TEAM):
        week_rows += f'<tr><td style="padding:5px 8px; border-bottom:1px solid #f3f4f6; font-size:12px;">{name}</td>'
        for i, full_day in enumerate(DAYS):
            s = week_data.get(name, {}).get(full_day, "—")
            short, scolor = status_short.get(s, ("—", "#9ca3af"))
            highlight = "background-color:#f8faff;" if full_day == today_day else ""
            week_rows += f'<td style="padding:5px 4px; text-align:center; border-bottom:1px solid #f3f4f6; {highlight}"><span style="background:{scolor}; color:white; padding:2px 6px; border-radius:3px; font-size:11px;">{short}</span></td>'
        week_rows += "</tr>"

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
                <tbody>{grouped_html}</tbody>
            </table>

            <div style="margin-top:24px; padding-top:16px; border-top:2px solid #e5e7eb;">
                <h3 style="margin:0 0 12px; font-size:15px; color:#374151;">📅 This Week at a Glance</h3>
                <table style="width:100%; border-collapse:collapse;">
                    <thead>
                        <tr>
                            <th style="padding:6px 8px; text-align:left; font-size:12px; color:#6b7280; border-bottom:2px solid #e5e7eb;">Name</th>
                            {week_header}
                        </tr>
                    </thead>
                    <tbody>{week_rows}</tbody>
                </table>
            </div>
        </div>

        <div style="background:#f9fafb; padding:12px 24px; border:1px solid #e5e7eb; border-top:none; border-radius:0 0 10px 10px;">
            <p style="margin:0; font-size:12px; color:#9ca3af; text-align:center;">
                Need to update your status? Visit the <a href="https://inout-board-fdu78e24bjvyservvvnrax.streamlit.app/">In/Out Board Dashboard</a>
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
