import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import json
import pandas as pd

# ── Configuration ──────────────────────────────────────────────────────────────
TEAM = [
    "Adelaide", "Amy", "Angela", "Chris", "Daniel",
    "Erica", "Erick", "Karen", "Katie"
]
DESKS = [f"Desk {i}" for i in range(1, 9)]
PERMANENT_DESKS = {"Chris": "Desk 8"}  # Always assigned, not bookable
BOOKABLE_DESKS = [d for d in [f"Desk {i}" for i in range(1, 9)] if d not in PERMANENT_DESKS.values()]
STATUSES = ["In Office", "In Office - AM Only", "Remote", "Out"]
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

STATUS_COLORS = {
    "In Office": "#22c55e",           # green
    "In Office - AM Only": "#f59e0b", # amber
    "Remote": "#3b82f6",              # blue
    "Out": "#ef4444",                 # red
    "—": "#9ca3af",                   # gray (no entry)
}

STATUS_EMOJI = {
    "In Office": "🟢",
    "In Office - AM Only": "🟡",
    "Remote": "🔵",
    "Out": "🔴",
    "—": "⚪",
}

# ── Google Sheets Connection ───────────────────────────────────────────────────
@st.cache_resource
def get_gsheet_connection():
    """Connect to Google Sheets using service account credentials."""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client

def get_spreadsheet():
    """Get the spreadsheet by ID from secrets."""
    client = get_gsheet_connection()
    return client.open_by_key(st.secrets["spreadsheet_id"])

# ── Helper Functions ───────────────────────────────────────────────────────────
def get_current_week_monday():
    """Get the Monday of the current week."""
    today = datetime.now().date()
    return today - timedelta(days=today.weekday())

def get_next_week_monday():
    """Get the Monday of next week."""
    return get_current_week_monday() + timedelta(days=7)

def get_today_day_name():
    """Get today's day name (Monday, Tuesday, etc.)."""
    return datetime.now().strftime("%A")

def ensure_weekly_status_sheet(spreadsheet):
    """Ensure the Weekly Status sheet exists with proper headers."""
    try:
        ws = spreadsheet.worksheet("Weekly Status")
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title="Weekly Status", rows=200, cols=8)
        ws.update("A1:H1", [["Your Name", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Submission ID", "Week_Of"]])
        ws.format("A1:H1", {"textFormat": {"bold": True}})
    return ws

def ensure_desk_bookings_sheet(spreadsheet):
    """Ensure the Desk Bookings sheet exists with proper headers."""
    try:
        ws = spreadsheet.worksheet("Desk Bookings")
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title="Desk Bookings", rows=500, cols=3)
        ws.update("A1:C1", [["Date", "Name", "Desk"]])
        ws.format("A1:C1", {"textFormat": {"bold": True}})
    return ws

# ── Data Operations ────────────────────────────────────────────────────────────
def load_weekly_status(ws, week_of):
    """Load status data for a given week."""
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
    # Fill in missing team members
    for name in TEAM:
        if name not in week_data:
            week_data[name] = {day: "—" for day in DAYS}
    return week_data

def save_weekly_status(ws, week_of, name, statuses_dict):
    """Save or update a person's weekly status."""
    records = ws.get_all_values()
    week_str = week_of.strftime("%Y-%m-%d")

    # Find existing row
    row_idx = None
    for i, row in enumerate(records):
        if i == 0:
            continue
        if len(row) >= 8 and row[0] == name and row[7] == week_str:
            row_idx = i + 1  # 1-indexed for gspread
            break
        # Also check if Week_Of is empty but name matches and was recently added
        if len(row) >= 1 and row[0] == name and (len(row) < 8 or row[7] == "" or row[7] == week_str):
            row_idx = i + 1
            break

    row_data = [
        name,
        statuses_dict.get("Monday", "—"),
        statuses_dict.get("Tuesday", "—"),
        statuses_dict.get("Wednesday", "—"),
        statuses_dict.get("Thursday", "—"),
        statuses_dict.get("Friday", "—"),
        "",  # Submission ID (blank for manual entries)
        week_str,
    ]

    if row_idx:
        ws.update(f"A{row_idx}:H{row_idx}", [row_data])
    else:
        ws.append_row(row_data)

def load_desk_bookings(ws, date):
    """Load desk bookings for a given date."""
    records = ws.get_all_records()
    date_str = date.strftime("%Y-%m-%d")
    bookings = {}
    for row in records:
        if row.get("Date") == date_str:
            bookings[row.get("Name")] = row.get("Desk")
    return bookings

def save_desk_booking(ws, date, name, desk):
    """Save or update a desk booking."""
    records = ws.get_all_values()
    date_str = date.strftime("%Y-%m-%d")

    # Find existing row
    row_idx = None
    for i, row in enumerate(records):
        if i == 0:
            continue
        if len(row) >= 2 and row[0] == date_str and row[1] == name:
            row_idx = i + 1
            break

    if desk == "None":
        # Remove booking
        if row_idx:
            ws.delete_rows(row_idx)
        return

    row_data = [date_str, name, desk]
    if row_idx:
        ws.update(f"A{row_idx}:C{row_idx}", [row_data])
    else:
        ws.append_row(row_data)

def clear_desk_booking(ws, date, name):
    """Remove a desk booking."""
    records = ws.get_all_values()
    date_str = date.strftime("%Y-%m-%d")
    for i, row in enumerate(records):
        if i == 0:
            continue
        if len(row) >= 2 and row[0] == date_str and row[1] == name:
            ws.delete_rows(i + 1)
            return

# ── Page Config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="In/Out Board — Kids in Crisis",
    page_icon="🏢",
    layout="wide",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Clean up spacing */
    .block-container { padding-top: 1rem; }

    /* Status badge styles */
    .status-badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 6px;
        color: white;
        font-weight: 600;
        font-size: 0.85rem;
        text-align: center;
        min-width: 80px;
    }
    .status-in { background-color: #22c55e; }
    .status-am { background-color: #f59e0b; }
    .status-remote { background-color: #3b82f6; }
    .status-out { background-color: #ef4444; }
    .status-none { background-color: #d1d5db; color: #6b7280; }

    /* Desk grid */
    .desk-available {
        background-color: #f0fdf4;
        border: 2px solid #22c55e;
        border-radius: 8px;
        padding: 12px;
        text-align: center;
        min-height: 80px;
    }
    .desk-taken {
        background-color: #fef2f2;
        border: 2px solid #ef4444;
        border-radius: 8px;
        padding: 12px;
        text-align: center;
        min-height: 80px;
    }

    /* Today highlight */
    .today-header {
        background: linear-gradient(135deg, #3b82f6, #1d4ed8);
        color: white;
        padding: 16px 24px;
        border-radius: 10px;
        margin-bottom: 1rem;
    }

    /* Summary cards */
    .summary-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    .summary-number {
        font-size: 2rem;
        font-weight: 700;
        line-height: 1.2;
    }
    .summary-label {
        font-size: 0.85rem;
        color: #6b7280;
        margin-top: 4px;
    }
</style>
""", unsafe_allow_html=True)

# ── Main App ───────────────────────────────────────────────────────────────────
try:
    spreadsheet = get_spreadsheet()
    status_ws = ensure_weekly_status_sheet(spreadsheet)
    desk_ws = ensure_desk_bookings_sheet(spreadsheet)
except Exception as e:
    st.error(f"⚠️ Could not connect to Google Sheets. Check your secrets configuration.\n\n{e}")
    st.stop()

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("🏢 In/Out Board")
    st.caption("Third Floor — Kids in Crisis")
    st.divider()
    page = st.radio(
        "Navigate",
        ["📊 Today's Dashboard", "🪑 Book a Desk", "📝 Update My Week", "📅 Weekly View"],
        label_visibility="collapsed",
    )
    st.divider()
    st.caption(f"Today: {datetime.now().strftime('%A, %B %d, %Y')}")
    if st.button("🔄 Refresh Data"):
        st.cache_resource.clear()
        st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: Today's Dashboard
# ══════════════════════════════════════════════════════════════════════════════
if page == "📊 Today's Dashboard":
    today = datetime.now().date()
    day_name = get_today_day_name()
    week_monday = get_current_week_monday()

    # Check if today is a weekend
    if day_name not in DAYS:
        st.info("It's the weekend! Showing Friday's status.")
        day_name = "Friday"

    week_data = load_weekly_status(status_ws, week_monday)
    desk_bookings = load_desk_bookings(desk_ws, today)

    # Header
    st.markdown(f"""
    <div class="today-header">
        <h2 style="margin:0; color:white;">📊 {day_name}'s Board</h2>
        <p style="margin:0; opacity:0.9; color: #e0e0e0;">
            {today.strftime('%B %d, %Y')} — Third Floor
        </p>
    </div>
    """, unsafe_allow_html=True)

    # Summary counts
    in_office = sum(1 for n in TEAM if week_data.get(n, {}).get(day_name, "—") == "In Office")
    am_only = sum(1 for n in TEAM if week_data.get(n, {}).get(day_name, "—") == "In Office - AM Only")
    remote = sum(1 for n in TEAM if week_data.get(n, {}).get(day_name, "—") == "Remote")
    out = sum(1 for n in TEAM if week_data.get(n, {}).get(day_name, "—") == "Out")
    no_entry = sum(1 for n in TEAM if week_data.get(n, {}).get(day_name, "—") == "—")

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.markdown(f"""<div class="summary-card">
            <div class="summary-number" style="color: #22c55e;">{in_office}</div>
            <div class="summary-label">🟢 In Office</div>
        </div>""", unsafe_allow_html=True)
    with col2:
        st.markdown(f"""<div class="summary-card">
            <div class="summary-number" style="color: #f59e0b;">{am_only}</div>
            <div class="summary-label">🟡 AM Only</div>
        </div>""", unsafe_allow_html=True)
    with col3:
        st.markdown(f"""<div class="summary-card">
            <div class="summary-number" style="color: #3b82f6;">{remote}</div>
            <div class="summary-label">🔵 Remote</div>
        </div>""", unsafe_allow_html=True)
    with col4:
        st.markdown(f"""<div class="summary-card">
            <div class="summary-number" style="color: #ef4444;">{out}</div>
            <div class="summary-label">🔴 Out</div>
        </div>""", unsafe_allow_html=True)
    with col5:
        st.markdown(f"""<div class="summary-card">
            <div class="summary-number" style="color: #9ca3af;">{no_entry}</div>
            <div class="summary-label">⚪ No Entry</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Status table — grouped by status
    STATUS_ORDER = ["In Office", "In Office - AM Only", "Remote", "Out", "—"]
    STATUS_GROUP_LABELS = {
        "In Office": "🟢 In Office",
        "In Office - AM Only": "🟡 In Office - AM Only",
        "Remote": "🔵 Remote",
        "Out": "🔴 Out",
        "—": "⚪ No Entry",
    }
    STATUS_GROUP_COLORS = {
        "In Office": "#f0fdf4",
        "In Office - AM Only": "#fffbeb",
        "Remote": "#eff6ff",
        "Out": "#fef2f2",
        "—": "#f9fafb",
    }

    for status_group in STATUS_ORDER:
        members = [n for n in TEAM if week_data.get(n, {}).get(day_name, "—") == status_group]
        if not members:
            continue

        if status_group == "In Office":
            css_class = "status-in"
        elif status_group == "In Office - AM Only":
            css_class = "status-am"
        elif status_group == "Remote":
            css_class = "status-remote"
        elif status_group == "Out":
            css_class = "status-out"
        else:
            css_class = "status-none"

        bg_color = STATUS_GROUP_COLORS.get(status_group, "#f9fafb")
        label = STATUS_GROUP_LABELS.get(status_group, status_group)

        st.markdown(f"""
        <div style="background:{bg_color}; padding:8px 12px; border-radius:6px 6px 0 0; margin-top:12px; border-bottom:2px solid #e5e7eb;">
            <span style="font-weight:600; font-size:0.95rem; color:#1f2937;">{label} ({len(members)})</span>
        </div>
        """, unsafe_allow_html=True)

        for name in members:
            desk = desk_bookings.get(name, "") or PERMANENT_DESKS.get(name, "")
            desk_text = f"&nbsp;&nbsp;📍 {desk}" if desk else ""

            st.markdown(f"""
            <div style="display:flex; align-items:center; padding:8px 12px; border-bottom:1px solid #f3f4f6; background:white;">
                <span style="flex:1; font-size:1.05rem; font-weight:500; color:#1f2937;">{name}</span>
                <span style="min-width:100px; color:#6b7280; font-size:0.9rem;">{desk_text}</span>
            </div>
            """, unsafe_allow_html=True)

    # Desk overview
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("🪑 Today's Desk Assignments")
    desk_cols = st.columns(min(len(DESKS), 8))
    for i, desk in enumerate(DESKS):
        # Check permanent assignment first
        perm_occupant = None
        for person, perm_desk in PERMANENT_DESKS.items():
            if perm_desk == desk:
                perm_occupant = person
                break

        occupant = None
        if not perm_occupant:
            for person, booked_desk in desk_bookings.items():
                if booked_desk == desk:
                    occupant = person
                    break

        with desk_cols[i]:
            if perm_occupant:
                st.markdown(f"""<div class="desk-taken" style="border-color:#7c3aed; background-color:#f5f3ff;">
                    <div style="font-weight:600;">{desk}</div>
                    <div style="color:#7c3aed; font-size:0.85rem;">🔒 {perm_occupant}</div>
                </div>""", unsafe_allow_html=True)
            elif occupant:
                st.markdown(f"""<div class="desk-taken">
                    <div style="font-weight:600;">{desk}</div>
                    <div style="color:#ef4444; font-size:0.85rem;">🔴 {occupant}</div>
                </div>""", unsafe_allow_html=True)
            else:
                st.markdown(f"""<div class="desk-available">
                    <div style="font-weight:600;">{desk}</div>
                    <div style="color:#22c55e; font-size:0.85rem;">✅ Available</div>
                </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: Book a Desk
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🪑 Book a Desk":
    st.header("🪑 Book a Desk")

    today = datetime.now().date()
    week_monday = get_current_week_monday()

    # Date selector — only show weekdays of current week
    available_dates = []
    for i in range(5):
        d = week_monday + timedelta(days=i)
        if d >= today:
            available_dates.append(d)

    if not available_dates:
        st.info("No more weekdays this week. Check back Monday!")
        st.stop()

    selected_date = st.selectbox(
        "Select date",
        available_dates,
        format_func=lambda d: d.strftime("%A, %B %d"),
    )

    who = st.selectbox("Who are you?", TEAM)

    # Load current bookings for the selected date
    bookings = load_desk_bookings(desk_ws, selected_date)

    # Show desk grid
    st.markdown("#### Available Desks")
    desk_cols = st.columns(min(len(DESKS), 8))

    taken_desks = {}
    for person, desk in bookings.items():
        taken_desks[desk] = person

    my_current_desk = bookings.get(who)

    # Check if user has a permanent desk
    has_permanent = who in PERMANENT_DESKS

    for i, desk in enumerate(DESKS):
        # Check permanent assignment
        perm_occupant = None
        for person, perm_desk in PERMANENT_DESKS.items():
            if perm_desk == desk:
                perm_occupant = person
                break

        occupant = taken_desks.get(desk) if not perm_occupant else None

        with desk_cols[i]:
            if perm_occupant:
                st.markdown(f"""<div class="desk-taken" style="border-color:#7c3aed; background-color:#f5f3ff;">
                    <div style="font-weight:600;">{desk}</div>
                    <div style="color:#7c3aed; font-size:0.85rem;">🔒 {perm_occupant}</div>
                </div>""", unsafe_allow_html=True)
            elif occupant == who:
                st.markdown(f"""<div class="desk-taken" style="border-color:#3b82f6; background-color:#eff6ff;">
                    <div style="font-weight:600;">{desk}</div>
                    <div style="color:#3b82f6; font-size:0.85rem;">📍 You</div>
                </div>""", unsafe_allow_html=True)
            elif occupant:
                st.markdown(f"""<div class="desk-taken">
                    <div style="font-weight:600;">{desk}</div>
                    <div style="color:#ef4444; font-size:0.85rem;">🔴 {occupant}</div>
                </div>""", unsafe_allow_html=True)
            else:
                st.markdown(f"""<div class="desk-available">
                    <div style="font-weight:600;">{desk}</div>
                    <div style="color:#22c55e; font-size:0.85rem;">✅ Open</div>
                </div>""", unsafe_allow_html=True)

    st.markdown("---")

    if has_permanent:
        st.info(f"You're permanently assigned to **{PERMANENT_DESKS[who]}** — no booking needed!")
    else:
        # Booking actions — only show bookable desks
        available_desks = [d for d in BOOKABLE_DESKS if d not in taken_desks or taken_desks[d] == who]

        col1, col2 = st.columns(2)
        with col1:
            selected_desk = st.selectbox(
                "Choose a desk",
                ["None"] + available_desks,
                index=(available_desks.index(my_current_desk) + 1) if my_current_desk in available_desks else 0,
            )

        with col2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("💾 Save Booking", type="primary", use_container_width=True):
                if selected_desk == "None" and my_current_desk:
                    clear_desk_booking(desk_ws, selected_date, who)
                    st.success(f"Cleared desk booking for {who} on {selected_date.strftime('%A')}.")
                elif selected_desk != "None":
                    save_desk_booking(desk_ws, selected_date, who, selected_desk)
                    st.success(f"✅ {who} booked **{selected_desk}** for {selected_date.strftime('%A, %B %d')}!")
                else:
                    st.info("No changes made.")
                st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: Update My Week
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📝 Update My Week":
    st.header("📝 Update My Week")

    # Choose which week to update
    week_option = st.radio(
        "Which week?",
        ["This Week", "Next Week"],
        horizontal=True,
    )
    week_monday = get_current_week_monday() if week_option == "This Week" else get_next_week_monday()

    st.caption(f"Week of {week_monday.strftime('%B %d, %Y')}")

    who = st.selectbox("Who are you?", TEAM)

    # Load existing data
    week_data = load_weekly_status(status_ws, week_monday)
    current = week_data.get(who, {day: "—" for day in DAYS})

    st.markdown("Set your status for each day:")

    selections = {}
    cols = st.columns(5)
    for i, day in enumerate(DAYS):
        with cols[i]:
            current_val = current.get(day, "—")
            default_idx = STATUSES.index(current_val) if current_val in STATUSES else 0
            selections[day] = st.selectbox(
                day,
                STATUSES,
                index=default_idx,
                key=f"status_{day}",
            )

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("💾 Save My Week", type="primary", use_container_width=True):
        save_weekly_status(status_ws, week_monday, who, selections)
        st.success(f"✅ Saved {who}'s schedule for the week of {week_monday.strftime('%B %d')}!")
        st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: Weekly View
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📅 Weekly View":
    st.header("📅 Weekly View")

    week_option = st.radio(
        "Which week?",
        ["This Week", "Next Week"],
        horizontal=True,
    )
    week_monday = get_current_week_monday() if week_option == "This Week" else get_next_week_monday()

    st.caption(f"Week of {week_monday.strftime('%B %d, %Y')}")

    week_data = load_weekly_status(status_ws, week_monday)

    # Build styled table
    def status_badge_html(status):
        if status == "In Office":
            return f'<span class="status-badge status-in">In Office</span>'
        elif status == "In Office - AM Only":
            return f'<span class="status-badge status-am">AM Only</span>'
        elif status == "Remote":
            return f'<span class="status-badge status-remote">Remote</span>'
        elif status == "Out":
            return f'<span class="status-badge status-out">Out</span>'
        else:
            return f'<span class="status-badge status-none">—</span>'

    today_name = get_today_day_name()

    # Table header with today highlighted
    header_cells = "<th style='padding:10px; text-align:left; color:#1f2937; border-bottom:2px solid #e5e7eb;'>Name</th>"
    for day in DAYS:
        highlight = "background-color: #eff6ff;" if day == today_name else ""
        header_cells += f"<th style='padding:10px; text-align:center; color:#1f2937; border-bottom:2px solid #e5e7eb; {highlight}'>{day}</th>"

    rows = ""
    for name in TEAM:
        row_cells = f"<td style='padding:8px 10px; font-weight:500; color:#1f2937; border-bottom:1px solid #f3f4f6;'>{name}</td>"
        for day in DAYS:
            status = week_data.get(name, {}).get(day, "—")
            highlight = "background-color: #f8faff;" if day == today_name else ""
            row_cells += f"<td style='padding:8px; text-align:center; border-bottom:1px solid #f3f4f6; {highlight}'>{status_badge_html(status)}</td>"
        rows += f"<tr>{row_cells}</tr>"

    st.markdown(f"""
    <table style="width:100%; border-collapse:collapse; background:white; border-radius:8px; overflow:hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.08);">
        <thead><tr>{header_cells}</tr></thead>
        <tbody>{rows}</tbody>
    </table>
    """, unsafe_allow_html=True)

    # Daily summaries
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("Daily Summaries")
    day_cols = st.columns(5)
    for i, day in enumerate(DAYS):
        with day_cols[i]:
            in_count = sum(1 for n in TEAM if week_data.get(n, {}).get(day, "—") == "In Office")
            am_count = sum(1 for n in TEAM if week_data.get(n, {}).get(day, "—") == "In Office - AM Only")
            remote_count = sum(1 for n in TEAM if week_data.get(n, {}).get(day, "—") == "Remote")
            out_count = sum(1 for n in TEAM if week_data.get(n, {}).get(day, "—") == "Out")

            is_today = "border: 2px solid #3b82f6;" if day == today_name else "border: 1px solid #e5e7eb;"
            st.markdown(f"""<div style="background:white; {is_today} border-radius:8px; padding:12px; text-align:center;">
                <div style="font-weight:600; margin-bottom:8px;">{day[:3]}</div>
                <div style="color:#22c55e;">🟢 {in_count}</div>
                <div style="color:#f59e0b;">🟡 {am_count}</div>
                <div style="color:#3b82f6;">🔵 {remote_count}</div>
                <div style="color:#ef4444;">🔴 {out_count}</div>
            </div>""", unsafe_allow_html=True)
