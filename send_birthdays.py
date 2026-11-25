#!/usr/bin/env python3
"""
send_birthdays.py

This script automates staff birthday reminders.

In simple terms, it does three big things:
1) Gets birthday data (from SharePoint/Google Sheets if available, otherwise from a local CSV file)
2) Figures out who has birthdays in the reminder windows (10, 5, 2, or 0 days from today)
3) Sends HR an email summary
"""
# Built-in CSV reader for comma-separated files.
import csv
# Date/time tools for parsing dates and calculating day differences.
import datetime
# Allows reading byte data like a file in memory.
import io
# Access environment variables and file operations.
import os
# Regular expressions for cleaning date text.
import re
# SMTP library used to send email.
import smtplib
# System utilities (used here for clean program exits).
import sys
# Creates temporary files safely.
import tempfile
# Dictionary variant that auto-creates default values (lists in our case).
from collections import defaultdict
# Easy way to build an email message object.
from email.message import EmailMessage

# Third-party HTTP library to download files from URLs.
import requests

# Try to load timezone support (available in modern Python versions).
try:
    from zoneinfo import ZoneInfo
except Exception:
    # If zoneinfo is unavailable, we will fall back to UTC later.
    ZoneInfo = None

# Try to load pandas for reading Excel files.
try:
    import pandas as pd
except Exception:
    # If pandas/openpyxl are missing, Excel reading will fail clearly.
    pd = None

# Default path to local fallback birthday data.
CSV_PATH = "birthdays.csv"
# Default HR recipient if HR_EMAIL env var is not provided.
DEFAULT_HR_EMAIL = "nofiumoruf17@gmail.com"
# Gmail SMTP server host.
SMTP_HOST = "smtp.gmail.com"
# SSL port for Gmail SMTP.
SMTP_PORT_SSL = 465

# Regex to remove ordinal suffixes in dates (e.g., "21st" -> "21").
ORDINAL_RE = re.compile(r"(\d+)(st|nd|rd|th)", flags=re.IGNORECASE)
# Allowed reminder windows in days: 10 days, 5 days, 2 days, and same day.
VALID_WINDOWS = {10, 5, 2, 0}


def normalize_date_str(s: str) -> str:
    # If input is empty/None, return empty string early.
    if not s:
        return ""
    # Remove spaces at the beginning/end.
    s = s.strip()
    # Replace ordinals like "3rd" with "3" so parsing is easier.
    s = ORDINAL_RE.sub(r"\1", s)
    # Replace commas with spaces (e.g., "May, 5" -> "May 5").
    s = s.replace(",", " ")
    # Convert "5th of May" style into "5 May".
    s = re.sub(r"\s+of\s+", " ", s, flags=re.IGNORECASE)
    # Collapse repeated spaces into a single space.
    s = " ".join(s.split())
    # Convert to title case for predictable parsing (e.g., "may" -> "May").
    return s.title()


def parse_day_month(s: str):
    # Standardize the incoming date text first.
    s = normalize_date_str(s)
    # Try several common date text patterns.
    for fmt in ("%d %B", "%d %b", "%B %d", "%b %d", "%d-%B", "%d/%B"):
        try:
            # Parse text to a datetime object using the current format pattern.
            dt = datetime.datetime.strptime(s, fmt)
            # Return only day and month (year is irrelevant for birthdays).
            return dt.day, dt.month
        except Exception:
            # If this format fails, try the next one.
            continue
    # Last-chance fallback: manually extract "number + month-word".
    m = re.search(r"(\d{1,2})\s+([A-Za-z]+)", s)
    if m:
        # First group is day.
        day = int(m.group(1))
        # Second group is month name text.
        mon = m.group(2)
        try:
            # Try full month name first (e.g., "September").
            dt = datetime.datetime.strptime(mon.title(), "%B")
            return day, dt.month
        except Exception:
            try:
                # Then try abbreviated month name (e.g., "Sep").
                dt = datetime.datetime.strptime(mon.title(), "%b")
                return day, dt.month
            except Exception:
                # Could not parse month.
                return None
    # No valid date pattern found.
    return None


def today_date():
    # Read timezone from environment; default to Africa/Lagos.
    tz_name = os.environ.get("TZ", "Africa/Lagos")
    # If timezone support exists, try timezone-aware current date.
    if ZoneInfo:
        try:
            return datetime.datetime.now(ZoneInfo(tz_name)).date()
        except Exception:
            # If timezone lookup fails, fall back to UTC date.
            return datetime.datetime.utcnow().date()
    # If zoneinfo is unavailable at all, use UTC date.
    return datetime.datetime.utcnow().date()


def normalize_workbook_url(url: str) -> str:
    # Convert a Google Sheets share URL into a direct downloadable Excel export URL.
    if "docs.google.com" in url and "/spreadsheets/d/" in url:
        m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
        if not m:
            raise RuntimeError("Google Sheets URL is missing a spreadsheet id.")
        sheet_id = m.group(1)
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    # Non-Google URLs are returned unchanged (e.g., SharePoint links).
    return url


def download_sharepoint_excel(url: str, timeout: int = 30):
    # Guard against missing/empty URL input.
    if not url:
        raise ValueError("SHAREPOINT_URL is empty")

    # Normalize known share links (Google Sheets) to direct file-download URLs.
    resolved_url = normalize_workbook_url(url)
    print("Resolved workbook URL:", resolved_url)

    # Attempt direct HTTP download.
    response = requests.get(resolved_url, allow_redirects=True, timeout=timeout)
    # Print request diagnostics to help troubleshooting.
    print("Workbook HTTP status:", response.status_code)
    print("Workbook Content-Type:", response.headers.get("Content-Type"))
    print("Workbook final URL:", response.url)

    # Any non-200 response means the file was not retrieved successfully.
    if response.status_code != 200:
        raise RuntimeError(f"SharePoint URL returned status {response.status_code}")

    # Inspect content type to avoid treating HTML/login pages as Excel.
    content_type = response.headers.get("Content-Type", "").lower()
    if "html" in content_type or "text/plain" in content_type:
        raise RuntimeError(
            "URL did not return an Excel file. It likely returned an HTML login page or sharing landing page."
        )

    # Warn if content type looks unusual, but continue in case server is generic.
    if "excel" not in content_type and "spreadsheetml" not in content_type and "octet-stream" not in content_type and "xlsx" not in content_type and "zip" not in content_type:
        # Still allow it in case the file is valid but headers are not descriptive.
        print("Warning: unexpected content type for an Excel file. Continuing anyway.")

    # Save downloaded bytes into a temporary .xlsx file.
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(response.content)
        tmp_path = tmp.name

    # Print where the file was stored.
    print(f"Saved workbook to: {tmp_path}")
    # Return temporary file path to caller.
    return tmp_path


def read_birthdays_from_csv(csv_path=CSV_PATH):
    # This will hold cleaned row objects.
    rows = []
    # Open CSV using UTF-8 so names/symbols are read correctly.
    with open(csv_path, newline="", encoding="utf-8") as f:
        # Read each row as a dictionary using CSV headers.
        reader = csv.DictReader(f)
        for r in reader:
            # Normalize CSV columns into consistent internal keys.
            rows.append({
                # Try common variants of serial-number header.
                "sn": r.get("S/N") or r.get("S/N ") or r.get("SN"),
                # Staff name, trimmed.
                "name": (r.get("NAMES") or "").strip(),
                # Staff designation/title, trimmed.
                "designation": (r.get("DESIGNATION") or "").strip(),
                # Raw date-of-birth text, trimmed.
                "dob": (r.get("DATE OF BIRTH") or "").strip()
            })
    # Return normalized list of staff rows.
    return rows


def read_birthdays_from_excel(path: str):
    # Excel parsing requires pandas + openpyxl.
    if pd is None:
        raise RuntimeError("pandas/openpyxl not installed")

    # Read file bytes first to avoid direct file-handle edge cases.
    with open(path, "rb") as f:
        data = f.read()

    # Load the Excel sheet into a pandas DataFrame as strings.
    df = pd.read_excel(io.BytesIO(data), engine="openpyxl", dtype=str, keep_default_na=False)
    # Log detected columns for visibility/debugging.
    print("Excel columns detected:", list(df.columns))

    # Build map: lowercase/trimmed name -> original actual column label.
    normalized = {str(c).strip().lower(): c for c in df.columns}

    # Helper to find the first matching column among possible aliases.
    def find_col(*names):
        for n in names:
            key = str(n).strip().lower()
            if key in normalized:
                return normalized[key]
        return None

    # Resolve likely column names even if workbook header wording differs.
    sn_col = find_col("s/n", "sn", "s n", "serial")
    name_col = find_col("names", "name", "full name", "employee name")
    desig_col = find_col("designation", "role", "job title", "position")
    dob_col = find_col("date of birth", "dob", "date", "birth date")

    # Convert each DataFrame row to our standard dictionary structure.
    rows = []
    for _, row in df.iterrows():
        rows.append({
            # If column not found, keep field empty to avoid crashes.
            "sn": "" if sn_col is None else str(row.get(sn_col, "")).strip(),
            "name": "" if name_col is None else str(row.get(name_col, "")).strip(),
            "designation": "" if desig_col is None else str(row.get(desig_col, "")).strip(),
            "dob": "" if dob_col is None else str(row.get(dob_col, "")).strip()
        })
    # Return normalized rows from workbook.
    return rows


def days_until_next_birthday(day: int, month: int, from_date: datetime.date):
    # Start by assuming birthday is in the current year.
    year = from_date.year
    try:
        # Create birthday date in current year.
        next_bd = datetime.date(year, month, day)
    except ValueError:
        # Invalid day/month combination (e.g., impossible date).
        return None
    # If birthday already passed this year, move to next year.
    if next_bd < from_date:
        try:
            next_bd = datetime.date(year + 1, month, day)
        except ValueError:
            # Invalid date in next year as well.
            return None
    # Return difference in whole days.
    return (next_bd - from_date).days


def build_email_body(grouped: dict, for_date: datetime.date):
    # Build the email as a list of lines, then join at the end.
    lines = []
    # Greeting line.
    lines.append("Hello HR team,")
    # Blank line for readability.
    lines.append("")
    # Explain what date this reminder is generated for.
    lines.append(f"This is an automated birthday reminder for {for_date.strftime('%d %B %Y')}.")
    lines.append("")
    # Count total people across all reminder groups.
    total = sum(len(v) for v in grouped.values())
    # Handle case where no birthdays matched.
    if total == 0:
        lines.append("There are no upcoming birthdays in the configured windows (10, 5, 2, 0 days).")
        lines.append("")
        lines.append("Regards,")
        lines.append("Birthday Automation")
        return "\n".join(lines)

    # Intro line when there are matching birthdays.
    lines.append(f"There are {total} staff with birthdays in the next configured windows:")
    lines.append("")

    # Print groups in order: 10, 5, 2, 0.
    for window in sorted(VALID_WINDOWS, reverse=True):
        # Fetch records in this window; default empty list if absent.
        items = grouped.get(window, [])
        if not items:
            # Skip empty sections.
            continue
        # Human-friendly section title.
        header = "Same day" if window == 0 else f"{window} day{'s' if window != 1 else ''} remaining"
        lines.append(header + ":")
        for m in items:
            # Re-parse DOB text for display of exact upcoming date.
            parsed = parse_day_month(m["dob"])
            if parsed:
                d, mo = parsed
                # Recompute days from current run date.
                days = days_until_next_birthday(d, mo, for_date)
                if days is None:
                    # If date invalid, show original DOB text as fallback.
                    nb_str = m["dob"]
                else:
                    # Convert day-difference to concrete calendar date.
                    nb_date = for_date + datetime.timedelta(days=days)
                    nb_str = nb_date.strftime("%d %B %Y")
            else:
                # If parsing fails entirely, keep original DOB text.
                nb_str = m["dob"]
            # Add one bullet for each staff member in this window.
            lines.append(f"- {m['name']} — {m['designation']} (DOB: {m['dob']}) — Next: {nb_str}")
        # Blank line between sections.
        lines.append("")

    # Friendly closing lines.
    lines.append("Please join in wishing them well.")
    lines.append("")
    lines.append("Regards,")
    lines.append("Birthday Automation")
    # Convert lines list into the final plain-text email body.
    return "\n".join(lines)


def send_email(subject, body, from_addr, to_addrs, smtp_user, smtp_pass):
    # Create an email container object.
    msg = EmailMessage()
    # Set email subject line.
    msg["Subject"] = subject
    # Set sender address.
    msg["From"] = from_addr
    # Support one recipient or a list/tuple of recipients.
    msg["To"] = ", ".join(to_addrs) if isinstance(to_addrs, (list, tuple)) else to_addrs
    # Set plain-text body content.
    msg.set_content(body)

    # Open secure SMTP connection, authenticate, then send.
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT_SSL, timeout=30) as server:
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)


def main():
    # Read required Gmail credentials from environment variables.
    gmail_user = os.environ.get("GMAIL_USER")
    gmail_pass = os.environ.get("GMAIL_APP_PASSWORD")
    # Read optional HR target; use default if not provided.
    hr_email = os.environ.get("HR_EMAIL", DEFAULT_HR_EMAIL)
    # Optional SharePoint workbook URL.
    sharepoint_url = os.environ.get("SHAREPOINT_URL")
    # Optional Google Sheets sharing URL.
    google_sheet_url = os.environ.get("GOOGLE_SHEET_URL")

    # If email credentials are missing, stop gracefully.
    if not gmail_user or not gmail_pass:
        print("Missing GMAIL_USER or GMAIL_APP_PASSWORD environment variables. Exiting.")
        sys.exit(0)

    # Will hold birthday records once loaded.
    rows = []

    # Prefer remote workbook source when URL is provided.
    source_url = sharepoint_url or google_sheet_url
    source_label = "SHAREPOINT_URL" if sharepoint_url else "GOOGLE_SHEET_URL"
    if source_url:
        print(f"Attempting to use {source_label}")
        try:
            # Download workbook file from the configured source URL.
            workbook_path = download_sharepoint_excel(source_url)
            # Parse workbook rows into normalized structure.
            rows = read_birthdays_from_excel(workbook_path)
            print(f"Parsed {len(rows)} rows from workbook source.")
            # Clean up temporary file after successful parse.
            os.remove(workbook_path)
        except Exception as e:
            # If workbook source fails for any reason, fallback to CSV.
            print("Excel download/parse failed:", str(e))
            print("Falling back to local CSV...")
            try:
                rows = read_birthdays_from_csv(CSV_PATH)
                print(f"Parsed {len(rows)} rows from CSV fallback.")
            except FileNotFoundError:
                # Neither workbook URL nor CSV exists.
                print(f"{CSV_PATH} not found. Exiting.")
                sys.exit(1)
    else:
        # No SharePoint URL; read directly from local CSV.
        try:
            rows = read_birthdays_from_csv(CSV_PATH)
            print(f"Parsed {len(rows)} rows from CSV.")
        except FileNotFoundError:
            # No data source available at all.
            print(f"{CSV_PATH} not found and no remote workbook URL is set. Exiting.")
            sys.exit(1)

    # Determine "today" date using timezone logic.
    today = today_date()
    # Group matching staff records by day-window value (10/5/2/0).
    grouped = defaultdict(list)

    # Process every staff record.
    for r in rows:
        # Extract and sanitize DOB string.
        dob = (r.get("dob") or "").strip()
        # Convert DOB text into numeric day/month.
        parsed = parse_day_month(dob)
        if not parsed:
            # Skip rows with unparseable DOB.
            continue
        day, month = parsed
        # Compute days until next birthday from today.
        delta = days_until_next_birthday(day, month, today)
        # Keep only people in configured reminder windows.
        if delta in VALID_WINDOWS:
            grouped[delta].append(r)

    # If no one matches, print message and end without sending email.
    if not any(grouped.values()):
        print(f"No birthdays in the configured windows for {today.isoformat()}.")
        return

    # Build short subject fragments like "2 x 10d".
    subject_parts = []
    for w in sorted(grouped.keys()):
        label = "Same day" if w == 0 else f"{w}d"
        subject_parts.append(f"{len(grouped[w])} x {label}")

    # Final email subject line with summary and date.
    subject = f"Birthday Reminder — {' | '.join(subject_parts)} — {today.strftime('%d %b %Y')}"
    # Generate detailed email body.
    body = build_email_body(grouped, today)

    try:
        # Send email to HR using Gmail credentials.
        send_email(subject, body, gmail_user, [hr_email], gmail_user, gmail_pass)
        print(f"Sent reminder to {hr_email}")
    except Exception as e:
        # Print the error and re-raise so automation can mark failure.
        print("Failed to send email:", str(e))
        raise


# Standard Python entrypoint guard:
# run main() only when this file is executed directly.
if __name__ == "__main__":
    main()
