import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# ========== CONFIG ==========
SOURCE_SHEET = "Attendes-Eventbrite"
SOURCE_TABS = ["Corporate Wellbeing", "Southampton", "Business Revival Series","Milton Keynes Expo"]



TARGET_SHEET = "Expo-Sales-Management"
EXHIBITOR_TAB = "exhibitors-1"
SPEAKER_TAB = "speakers-2"
VISITOR_TAB = "visitors"


SHOW_MAPPING = {
    "Corporate Wellbeing": "The Corporate Wellbeing Expo",
    "Southampton": "Southampton Expo 29 Jan 2026",
    "Milton Keynes Expo": "Milton Keynes Expo 2026"
}

# ---- META SOURCE ----
META_SOURCE_SHEET = "Meta Leads"

# ---- ACTIVE META EXHIBITOR ADS ----
ACTIVE_META_EXHIBITOR_TABS = {
    "MKE-EXIBITORS 29 APRIL 2026-2": "Milton Keynes Expo 2026",
    "Bournemouth - Exhibitor 2026": "Bournemouth Expo 2026",
    "Milton Keynes | Expo'26 - Exhibitor": "Milton Keynes Expo 2026",
}

# === Google Auth ===
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
CREDS = Credentials.from_service_account_file("/etc/secrets/service.json", scopes=SCOPES)
client = gspread.authorize(CREDS)

today = datetime.now().strftime("%Y-%m-%d")

# ---------------- UTILITY ----------------
def get_column_map(header_row):
    return {header.strip(): idx for idx, header in enumerate(header_row) if header.strip() != ""}

def safe_get(row, idx, default=""):
    if idx is None:
        return default
    if idx >= len(row):
        return default
    return row[idx].strip()


def split_full_name(full):
    if not full:
        return "", ""
    parts = full.strip().split()
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])





# ---------------- EVENTBRITE BUILD ----------------
def build_eventbrite_row(mapping, data, mode):
    new = [""] * len(mapping)

    new[mapping["Lead Date"]] = today
    new[mapping["Lead Source"]] = "Eventbrite"
    new[mapping["First_Name"]] = data.get("first")
    new[mapping["Last Name"]] = data.get("last")
    new[mapping["Company Name"]] = data.get("company")
    new[mapping["Mobile"]] = data.get("mobile")
    new[mapping["Email"]] = data.get("email")
    new[mapping["Show"]] = data.get("show")

    if mode == "exhibitor":
        new[mapping["Interested for"]] = "Exhibitor"
    else:
        new[mapping["Interested In?"]] = "Speaker"

    return new


# ---------------- META BUILDERS ----------------

def build_meta_row(mapping, data, lead_source="Meta Lead"):
    new = [""] * len(mapping)

    if "Lead Date" in mapping:
        new[mapping["Lead Date"]] = today

    if "Lead Source" in mapping:
        new[mapping["Lead Source"]] = lead_source

    if "First_Name" in mapping:
        new[mapping["First_Name"]] = data.get("first", "")

    if "Last Name" in mapping:
        new[mapping["Last Name"]] = data.get("last", "")

    if "Company Name" in mapping:
        new[mapping["Company Name"]] = data.get("company", "")

    if "Mobile" in mapping:
        new[mapping["Mobile"]] = data.get("mobile", "")

    if "Email" in mapping:
        new[mapping["Email"]] = data.get("email", "")

    if "Show" in mapping:
        new[mapping["Show"]] = data.get("show", "")

    if "Interested for" in mapping:
        new[mapping["Interested for"]] = "Exhibitor Opportunity"

    return new

# ---------------- LOAD EXISTING EMAILS ----------------
def load_existing_emails():
    expo_sheet = client.open(TARGET_SHEET)

    exhibitor_ws = expo_sheet.worksheet(EXHIBITOR_TAB)
    speaker_ws = expo_sheet.worksheet(SPEAKER_TAB)

    exhibitor_map = get_column_map(exhibitor_ws.row_values(1))
    speaker_map = get_column_map(speaker_ws.row_values(1))

    ex_rows = exhibitor_ws.get_all_values()
    sp_rows = speaker_ws.get_all_values()

    visitor_ws = expo_sheet.worksheet(VISITOR_TAB)
    visitor_map = get_column_map(visitor_ws.row_values(1))
    visitor_rows = visitor_ws.get_all_values()

    exhibitor_existing = {row[exhibitor_map["Email"]].lower().strip() for row in ex_rows[1:] if row[exhibitor_map["Email"]]}
    speaker_existing = {row[speaker_map["Email"]].lower().strip() for row in sp_rows[1:] if row[speaker_map["Email"]]}
    visitor_existing = {row[visitor_map["Email"]].lower().strip() for row in visitor_rows[1:] if row[visitor_map["Email"]]}

    return exhibitor_existing, speaker_existing, visitor_existing

# ---------------- PROCESS 1: EVENTBRITE ----------------
def process_eventbrite(tab_name, exhibitor_batch, speaker_batch, existing_exhibitors, existing_speakers):
    sheet = client.open(SOURCE_SHEET).worksheet(tab_name)

    rows = sheet.get_all_values()
    headers = rows[0]
    col = get_column_map(headers)

    col_first = col.get("First Name")
    col_last = col.get("Surname")
    col_company = col.get("Company")
    col_mobile = col.get("Mobile Phone")
    col_email = col.get("Email")
    col_speaker = col.get("Do you want to be a Speaker?")
    col_exhibitor = col.get("Do you want to be an Exhibitor?")
    col_campaign = col.get("Campaign")

    show_name = SHOW_MAPPING.get(tab_name, None)


    # Normalize tab name
    normalized = tab_name.lower().replace(" ", "").strip()
    is_business_revival = ("businessrevival" in normalized)

    for row in rows[1:]:
        try:
            email = row[col_email].lower().strip()
        except:
            continue
        if not email:
            continue

        # --- FIXED LOGIC FOR SHOW ---
        if is_business_revival and col_campaign is not None:
            show_value = row[col_campaign].strip()
        else:
            show_value = show_name if show_name is not None else ""
        # -----------------------------

        data = {
            "first": row[col_first],
            "last": row[col_last],
            "company": row[col_company],
            "mobile": row[col_mobile],
            "email": email,
            "show": show_value
        }

        want_exhibitor = row[col_exhibitor].lower().strip() == "yes"
        want_speaker = row[col_speaker].lower().strip() == "yes"

        if want_exhibitor and email not in existing_exhibitors:
            exhibitor_batch.append(data)
            existing_exhibitors.add(email)

        if want_speaker and email not in existing_speakers:
            speaker_batch.append(data)
            existing_speakers.add(email)

# ---------------- PROCESS 2: ACTIVE META EXHIBITOR ADS ----------------
def process_active_meta_exhibitor_ads(exh_batch, existing_exh):
    sheet = client.open(META_SOURCE_SHEET)

    for tab_name, show_name in ACTIVE_META_EXHIBITOR_TABS.items():
        ws = sheet.worksheet(tab_name)
        rows = ws.get_all_values()

        if not rows:
            continue

        col = get_column_map(rows[0])

        # Common fields
        idx_full = col.get("full_name")
        idx_email = col.get("work_email")
        idx_phone = col.get("phone_number")
        idx_company = col.get("company_name")

        # Alternate field names for new tab
        if idx_email is None:
            idx_email = col.get("email")
        if idx_phone is None:
            idx_phone = col.get("phone")
        if idx_company is None:
            idx_company = col.get("website")

        # Optional old-tab extra fields
        idx_created_time = col.get("created_time")
        idx_budget = col.get("what_is_your_budget_for_exhibition_?")
        idx_call_window = col.get("preffered_time_to_connect_over_phone_call?")
        idx_lead_status = col.get("lead_status")
        idx_participation = col.get("how_do_you_want_to_participate_in_b2b_growth_expo?")

        for row in rows[1:]:
            email = safe_get(row, idx_email).lower().strip()

            if not email or email in existing_exh:
                continue

            first, last = split_full_name(safe_get(row, idx_full))

            exh_batch.append({
                "first": first,
                "last": last,
                "company": safe_get(row, idx_company),   # website will go into Company Name for this new tab
                "mobile": safe_get(row, idx_phone),
                "email": email,
                "show": show_name,
                "interest": "Exhibitor",
                "budget": safe_get(row, idx_budget),
                "lead_status": safe_get(row, idx_lead_status),
                "preferred_call_window": safe_get(row, idx_call_window),
                "lead_time": safe_get(row, idx_created_time),
                "enquiry_type": safe_get(row, idx_participation),
            })

            existing_exh.add(email)

# ---------------- INSERTION ----------------
def batch_insert_rows(ws_name, rows, mode="exhibitor", lead_source="Eventbrite"):
    if not rows:
        return

    expo_sheet = client.open(TARGET_SHEET)
    ws = expo_sheet.worksheet(ws_name)

    header = ws.row_values(1)
    mapping = get_column_map(header)

    if lead_source == "Eventbrite":
        formatted = [build_eventbrite_row(mapping, r, mode) for r in rows]
    else:
        formatted = [build_meta_row(mapping, r, lead_source) for r in rows]

    blank = [[""] * len(header) for _ in range(len(rows))]
    ws.insert_rows(blank, row=2)
    ws.update("A2", formatted, value_input_option="USER_ENTERED")

# ---------------- RUN ALL ----------------
def run_all():
    print("\n=== Combined Processing Started ===")

    exh_existing, spk_existing, vis_existing = load_existing_emails()

    event_exh = []
    event_spk = []
    active_meta_exh = []

    # ---- Process Eventbrite ----
    for tab in SOURCE_TABS:
        process_eventbrite(tab, event_exh, event_spk, exh_existing, spk_existing)

    print("\nEventbrite Summary:")
    print(" Exhibitors:", len(event_exh))
    print(" Speakers:", len(event_spk))

    # ---- Process Active Meta Exhibitor Ads ----
    process_active_meta_exhibitor_ads(active_meta_exh, exh_existing)

    print("\nActive Meta Exhibitor Ads Summary:")
    print(" Exhibitors:", len(active_meta_exh))

    # ---- Insert all ----
    if event_exh:
        batch_insert_rows(EXHIBITOR_TAB, event_exh, "exhibitor", "Eventbrite")

    if event_spk:
        batch_insert_rows(SPEAKER_TAB, event_spk, "speaker", "Eventbrite")

    if active_meta_exh:
        batch_insert_rows(EXHIBITOR_TAB, active_meta_exh, "exhibitor", "Meta Lead – Active Ads")

    print("\n=== All Processing Completed ===\n")


if __name__ == "__main__":
    run_all()
