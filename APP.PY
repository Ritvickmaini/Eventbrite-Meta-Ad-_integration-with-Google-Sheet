import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# ========== CONFIG ==========
SOURCE_SHEET = "Attendes-Eventbrite"
SOURCE_TABS = ["Corporate Wellbeing", "Southampton", "Business Revival Series"]

TARGET_SHEET = "Expo-Sales-Management"
EXHIBITOR_TAB = "exhibitors-1"
SPEAKER_TAB = "speakers-2"
VISITOR_TAB = "visitors"


SHOW_MAPPING = {
    "Corporate Wellbeing": "The Corporate Wellbeing Expo",
    "Southampton": "Southampton Expo 29 Jan 2026"
}

# ---- META SOURCES ----
META_SOURCE_SHEET = "Meta Leads"
META_STANDARD_TAB = "Southampton expo exhibitor 29 Jan 2026"   # Process 2
META_CAMPAIGN_TAB = "Southamapton Leads Campaign"              # Process 3
META_SHOW_NAME = "Southampton Expo 29 Jan 2026"
# ---- META LONDON CAMPAIGN ----
META_LONDON_TAB = "london Leads Ad Set"
META_LONDON_SHOW_NAME = "London Expo"

# ---- META LEAD BASED ADS (PROCESS 5) ----
META_LEAD_BASED_TABS = [
    "BGE-Lead Based AD-09/12/2025",
    "SME-BE-Lead Based AD-09/12/2025",
    "FTE-Lead Based AD-09/12/2025",
    "CTE-Lead Based AD-09/12/2025",
    "BIE-Lead Based AD-09/12/2025",
    "CWE-Lead Based AD-09/12/2025",
    "Southampton Expo Exhibitor Leads"
]




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


def split_full_name(full):
    if not full:
        return "", ""
    parts = full.strip().split()
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])

def extract_expo_name(form_name):
    """
    Example:
    Business Growth Expo - Exhibitor
    → Business Growth Expo
    """
    if not form_name:
        return ""
    return form_name.split("-")[0].strip()



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
    new[mapping["Lead Date"]] = today
    new[mapping["Lead Source"]] = lead_source
    new[mapping["First_Name"]] = data.get("first")
    new[mapping["Last Name"]] = data.get("last")
    new[mapping["Company Name"]] = data.get("company")
    new[mapping["Mobile"]] = data.get("mobile")
    new[mapping["Email"]] = data.get("email")
    new[mapping["Show"]] = data.get("show", META_SHOW_NAME)

    if "Interested for" in mapping:
        new[mapping["Interested for"]] = data.get("interest", "")

    if "Lead-Budget" in mapping:
        new[mapping["Lead-Budget"]] = data.get("budget", "")

    return new

def build_visitor_row(mapping, data, lead_source):
    new = [""] * len(mapping)

    new[mapping["Lead Date"]] = today
    new[mapping["Lead Source"]] = lead_source
    new[mapping["First_Name"]] = data.get("first")
    new[mapping["Last Name"]] = data.get("last")
    new[mapping["Company"]] = data.get("company")
    new[mapping["Designation"]] = data.get("designation", "")
    new[mapping["Interested In?"]] = "Visitor"
    new[mapping["Comments"]] = ""
    new[mapping["Mobile"]] = data.get("mobile")
    new[mapping["Email"]] = data.get("email")
    new[mapping["Show"]] = data.get("show", META_SHOW_NAME)

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

# ---------------- PROCESS 2: META STANDARD LEADS ----------------
def process_meta_standard(exhibitor_batch, existing_exhibitors):
    sheet = client.open(META_SOURCE_SHEET).worksheet(META_STANDARD_TAB)
    rows = sheet.get_all_values()
    col = get_column_map(rows[0])

    idx_full = col.get("full_name")
    idx_company = col.get("provide_your_business_name")
    idx_phone = col.get("phone_number")
    idx_email = col.get("email")
    idx_budget = col.get("please_specify_your_budget")

    for row in rows[1:]:
        try:
            email = row[idx_email].lower().strip()
        except:
            continue

        if not email or email in existing_exhibitors:
            continue

        first, last = split_full_name(row[idx_full])

        exhibitor_batch.append({
            "first": first,
            "last": last,
            "company": row[idx_company],
            "mobile": row[idx_phone],
            "email": email,
            "budget": row[idx_budget]
        })

        existing_exhibitors.add(email)


# ---------------- PROCESS 3: META CAMPAIGN LEADS ----------------
def process_meta_campaign(exh_batch, spk_batch, visitor_batch, existing_exh, existing_spk, vis_existing):

    sheet = client.open(META_SOURCE_SHEET).worksheet(META_CAMPAIGN_TAB)
    rows = sheet.get_all_values()
    col = get_column_map(rows[0])

    idx_full = col.get("full_name")
    idx_company = col.get("provide_your_business_name")
    idx_phone = col.get("phone_number")
    idx_email = col.get("email")

    idx_exhibitor = col.get("do_you_want_to_be_an_exhibitor?")
    idx_speaker = col.get("do_you_want_to_be_a_speaker?")

    for row in rows[1:]:
        try:
            email = row[idx_email].strip().lower()
        except:
            continue

        if not email:
            continue

        first, last = split_full_name(row[idx_full])

        # Exhibitor
        if row[idx_exhibitor].strip().lower() == "yes" and email not in existing_exh:
            exh_batch.append({
                "first": first,
                "last": last,
                "company": row[idx_company],
                "mobile": row[idx_phone],
                "email": email,
                "interest": "Exhibitor"
            })
            existing_exh.add(email)

        # Speaker
        if row[idx_speaker].strip().lower() == "yes" and email not in existing_spk:
            spk_batch.append({
                "first": first,
                "last": last,
                "company": row[idx_company],
                "mobile": row[idx_phone],
                "email": email,
                "interest": "Speaker"
            })
            existing_spk.add(email)

        # Visitor (both NO)
        if (
            row[idx_exhibitor].strip().lower() == "no"
            and row[idx_speaker].strip().lower() == "no"
            and email not in vis_existing
        ):
            visitor_batch.append({
                "first": first,
                "last": last,
                "company": row[idx_company],
                "mobile": row[idx_phone],
                "email": email,
                "show": META_SHOW_NAME
            })
            vis_existing.add(email)



# ---------------- PROCESS 4: META LONDON CAMPAIGN LEADS ----------------
def process_meta_london_campaign(exh_batch, spk_batch, visitor_batch, existing_exh, existing_spk, vis_existing):

    sheet = client.open(META_SOURCE_SHEET).worksheet(META_LONDON_TAB)
    rows = sheet.get_all_values()
    col = get_column_map(rows[0])

    idx_full = col.get("full_name")
    idx_company = col.get("provide_your_business_name")
    idx_phone = col.get("phone_number")
    idx_email = col.get("email")

    idx_exhibitor = col.get("do_you_want_to_be_an_exhibitor?")
    idx_speaker = col.get("do_you_want_to_be_a_speaker?")

    for row in rows[1:]:
        try:
            email = row[idx_email].strip().lower()
        except:
            continue

        if not email:
            continue

        first, last = split_full_name(row[idx_full])

        # Exhibitor
        if row[idx_exhibitor].strip().lower() == "yes" and email not in existing_exh:
            exh_batch.append({
                "first": first,
                "last": last,
                "company": row[idx_company],
                "mobile": row[idx_phone],
                "email": email,
                "interest": "Exhibitor",
                "show": META_LONDON_SHOW_NAME
            })
            existing_exh.add(email)

        # Speaker
        if row[idx_speaker].strip().lower() == "yes" and email not in existing_spk:
            spk_batch.append({
                "first": first,
                "last": last,
                "company": row[idx_company],
                "mobile": row[idx_phone],
                "email": email,
                "interest": "Speaker",
                "show": META_LONDON_SHOW_NAME
            })
            existing_spk.add(email)
        
        # Visitor (both NO)
        if (
            row[idx_exhibitor].strip().lower() == "no"
            and row[idx_speaker].strip().lower() == "no"
            and email not in vis_existing
        ):
            visitor_batch.append({
                "first": first,
                "last": last,
                "company": row[idx_company],
                "mobile": row[idx_phone],
                "email": email,
                "show": META_LONDON_SHOW_NAME
            })
            vis_existing.add(email)


# ---------------- PROCESS 5: META LEAD BASED ADS ----------------
def process_meta_lead_based_ads(exh_batch, existing_exh):
    sheet = client.open(META_SOURCE_SHEET)

    for tab_name in META_LEAD_BASED_TABS:
        ws = sheet.worksheet(tab_name)
        rows = ws.get_all_values()
        col = get_column_map(rows[0])

        idx_full = col.get("full_name")
        idx_company = col.get("company_name")
        idx_phone = col.get("phone") or col.get("phone_number")
        idx_email = col.get("email") or col.get("work_email")
        idx_form = col.get("form_name")

        for row in rows[1:]:
            try:
                email = row[idx_email].strip().lower()
            except:
                continue

            if not email or email in existing_exh:
                continue

            first, last = split_full_name(row[idx_full])
            show_name = extract_expo_name(row[idx_form])

            exh_batch.append({
                "first": first,
                "last": last,
                "company": row[idx_company],
                "mobile": row[idx_phone],
                "email": email,
                "show": show_name,
                "interest": "Exhibitor"
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

    # -----------------------------------------------
    # SAFE BUILDER SWITCH
    # EVENTBRITE rows use build_eventbrite_row()
    # META rows use build_meta_row()
    # -----------------------------------------------
    if lead_source == "Eventbrite":
        formatted = [build_eventbrite_row(mapping, r, mode) for r in rows]
    elif mode == "visitor":
        formatted = [build_visitor_row(mapping, r, lead_source) for r in rows]
    else:
        formatted = [build_meta_row(mapping, r, lead_source) for r in rows]
    # -----------------------------------------------

    blank = [[""] * len(header) for _ in range(len(rows))]
    ws.insert_rows(blank, row=2)

    ws.update("A2", formatted, value_input_option="USER_ENTERED")

# ---------------- RUN ALL ----------------
def run_all():
    print("\n=== Combined Processing Started ===")

    exh_existing, spk_existing, vis_existing = load_existing_emails()

    event_exh = []
    event_spk = []

    meta_std = []
    meta_campaign_exh = []
    meta_campaign_spk = []
    meta_london_exh = []
    meta_london_spk = []
    meta_lead_based_exh = []
    meta_visitors = []


    # ---- Process Eventbrite ----
    for tab in SOURCE_TABS:
        process_eventbrite(tab, event_exh, event_spk, exh_existing, spk_existing)

    print("\nEventbrite Summary:")
    print(" Exhibitors:", len(event_exh))
    print(" Speakers:", len(event_spk))

    # ---- Process Meta Standard Leads ----
    process_meta_standard(meta_std, exh_existing)
    print("\nMeta Standard Summary:")
    print(" Exhibitors:", len(meta_std))

    # ---- Process Meta Campaign Leads ----
    process_meta_campaign(meta_campaign_exh, meta_campaign_spk, meta_visitors, exh_existing, spk_existing, vis_existing)

    print("\nMeta Campaign Summary:")
    print(" Exhibitors:", len(meta_campaign_exh))
    print(" Speakers:", len(meta_campaign_spk))
    print(" Visitors:", len(meta_visitors))

    # ---- Process Meta London Campaign Leads ----
    process_meta_london_campaign(
        meta_london_exh,
        meta_london_spk,
        meta_visitors,
        exh_existing,
        spk_existing,
        vis_existing
    )

    print("\nMeta London Campaign Summary:")
    print(" Exhibitors:", len(meta_london_exh))
    print(" Speakers:", len(meta_london_spk))
    print(" Visitors:", len(meta_visitors))

        # ---- Process Meta Lead Based Ads ----
    process_meta_lead_based_ads(
        meta_lead_based_exh, 
        exh_existing
    )

    print("\nMeta Lead-Based Ads Summary:")
    print(" Exhibitors:", len(meta_lead_based_exh))



    # ---- Insert all ----
    if event_exh:
        batch_insert_rows(EXHIBITOR_TAB, event_exh, "exhibitor", "Eventbrite")

    if event_spk:
        batch_insert_rows(SPEAKER_TAB, event_spk, "speaker", "Eventbrite")

    if meta_std:
        batch_insert_rows(EXHIBITOR_TAB, meta_std, "exhibitor", "Meta Lead")

    if meta_campaign_exh:
        batch_insert_rows(EXHIBITOR_TAB, meta_campaign_exh, "exhibitor", "Meta Lead – Campaign")

    if meta_campaign_spk:
        batch_insert_rows(SPEAKER_TAB, meta_campaign_spk, "speaker", "Meta Lead – Campaign")
    
    if meta_london_exh:
        batch_insert_rows(EXHIBITOR_TAB,meta_london_exh,"exhibitor","Meta Lead – London Campaign")

    if meta_london_spk:
        batch_insert_rows(SPEAKER_TAB,meta_london_spk,"speaker","Meta Lead – London Campaign")
    
    if meta_lead_based_exh:
        batch_insert_rows(EXHIBITOR_TAB,meta_lead_based_exh,"exhibitor","Meta Lead – Lead Based Ad")

    if meta_visitors:
        batch_insert_rows(VISITOR_TAB,meta_visitors,mode="visitor",lead_source="Meta Lead – Visitor")




    print("\n=== All Processing Completed ===\n")


if __name__ == "__main__":
    run_all()
