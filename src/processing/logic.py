import re
from datetime import datetime

# Constants ported from the original application
US_STATES = {
    'alabama': 'AL', 'alaska': 'AK', 'arizona': 'AZ', 'arkansas': 'AR', 'california': 'CA',
    'colorado': 'CO', 'connecticut': 'CT', 'delaware': 'DE', 'florida': 'FL', 'georgia': 'GA',
    'hawaii': 'HI', 'idaho': 'ID', 'illinois': 'IL', 'indiana': 'IN', 'iowa': 'IA',
    'kansas': 'KS', 'kentucky': 'KY', 'louisiana': 'LA', 'maine': 'ME', 'maryland': 'MD',
    'massachusetts': 'MA', 'michigan': 'MI', 'minnesota': 'MN', 'mississippi': 'MS', 'missouri': 'MO',
    'montana': 'MT', 'nebraska': 'NE', 'nevada': 'NV', 'new hampshire': 'NH', 'new jersey': 'NJ',
    'new mexico': 'NM', 'new york': 'NY', 'north carolina': 'NC', 'north dakota': 'ND', 'ohio': 'OH',
    'oklahoma': 'OK', 'oregon': 'OR', 'pennsylvania': 'PA', 'rhode island': 'RI', 'south carolina': 'SC',
    'south dakota': 'SD', 'tennessee': 'TN', 'texas': 'TX', 'utah': 'UT', 'vermont': 'VT',
    'virginia': 'VA', 'washington': 'WA', 'west virginia': 'WV', 'wisconsin': 'WI', 'wyoming': 'WY'
}
CAN_PROVINCES = {
    'alberta': 'AB', 'british columbia': 'BC', 'manitoba': 'MB', 'new brunswick': 'NB',
    'newfoundland': 'NL', 'nova scotia': 'NS', 'ontario': 'ON', 'prince edward island': 'PE',
    'quebec': 'QC', 'saskatchewan': 'SK'
}
CAN_POSTAL_PREFIX_TO_PROVINCE = {
    "A": "NL",
    "B": "NS",
    "C": "PE",
    "E": "NB",
    "G": "QC",
    "H": "QC",
    "J": "QC",
    "K": "ON",
    "L": "ON",
    "M": "ON",
    "N": "ON",
    "P": "ON",
    "R": "MB",
    "S": "SK",
    "T": "AB",
    "V": "BC",
    "Y": "YT",
}


def clean_text(value: str) -> str:
    """Collapses whitespace and normalizes empty-like values."""
    if value is None:
        return "Empty"
    cleaned = re.sub(r"\s+", " ", str(value)).strip(" ,")
    return cleaned if cleaned else "Empty"


def _infer_canadian_province(zip_code: str) -> str:
    """Infers province from postal-code prefix when it is uniquely mappable."""
    if not zip_code or zip_code == "Empty":
        return "Empty"
    compact = zip_code.replace(" ", "").upper()
    if not re.match(r"^[A-Z]\d[A-Z]\d[A-Z]\d$", compact):
        return "Empty"
    return CAN_POSTAL_PREFIX_TO_PROVINCE.get(compact[0], "Empty")


def normalize_address_fields(
    address: str,
    city_state: str,
    payable_to: str,
) -> tuple[str, str, str]:
    """Cleans address-related fields without relying on PDF-specific layout fixes."""
    clean_address = clean_text(address)
    clean_city_state = clean_text(city_state)
    clean_payable_to = clean_text(payable_to)

    if "Employees Health Trust" in clean_city_state:
        if clean_payable_to == "Empty":
            clean_payable_to = "Employees Health Trust"
        elif "Employees Health Trust" not in clean_payable_to:
            clean_payable_to = f"{clean_payable_to} Employees Health Trust".strip()
        clean_city_state = clean_text(
            clean_city_state.replace("Employees Health Trust", "")
        )

    return clean_address, clean_city_state, clean_payable_to

def parse_city_state_zip(text: str) -> tuple[str, str, str, str]:
    """Parses 'Mason OH 45040' or 'Whitby, ON L1R 2S7, Canada' into City, State, Zip, Country."""
    if not text or text == "Empty":
        return "Empty", "Empty", "Empty", "US"

    normalized_text = clean_text(text)
    normalized_text = normalized_text.replace("/", " ").replace(",", " ")
    normalized_text = re.sub(
        r'\b(Canada|USA|United States)\b',
        '',
        normalized_text,
        flags=re.IGNORECASE,
    ).strip()
    
    parts = normalized_text.split()
    
    zip_code, state, city, country = "Empty", "Empty", "Empty", "US"
    
    if not parts: return city, state, zip_code, country

    # 1. Detect ZIP
    if len(parts) >= 1 and re.match(r'^[A-Za-z]\d[A-Za-z]\d[A-Za-z]\d$', parts[-1]):
        zip_code = f"{parts[-1][:3]} {parts[-1][3:]}".upper()
        country = "CA"
        parts.pop()
    elif len(parts) >= 2 and re.match(r'^[A-Za-z]\d[A-Za-z]$', parts[-2]) and re.match(r'^\d[A-Za-z]\d$', parts[-1]):
        zip_code = f"{parts.pop(-2)} {parts.pop(-1)}".upper()
        country = "CA"
    elif len(parts) >= 1 and re.match(r'^\d{5}(-\d{4})?$', parts[-1]):
        zip_code = parts.pop()
    
    if not parts: return city, state, zip_code, country

    # 2. Detect State/Prov
    potential_state = parts[-1].lower()
    
    if len(parts[-1]) == 2 and parts[-1].isalpha():
        state_code = parts.pop().upper()
        state = state_code
        if state_code in US_STATES.values():
            country = "US"
        elif state_code in CAN_PROVINCES.values():
            state = state_code
            country = "CA"
    elif potential_state in US_STATES:
        state = US_STATES[potential_state]
        country = "US"
        parts.pop()
    elif potential_state in CAN_PROVINCES:
        state = CAN_PROVINCES[potential_state]
        country = "CA"
        parts.pop()

    if country == "CA" and state == "Empty":
        state = _infer_canadian_province(zip_code)

    city = " ".join(parts).title() if parts else "Empty"
    return city, state, zip_code, country


def normalize_date(value: str) -> str:
    """Normalizes a date string to MM/DD/YYYY when possible."""
    if not value or value == "Empty":
        return "Empty"

    raw_value = value.split(" ")[0].strip()
    formats = [
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%m-%d-%Y",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(raw_value, fmt).strftime("%m/%d/%Y")
        except ValueError:
            continue
    return raw_value


def normalize_numeric_code(value: str, width: int = 10) -> str:
    """Removes separators and pads numeric codes to a fixed width."""
    if not value or value == "Empty":
        return "Empty"

    raw_value = str(value).strip()
    cleaned_value = re.sub(r"[-\s]+", "", raw_value)
    if cleaned_value.isdigit():
        return cleaned_value.zfill(width)
    return cleaned_value


def classify_mail_group(company_code: str) -> str:
    normalized = str(company_code or "").strip().upper()
    if normalized in {"1000", "2000", "E100"}:
        return "FMS"
    if normalized in {"0032", "0016", "0133", "0060", "1010", "5500", "0224"}:
        return "AFS"
    if normalized.startswith("E"):
        return "AFS"
    return "AFS"

def apply_business_rules(raw_data: dict) -> dict:
    """Applies business logic to transform raw extracted data into a structured format for the CSV."""
    
    # Rule: Determine CompanyCode
    comp_text = raw_data.get("Company", "")
    comp_code_match = re.match(r'^([A-Za-z0-9]+)', comp_text)
    comp_code = comp_code_match.group(1) if comp_code_match else "Empty"
    
    # Rule: Determine VendorNum based on CompanyCode
    vendor_num = "8000001"
    if "E100" in comp_code.upper() or "E1OO" in comp_code.upper():
        vendor_num = "900000"
    elif comp_code in ["1000", "2000"]:
        vendor_num = "900010"
        
    # Rule: Fallback for InvoiceNum
    inv_num = raw_data.get("Invoice Number")
    if not inv_num or inv_num == "Empty":
        inv_num = raw_data.get("Id", "Empty") # 'Id' is a placeholder, might not exist

    # Rule: Fallback and formatting for InvoiceDate
    inv_date = raw_data.get("Invoice Date")
    if not inv_date or inv_date == "Empty":
        inv_date = raw_data.get("Created", "Empty") # 'Created' is a placeholder
    if inv_date and inv_date != "Empty":
        inv_date = normalize_date(inv_date)

    # Rule: Clean and convert Amount
    amt_str = str(raw_data.get("Amount", "0")).replace(",", "").replace("$", "").strip()
    try:
        amount_val = float(amt_str)
    except (ValueError, TypeError):
        amount_val = 0.0

    # Rule: Clean and parse address components
    address, city_state, payable_to = normalize_address_fields(
        raw_data.get("Address", "Empty"),
        raw_data.get("City/State", "Empty"),
        raw_data.get("Payable to", "Empty"),
    )
    city, state, zip_c, country_from_address = parse_city_state_zip(city_state)
    
    # Rule: Determine country based on currency if not clear from address
    curr = raw_data.get("Currency", "USD").upper()
    country = country_from_address if country_from_address != "US" else ("CA" if "CAD" in curr else "US")

    processed_row = {
        "CompanyCode": comp_code,
        "VendorNum": vendor_num,
        "InvoiceNum": inv_num,
        "InvoiceDate": inv_date,
        "Amount": amount_val,
        "Currency": curr,
        "CostCenter": raw_data.get("Cost/Profit center", "Empty"),
        "GLAccount": raw_data.get("GL Account", "Empty"),
        "WBS": raw_data.get("WBS Element", "Empty"),
        "PayableTo": payable_to,
        "Address": address,
        "City": city,
        "State": state,
        "Zip": zip_c,
        "Country": country,
        "BrandCode": raw_data.get("Brand code", "Empty"),
        # Keep raw GL for validation purposes
        "_raw_gl": raw_data.get("GL Account", "Empty"),
    }
    return processed_row

def validate_ticket_data(processed_data: dict) -> list[str]:
    """Validates the processed ticket data based on the original app's highlighting logic."""
    errors = []

    # Rule: Amount should not be zero
    if processed_data.get("Amount", 0.0) == 0.0:
        errors.append("Amount is zero.")

    # Rule: Zip code should not be empty
    if not processed_data.get("Zip") or processed_data.get("Zip") == "Empty":
        errors.append("Zip code is missing.")

    # Rule: Cost Center cleaning and validation
    raw_cc = str(processed_data.get("CostCenter", "")).strip()
    if not raw_cc or raw_cc.lower() in ["attached", "empty"]:
        errors.append("Cost/Profit Center is missing.")

    # Rule: GL Account cleaning and validation
    raw_gl = str(processed_data.get("_raw_gl", "")).replace(" ", "").strip()
    if not raw_gl or raw_gl.lower() in ["attached", "empty"]:
        errors.append("GL Account is missing.")
    elif raw_gl.isdigit() and len(raw_gl) != 10:
        errors.append(f"GL Account '{raw_gl}' is not 10 digits.")

    # Rule: Cross-reference GL Account and VendorNum
    if raw_gl.startswith("P") and str(processed_data.get("VendorNum")) != "8000001":
        errors.append(f"GL Account starts with 'P' but VendorNum is not '8000001'.")

    return errors
