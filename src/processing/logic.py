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

FMS_ALLOWED_COMBINATIONS = {
    ("1000", "900010", "USD"),
    ("2000", "900010", "CAD"),
    ("E100", "900000", "USD"),
}

AFS_EXPLICIT_ALLOWED_COMBINATIONS = {
    ("0032", "8000001", "USD"),
    ("0016", "8000001", "CAD"),
    ("1010", "8000001", "USD"),
    ("0060", "8000001", "CAD"),
    ("0060", "8000001", "USD"),
    ("0182", "8000001", "USD"),
    ("0124", "8000001", "USD"),
    ("0133", "8000001", "USD"),
    ("0224", "8000001", "USD"),
}

AFS_NOT_APPLICABLE_COMPANY_CODES = {"5500", "5700"}


def clean_text(value: str) -> str:
    """Collapses whitespace and normalizes empty-like values."""
    if value is None:
        return "Empty"
    cleaned = re.sub(r"\s+", " ", str(value)).strip(" ,")
    return cleaned if cleaned else "Empty"


def parse_amount(value: str) -> tuple[float, list[str]]:
    """Tries to parse amount safely. Returns (value, errors)."""
    raw_val = clean_text(value).replace("$", "").replace(" ", "")
    if not raw_val or raw_val == "Empty":
        return 0.0, ["Amount is missing."]
    
    # Check for multiple decimal-like separators or strange chars
    if raw_val.count(".") > 1 and raw_val.count(",") > 1:
            return 0.0, [f"Ambiguous amount format (multiple separators): '{raw_val}'"]

    normalized = raw_val
    # Case: 1,000.00 (Standard US)
    if "," in normalized and "." in normalized:
        if normalized.rfind(".") > normalized.rfind(","):
            normalized = normalized.replace(",", "")
        else:
            # Case: 1.000,00 (European/Latam)
            # In USD/CAD context, this is risky.
            normalized = normalized.replace(".", "").replace(",", ".")
    
    # Case: 4,000 or 4,50
    elif "," in normalized and normalized.count(",") == 1:
        parts = normalized.split(",")
        if len(parts[-1]) == 3:
            # Likely thousands
            normalized = normalized.replace(",", "")
        elif len(parts[-1]) in [1, 2]:
            # Likely decimal
            normalized = normalized.replace(",", ".")
    
    # Case: 4.000 (Could be 4 or 4000)
    elif "." in normalized and normalized.count(".") == 1:
        parts = normalized.split(".")
        if len(parts[-1]) == 3:
            # If it's something like 4.000, it's ambiguous. Is it 4 or 4000?
            # Standard USD/CAD uses . as decimal. So 4.000 is 4.
            # But users might use . as thousands.
            # Let's be conservative. If it's exactly 3 digits after the dot, 
            # and no other separators, we mark as ambiguous if we want to be 100% sure.
            # However, standard float(4.000) is 4.0.
            pass

    try:
        val = float(normalized)
        return val, []
    except ValueError:
        return 0.0, [f"Invalid amount format: '{raw_val}'"]


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


def _is_parseable_date(value: str) -> bool:
    if not value or value == "Empty":
        return False

    raw_value = str(value).split(" ")[0].strip()
    formats = [
        "%m/%d/%Y",
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%m-%d-%Y",
    ]
    for fmt in formats:
        try:
            datetime.strptime(raw_value, fmt)
            return True
        except ValueError:
            continue
    return False


def _is_future_normalized_date(value: str) -> bool:
    if not value or value == "Empty":
        return False
    try:
        parsed = datetime.strptime(value, "%m/%d/%Y").date()
    except ValueError:
        return False
    return parsed > datetime.now().date()


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


def is_allowed_one_time_combination(
    company_code: str,
    vendor_num: str,
    currency: str,
) -> tuple[bool, str]:
    normalized_company = str(company_code or "").strip().upper()
    normalized_vendor = str(vendor_num or "").strip().upper()
    normalized_currency = str(currency or "").strip().upper()
    mail_group = classify_mail_group(normalized_company)

    if not normalized_company:
        return False, "Company code is missing."
    if not normalized_vendor:
        return False, "Vendor number is missing."
    if not normalized_currency:
        return False, "Currency is missing."

    if mail_group == "FMS":
        if (normalized_company, normalized_vendor, normalized_currency) in FMS_ALLOWED_COMBINATIONS:
            return True, ""
        return (
            False,
            f"Combination CompanyCode/Vendor/Currency is not allowed for FMS: "
            f"{normalized_company}/{normalized_vendor}/{normalized_currency}.",
        )

    if normalized_company in AFS_NOT_APPLICABLE_COMPANY_CODES:
        return (
            False,
            f"Company code {normalized_company} is marked as not applicable for OneTime Check.",
        )

    if (normalized_company, normalized_vendor, normalized_currency) in AFS_EXPLICIT_ALLOWED_COMBINATIONS:
        return True, ""

    if (
        normalized_company.startswith("E")
        and normalized_company != "E100"
        and normalized_vendor == "8000001"
        and normalized_currency == "USD"
    ):
        return True, ""

    return (
        False,
        f"Combination CompanyCode/Vendor/Currency is not allowed for AFS: "
        f"{normalized_company}/{normalized_vendor}/{normalized_currency}.",
    )

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
    amount_val, amount_errors = parse_amount(raw_data.get("Amount", "0"))


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
        "_amount_errors": amount_errors,
    }
    return processed_row

def validate_ticket_data(processed_data: dict) -> list[str]:
    """Validates the processed ticket data based on the original app's highlighting logic."""
    errors = []

    # Rule: Amount validation
    amount_errors = processed_data.get("_amount_errors", [])
    if amount_errors:
        errors.extend(amount_errors)
    elif processed_data.get("Amount", 0.0) == 0.0:
        errors.append("Amount is zero.")

    combo_ok, combo_error = is_allowed_one_time_combination(
        processed_data.get("CompanyCode", ""),
        processed_data.get("VendorNum", ""),
        processed_data.get("Currency", ""),
    )
    if not combo_ok and combo_error:
        errors.append(combo_error)

    # Rule: Zip code should not be empty
    if not processed_data.get("Zip") or processed_data.get("Zip") == "Empty":
        errors.append("Zip code is missing.")

    invoice_date = processed_data.get("InvoiceDate", "Empty")
    if not invoice_date or invoice_date == "Empty":
        errors.append("Invoice Date is missing.")
    elif not _is_parseable_date(invoice_date):
        errors.append(f"Invoice Date '{invoice_date}' is invalid.")
    elif _is_future_normalized_date(invoice_date):
        errors.append(f"Invoice Date '{invoice_date}' cannot be in the future.")

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
