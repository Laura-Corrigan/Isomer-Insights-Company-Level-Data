"""
Combined Old Template → BBB Managed Funds New Equity Template
==============================================================
Processes multiple old-format fund templates, cleans, validates,
and consolidates into a single BBB output workbook.
"""

import pandas as pd
import numpy as np
import os
import re
import json
import tempfile
import shutil
from datetime import datetime, date, timedelta
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import urllib.request

# Temp directory for file conversions (works on any OS)
TEMP_DIR = tempfile.mkdtemp(prefix="bbb_")

# Find LibreOffice binary (different path on Mac vs Linux)
LIBREOFFICE = shutil.which("libreoffice") or shutil.which("soffice")
if not LIBREOFFICE:
    mac_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if os.path.exists(mac_path):
        LIBREOFFICE = mac_path
import base64

# ============================================================================
# CONFIG
# ============================================================================

NEW_COLS = [
    "Fund Reference", "Period End Date", "Investment Reference", "Investment Type",
    "Enterprise Name", "Company Number", "Investment Industry", "Investment Location",
    "Percentage of Equity", "Original Investment Date", "Investment Country",
    "Loan / Investment Denomination Currency", "Exit Type",
    "Gross Investment Amount", "Unrealised Cost", "Unrealised Value",
    "Realised Cost", "Realised Value",
    "BBB Gross Investment Amount", "BBB Unrealised Cost", "BBB Unrealised Value",
    "BBB Realised Cost", "BBB Realised Value",
    "Investment Valuation Basis", "BBB Participation %",
    "Annual Turnover", "Number of Employees", "Employees YE", "Turnover YE",
    "Investment Stage", "Latest Investment Round", "Cumulative Roundsize to Date",
    "Enterprise Value", "Other Income",
    "Total Number of Founders/Executives", "Number of Female Founders/Executives",
    "Number of Ethnic Minority Founders/Executives", "Founder-led Business",
    "Enterprise Scope 1 Green House Gas (GHG) emissions",
    "Enterprise Scope 2 Green House Gas (GHG) emissions",
    "Enterprise Scope 3 Green House Gas (GHG) emissions",
    "Source of reported emissions data",
    "Green House Gas (GHG) Protocol confirmation",
]

VALID_INV_TYPES = ["Equity", "Quasi-Equity", "Convertible Loan"]
VALID_STAGES = [
    "Angel", "Pre Seed", "Seed", "Seed Extension",
    "Series A", "Series B", "Series C", "Series D", "Series E and beyond",
    "Bridge Financing", "Expansion/Growth Capital", "Secondary Buyout",
    "MBI", "MBO", "IPO"
]
UK_PC_RE = re.compile(r'^[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}$')
CC2 = {"FR":"FRA","US":"USA","GB":"GBR","DE":"DEU","NL":"NLD","SE":"SWE",
       "ES":"ESP","IT":"ITA","IE":"IRL","DK":"DNK","FI":"FIN","NO":"NOR",
       "PT":"PRT","AT":"AUT","BE":"BEL","CH":"CHE","PL":"POL","CZ":"CZE",
       "RO":"ROU","EE":"EST","LT":"LTU","LV":"LVA","HU":"HUN","LU":"LUX",
       "SW":"SWE","GE":"DEU","BU":"BGR","UK":"GBR"}

# Full country names → ISO3
COUNTRY_NAMES = {
    "UNITED STATES": "USA", "UNITED STATES OF AMERICA": "USA", "US": "USA", "USA": "USA",
    "UNITED KINGDOM": "GBR", "UK": "GBR", "ENGLAND": "GBR", "SCOTLAND": "GBR",
    "WALES": "GBR", "NORTHERN IRELAND": "GBR", "GREAT BRITAIN": "GBR", "GBR": "GBR",
    "GERMANY": "DEU", "DEUTSCHLAND": "DEU", "DEU": "DEU",
    "FRANCE": "FRA", "FRA": "FRA",
    "SWEDEN": "SWE", "SWE": "SWE",
    "DENMARK": "DNK", "DNK": "DNK",
    "NETHERLANDS": "NLD", "NLD": "NLD", "HOLLAND": "NLD",
    "SPAIN": "ESP", "ESP": "ESP",
    "ITALY": "ITA", "ITA": "ITA",
    "IRELAND": "IRL", "IRL": "IRL",
    "FINLAND": "FIN", "FIN": "FIN",
    "NORWAY": "NOR", "NOR": "NOR",
    "PORTUGAL": "PRT", "PRT": "PRT",
    "AUSTRIA": "AUT", "AUT": "AUT",
    "BELGIUM": "BEL", "BEL": "BEL",
    "SWITZERLAND": "CHE", "CHE": "CHE",
    "POLAND": "POL", "POL": "POL",
    "CZECH REPUBLIC": "CZE", "CZECHIA": "CZE", "CZE": "CZE",
    "ROMANIA": "ROU", "ROU": "ROU",
    "ESTONIA": "EST", "EST": "EST",
    "LITHUANIA": "LTU", "LTU": "LTU",
    "LATVIA": "LVA", "LVA": "LVA",
    "HUNGARY": "HUN", "HUN": "HUN",
    "LUXEMBOURG": "LUX", "LUX": "LUX",
    "ISRAEL": "ISR", "ISR": "ISR",
    "CANADA": "CAN", "CAN": "CAN",
    "AUSTRALIA": "AUS", "AUS": "AUS",
    "JAPAN": "JPN", "JPN": "JPN",
    "SINGAPORE": "SGP", "SGP": "SGP",
    "INDIA": "IND", "IND": "IND",
    "BRAZIL": "BRA", "BRA": "BRA",
    "SLOVENIA": "SVN", "SVN": "SVN", "SI": "SVN",
    "SLOVAKIA": "SVK", "SVK": "SVK", "SK": "SVK",
    # Non-standard 2-letter codes sometimes used by managers
    "SW": "SWE", "GE": "DEU", "BU": "BGR", "SP": "ESP",
    "NE": "NLD", "PO": "PRT", "AU": "AUT", "SZ": "CHE",
    "BULGARIA": "BGR", "BGR": "BGR", "BG": "BGR",
}

# ============================================================================
# COMPANIES HOUSE API
# ============================================================================

CH_API_KEY = os.environ.get("COMPANIES_HOUSE_API_KEY", "69de4594-0765-4658-a1b4-b4d69eebed2e")
CH_API_BASE = "https://api.company-information.service.gov.uk"
CH_CACHE = {}

def companies_house_lookup(company_number: str) -> dict | None:
    if not company_number:
        return None
    if company_number in CH_CACHE:
        return CH_CACHE[company_number]
    if not CH_API_KEY:
        print(f"    ⚠ No CH API key — cannot look up {company_number}")
        return None
    url = f"{CH_API_BASE}/company/{company_number}"
    try:
        credentials = base64.b64encode(f"{CH_API_KEY}:".encode()).decode()
        req = urllib.request.Request(url, headers={
            "Authorization": f"Basic {credentials}", "Accept": "application/json"})
        resp = urllib.request.urlopen(req, timeout=10)
        data = json.loads(resp.read().decode())
        result = {"name": data.get("company_name", ""), "sic": None, "postcode": None,
                  "company_number": company_number}
        sic_codes = data.get("sic_codes", [])
        if sic_codes:
            result["sic"] = sic_codes[0]
        office = data.get("registered_office_address", {})
        result["postcode"] = office.get("postal_code", "")
        CH_CACHE[company_number] = result
        return result
    except Exception as e:
        print(f"    ✗ CH API error for {company_number}: {e}")
        CH_CACHE[company_number] = None
        return None


def companies_house_search(company_name: str) -> dict | None:
    """Search Companies House by name. Returns best match with company_number, sic, postcode."""
    if not company_name or not CH_API_KEY:
        return None
    cache_key = f"SEARCH:{company_name.upper().strip()}"
    if cache_key in CH_CACHE:
        return CH_CACHE[cache_key]
    url = f"{CH_API_BASE}/search/companies?q={urllib.request.quote(company_name)}&items_per_page=5"
    try:
        credentials = base64.b64encode(f"{CH_API_KEY}:".encode()).decode()
        req = urllib.request.Request(url, headers={
            "Authorization": f"Basic {credentials}", "Accept": "application/json"})
        resp = urllib.request.urlopen(req, timeout=10)
        data = json.loads(resp.read().decode())
        items = data.get("items", [])
        if not items:
            CH_CACHE[cache_key] = None
            return None
        # Try to find an active match
        best = None
        for item in items:
            status = item.get("company_status", "").lower()
            if status == "active":
                best = item
                break
        if not best:
            best = items[0]  # fallback to first result
        co_num = best.get("company_number", "")
        # Now do a full lookup to get SIC + postcode
        result = companies_house_lookup(co_num)
        if result:
            result["ch_matched_name"] = best.get("title", "")
            CH_CACHE[cache_key] = result
            return result
        CH_CACHE[cache_key] = None
        return None
    except Exception as e:
        print(f"    ✗ CH search error for '{company_name}': {e}")
        CH_CACHE[cache_key] = None
        return None

# ============================================================================
# HELPER: EXCEL SERIAL DATE → PYTHON DATE
# ============================================================================

def excel_serial_to_date(val):
    """Convert Excel serial number, string date, or Timestamp to datetime."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    if isinstance(val, (datetime, date)):
        return val
    if isinstance(val, pd.Timestamp):
        return val.to_pydatetime()
    if isinstance(val, (int, float)):
        try:
            return datetime(1899, 12, 30) + timedelta(days=int(val))
        except:
            return None
    if isinstance(val, str):
        for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"]:
            try:
                return datetime.strptime(val.strip(), fmt)
            except ValueError:
                continue
    return None

# ============================================================================
# COLUMN MAPS FOR DIFFERENT OLD TEMPLATE VARIANTS
# ============================================================================

# Template variant 1: "Portfolio Company Data" sheet (e.g., Old_Template_Example_1)
COL_MAP_V1 = {
    "Fund Reference": "Fund Reference",
    "Investment Reference": "Investment Reference",
    "Period End Date": "Period End Date",
    "Company Name": "Enterprise Name",
    "All Companies, at Time of Investment, Updated as Needed": "Enterprise Name",  # LibreOffice variant
    "Investment Type": "Investment Type",
    "Original Investment Date": "Original Investment Date",
    "Investment Country": "Investment Country",
    "Company Number": "Company Number",
    "Investment Industry": "Investment Industry",
    "Investment Location": "Investment Location",
    "Percentage  Ownership": "Percentage of Equity",
    "Investment Denomination Currency": "Loan / Investment Denomination Currency",
    "Investment Stage ": "Investment Stage",
    "Investment Stage": "Investment Stage",
    "Latest Investment Round": "Latest Investment Round",
    "Gross Investment Amount From Fund (Cumulative)": "Gross Investment Amount",
    "Company\u2019s Cumulative Investment To Date": "Cumulative Roundsize to Date",
    "Company's Cumulative Investment To Date": "Cumulative Roundsize to Date",  # ASCII variant
    "Unrealised Cost": "Unrealised Cost",
    "Unrealised Value": "Unrealised Value",
    "Realised Cost": "Realised Cost",
    "Realised Value": "Realised Value",
    "Other Income": "Other Income",
    "Annual Turnover": "Annual Turnover",
    "Employees YE": "Number of Employees",
    "Employees ": "Employees YE",
    "Turnover YE (with currency)": "Turnover YE",
    "Exit Type": "Exit Type",
    "Total Number of Directors amd Executives": "Total Number of Founders/Executives",
    "Total Number of Directors": "Total Number of Founders/Executives",
    "Number of Female Directors amd Executives": "Number of Female Founders/Executives",
    "Number of Female Directors": "Number of Female Founders/Executives",
    "Number of Ethnic Minority Directors and Executives": "Number of Ethnic Minority Founders/Executives",
    "Number of Ethnic Minority Directors": "Number of Ethnic Minority Founders/Executives",
}

# Template variant 2: "BBI_Equity Template" sheet (e.g., Old_Template_Example_2)
COL_MAP_V2 = {
    "Fund Reference": "Fund Reference",
    "Investment Reference": "Investment Reference",
    "Period End Date": "Period End Date",
    "Enterprise Name": "Enterprise Name",
    "Investment Type": "Investment Type",
    "Original Investment Date": "Original Investment Date",
    "Investment Country": "Investment Country",
    "Company Number": "Company Number",
    "Investment Industry": "Investment Industry",
    "Investment Location": "Investment Location",
    "Percentage of Equity": "Percentage of Equity",
    "Loan / Investment Denomination Currency": "Loan / Investment Denomination Currency",
    "Investment Stage": "Investment Stage",
    "Latest Investment Round": "Latest Investment Round",
    "Gross Investment Amount": "Gross Investment Amount",
    "Cumulative Roundsize to Date": "Cumulative Roundsize to Date",
    "Unrealised Cost": "Unrealised Cost",
    "Unrealised Value": "Unrealised Value",
    "Realised Cost": "Realised Cost",
    "Realised Value": "Realised Value",
    "Other Income": "Other Income",
    "Annual Turnover": "Annual Turnover",
    "Employees QE": "Number of Employees",
    "Employees YE": "Employees YE",
    "Turnover YE": "Turnover YE",
    "Exit Type": "Exit Type",
    "Total Number of Directors": "Total Number of Founders/Executives",
    "Number of Female Directors": "Number of Female Founders/Executives",
    "Number of Ethnic Minority Directors": "Number of Ethnic Minority Founders/Executives",
}

# ============================================================================
# PROCESS A SINGLE FUND
# ============================================================================

def process_fund(filepath, fund_label, fund_denomination_ccy=None):
    """Read an old template, detect variant, map & return cleaned DataFrame + issues."""

    print(f"\n--- Processing: {fund_label} ---")

    # Handle .xlsb / .xls conversion
    if filepath.endswith('.xlsb') or filepath.endswith('.xls'):
        if not LIBREOFFICE:
            print(f"  ✗ Cannot convert {os.path.basename(filepath)} — LibreOffice not installed")
            print(f"    Please convert to .xlsx manually or install LibreOffice")
            return pd.DataFrame(columns=NEW_COLS), [], [], None
        import subprocess
        subprocess.run([LIBREOFFICE, "--headless", "--convert-to", "xlsx",
                        filepath, "--outdir", TEMP_DIR], capture_output=True)
        base = os.path.splitext(os.path.basename(filepath))[0]
        filepath = f"{TEMP_DIR}/{base}.xlsx"

    # Try to open; if corrupt XML, re-save via LibreOffice
    try:
        wb_check = load_workbook(filepath, data_only=True)
        sheets = wb_check.sheetnames
        wb_check.close()
    except (ValueError, Exception) as e:
        if ("stylesheet" in str(e).lower() or "invalid" in str(e).lower()) and LIBREOFFICE:
            print(f"  ⚠ Corrupt XML detected, re-saving via LibreOffice...")
            import subprocess
            subprocess.run([LIBREOFFICE, "--headless", "--convert-to", "xlsx",
                            filepath, "--outdir", TEMP_DIR], capture_output=True)
            base = os.path.splitext(os.path.basename(filepath))[0]
            filepath = f"{TEMP_DIR}/{base}.xlsx"
            wb_check = load_workbook(filepath, data_only=True)
            sheets = wb_check.sheetnames
            wb_check.close()
        else:
            raise

    # Detect template variant
    # V4: Azava format — detect by checking first sheet for Azava-specific columns
    azava_sheet = None
    for sn in sheets:
        wb_tmp = load_workbook(filepath, data_only=True)
        ws_tmp = wb_tmp[sn]
        hdrs = [str(ws_tmp.cell(1, c).value or "").strip() for c in range(1, min(ws_tmp.max_column+1, 30))]
        wb_tmp.close()
        if "Total Invested" in hdrs and "Unrealized Value" in hdrs:
            azava_sheet = sn
            break

    if azava_sheet:
        df_old = pd.read_excel(filepath, sheet_name=azava_sheet, header=0)
        name_col = 'Name'
        variant = "V4"
        variant_sheet = azava_sheet
    elif 'Portfolio Company Data' in sheets:
        df_old = pd.read_excel(filepath, sheet_name='Portfolio Company Data', header=1)
        col_map = COL_MAP_V1
        # Name column varies: "Company Name" or LibreOffice variant
        if 'Company Name' in df_old.columns:
            name_col = 'Company Name'
        else:
            # Find the column that maps to Enterprise Name
            name_col = None
            for c in df_old.columns:
                if c in col_map and col_map[c] == "Enterprise Name":
                    name_col = c
                    break
            if not name_col:
                # Last resort: find first column with company-like data
                name_col = df_old.columns[3]  # typically 4th column
        variant = "V1"
        variant_sheet = 'Portfolio Company Data'
    elif 'BBI_Equity Template' in sheets:
        df_old = pd.read_excel(filepath, sheet_name='BBI_Equity Template', header=2)
        col_map = COL_MAP_V2
        name_col = 'Enterprise Name'
        variant = "V2"
        variant_sheet = 'BBI_Equity Template'
    elif 'Isomer Template' in sheets:
        # V5: Isomer in-house template — same columns as V2 but different sheet name
        df_old = pd.read_excel(filepath, sheet_name='Isomer Template', header=0)
        col_map = COL_MAP_V2
        name_col = 'Enterprise Name'
        variant = "V5"
        variant_sheet = 'Isomer Template'
    elif 'Managed Funds Template' in sheets:
        # V3: Already in new BBB format — columns match output directly
        df_old = pd.read_excel(filepath, sheet_name='Managed Funds Template', header=0)
        # 1:1 mapping — column names match NEW_COLS
        col_map = {c: c for c in NEW_COLS}
        name_col = 'Enterprise Name'
        variant = "V3"
        variant_sheet = 'Managed Funds Template'
    else:
        print(f"  ✗ Unknown template format. Sheets: {sheets}")
        return pd.DataFrame(columns=NEW_COLS), [], [], None

    df_old = df_old.dropna(how='all')
    df_old = df_old[df_old[name_col].notna()].reset_index(drop=True)
    print(f"  Variant: {variant} | Rows: {len(df_old)}")

    # =========================================================================
    # V4 AZAVA: Custom mapping (completely different column structure)
    # =========================================================================
    if variant == "V4":
        new_df = pd.DataFrame(index=range(len(df_old)), columns=NEW_COLS)

        # Core mappings
        new_df["Enterprise Name"] = df_old["Name"].values
        new_df["Investment Reference"] = df_old["Name"].values
        new_df["Investment Country"] = df_old.get("Country", pd.Series()).values if "Country" in df_old.columns else None
        new_df["Original Investment Date"] = df_old.get("Investment Date", pd.Series()).values if "Investment Date" in df_old.columns else None
        new_df["Loan / Investment Denomination Currency"] = df_old.get("Reporting Currency", pd.Series()).values if "Reporting Currency" in df_old.columns else None
        new_df["Fund Reference"] = df_old.get("investors", pd.Series()).values if "investors" in df_old.columns else None

        # Financial mappings
        # Total Invested = Gross Investment Amount = Unrealised Cost (cost basis)
        new_df["Gross Investment Amount"] = df_old.get("Total Invested", pd.Series()).values if "Total Invested" in df_old.columns else None
        new_df["Unrealised Cost"] = df_old.get("Total Invested", pd.Series()).values if "Total Invested" in df_old.columns else None
        new_df["Unrealised Value"] = df_old.get("Unrealized Value", pd.Series()).values if "Unrealized Value" in df_old.columns else None
        new_df["Realised Value"] = df_old.get("Realized Value", pd.Series()).values if "Realized Value" in df_old.columns else None

        # Investment Type: Azava is always equity
        new_df["Investment Type"] = "Equity"

        # Period End Date: use the max investment date as reporting period
        max_date = df_old["Investment Date"].max() if "Investment Date" in df_old.columns else None
        for i in range(len(new_df)):
            new_df.at[i, "Period End Date"] = max_date

        # Save Legal Name for CH search (not a BBB column, stored separately)
        azava_legal_names = df_old.get("Legal Name", pd.Series(dtype=str)).values.tolist()

        col_map = {}  # prevent generic mapping from running
    else:
        # Map columns (V1, V2, V3)
        new_df = pd.DataFrame(index=range(len(df_old)), columns=NEW_COLS)
        for old_col, new_col in col_map.items():
            if old_col in df_old.columns:
                new_df[new_col] = df_old[old_col].values
        azava_legal_names = [None] * len(new_df)

    for c in ["Company Number", "Investment Industry", "Investment Location",
              "Investment Country", "Latest Investment Round", "Investment Stage",
              "Investment Reference", "Fund Reference",
              "Period End Date", "Original Investment Date"]:
        new_df[c] = new_df[c].astype(object)

    # Forward-fill columns that are typically only on the first row
    for fill_col in ["Fund Reference", "Period End Date"]:
        if fill_col in new_df.columns:
            first_val = None
            for i in range(len(new_df)):
                v = new_df.at[i, fill_col]
                if pd.notna(v) and str(v).strip() not in ("", "nan", "None"):
                    first_val = v
                    break
            if first_val is not None:
                for i in range(len(new_df)):
                    v = new_df.at[i, fill_col]
                    if pd.isna(v) or str(v).strip() in ("", "nan", "None"):
                        new_df.at[i, fill_col] = first_val

    # Clean & validate
    issues = []
    def add(sev, row, col, msg):
        issues.append({"sev": sev, "row": row, "col": col, "msg": msg})

    # =========================================================================
    # PARSE CURRENCY SYMBOLS FROM FINANCIAL VALUES
    # =========================================================================
    # Real files often have values like "€1,234.56", "£-", "$1,000.00", "1,945,302.90 kr"
    # We need to: (a) extract the currency, (b) store plain numbers, (c) blank zeros

    FINANCIAL_COLS = [
        "Gross Investment Amount", "Unrealised Cost", "Unrealised Value",
        "Realised Cost", "Realised Value", "Other Income", "Interest Received",
        "Cumulative Roundsize to Date", "Private Sector Leverage",
    ]

    CCY_SYMBOLS = {"£": "GBP", "€": "EUR", "$": "USD", "kr": "SEK"}

    def parse_currency_value(val):
        """Parse a currency-formatted value. Returns (currency_code, numeric_value) or (None, val)."""
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return None, None
        if isinstance(val, (int, float)):
            return None, float(val)
        s = str(val).strip()
        if s in ("", "nan", "None"):
            return None, None

        # Detect currency symbol
        detected_ccy = None
        for sym, code in CCY_SYMBOLS.items():
            if sym in s:
                detected_ccy = code
                s = s.replace(sym, "")
                break

        # Clean remaining string: remove commas, spaces, handle "-" as zero
        s = s.strip().replace(",", "").replace(" ", "")
        if s in ("-", "", "−"):
            return detected_ccy, 0.0
        try:
            return detected_ccy, float(s)
        except (ValueError, TypeError):
            return detected_ccy, val  # return original if can't parse

    # Pass 1: Parse all financial values — extract currencies and clean numbers
    row_currencies = []  # track detected currency per row from N-R values
    for i in range(len(new_df)):
        row_ccys = set()
        for col in FINANCIAL_COLS:
            if col not in new_df.columns:
                continue
            val = new_df.at[i, col]
            ccy, num = parse_currency_value(val)
            if ccy:
                row_ccys.add(ccy)
            # Store clean numeric value (or None for zero)
            if isinstance(num, (int, float)):
                if num == 0 or abs(num) < 0.001:
                    new_df.at[i, col] = None  # zero → blank
                else:
                    new_df.at[i, col] = num
            elif num is None:
                new_df.at[i, col] = None
        row_currencies.append(row_ccys)

    # Detect overall currency from N-R values
    all_detected = set()
    for rc in row_currencies:
        all_detected.update(rc)

    if len(all_detected) > 1:
        add("ERROR", 0, "Currency (N-R)",
            f"'{fund_label}': Financial columns contain MIXED CURRENCIES: "
            f"{', '.join(sorted(all_detected))} — go back to manager, "
            f"all values must be in the fund denomination currency")
    elif len(all_detected) == 1:
        detected_fund_ccy = list(all_detected)[0]
        # Use this to inform column L if not already set
        if not fund_denomination_ccy:
            fund_denomination_ccy = detected_fund_ccy

    for i in range(len(new_df)):
        name = str(new_df.at[i, "Enterprise Name"] or "Unknown")

        # Investment Reference: default to Enterprise Name if blank
        ir = new_df.at[i, "Investment Reference"]
        if pd.isna(ir) or str(ir).strip() in ("", "nan"):
            new_df.at[i, "Investment Reference"] = name

        # Fund Reference
        fr = new_df.at[i, "Fund Reference"]
        if pd.isna(fr) or str(fr).strip() in ("", "nan"):
            add("ERROR", i, "Fund Reference", f"'{name}': Fund Reference blank")

        # Parse dates (handle Excel serial numbers)
        for dc in ["Period End Date", "Original Investment Date"]:
            new_df.at[i, dc] = excel_serial_to_date(new_df.at[i, dc])

        # Country codes — normalise to ISO3
        country = str(new_df.at[i, "Investment Country"] or "").strip().upper()
        # Filter out currency codes that sometimes appear in country column
        CURRENCY_CODES = {"SEK", "GBP", "CHF", "EUR", "USD", "DKK", "NOK", "PLN", "CZK", "HUF", "RON"}
        if country in CURRENCY_CODES:
            add("WARNING", i, "Investment Country",
                f"'{name}': Country column contained currency code '{country}' — cleared")
            new_df.at[i, "Investment Country"] = None
            country = ""
        elif country in COUNTRY_NAMES:
            new_df.at[i, "Investment Country"] = COUNTRY_NAMES[country]
            country = COUNTRY_NAMES[country]
        elif len(country) == 2 and country in CC2:
            new_df.at[i, "Investment Country"] = CC2[country]
            country = CC2[country]
        elif len(country) == 3:
            pass  # already ISO3
        elif country and country != "NAN":
            add("WARNING", i, "Investment Country",
                f"'{name}': Country '{country}' not recognised — check manually")

        # Clean junk text from numeric/date fields
        JUNK_PATTERNS = ["we are not collecting", "n/a", "#ref!", "-", ""]
        for col in ["Employees YE", "Turnover YE", "Number of Employees",
                     "Annual Turnover", "Percentage of Equity",
                     "Total Number of Founders/Executives",
                     "Number of Female Founders/Executives",
                     "Number of Ethnic Minority Founders/Executives"]:
            val = new_df.at[i, col]
            if isinstance(val, str):
                if val.strip().lower() in JUNK_PATTERNS or "not collecting" in val.lower():
                    new_df.at[i, col] = None

        # Percentage of Equity: handle 'NA' strings
        pct = new_df.at[i, "Percentage of Equity"]
        if isinstance(pct, str) and pct.strip().upper() in ("NA", "N/A", ""):
            new_df.at[i, "Percentage of Equity"] = None
            add("WARNING", i, "Percentage of Equity", f"'{name}': Ownership is NA")

        # GBR: ALWAYS pull from Companies House
        if country == "GBR":
            co_num = new_df.at[i, "Company Number"]
            co_str = None
            ch = None

            if pd.notna(co_num) and str(co_num).strip() not in ("", "nan", "None", "0", "-"):
                co_raw = str(co_num).strip()
                # Try to zero-pad pure numeric company numbers
                try:
                    co_str = str(int(float(co_raw))).zfill(8)
                except (ValueError, TypeError):
                    co_str = co_raw  # keep as-is (e.g. "HRB 257595 B", "559405-1236")
                new_df.at[i, "Company Number"] = co_str
                ch = companies_house_lookup(co_str)
            else:
                # No Company Number — search by Legal Name or Enterprise Name
                search_name = None
                if i < len(azava_legal_names) and pd.notna(azava_legal_names[i]):
                    search_name = str(azava_legal_names[i]).strip()
                if not search_name:
                    search_name = name

                ch = companies_house_search(search_name)
                if ch and ch.get("company_number"):
                    new_df.at[i, "Company Number"] = ch["company_number"]
                    matched = ch.get("ch_matched_name", "")
                    add("INFO", i, "Company Number",
                        f"'{name}': CH search matched → '{matched}' ({ch['company_number']})")
                else:
                    add("WARNING", i, "Company Number",
                        f"'{name}': UK company — no Company Number, CH search failed for '{search_name}'")

            # SIC: always use CH data
            old_sic = new_df.at[i, "Investment Industry"]
            if ch and ch.get("sic"):
                new_df.at[i, "Investment Industry"] = ch["sic"]
                if pd.notna(old_sic) and str(old_sic).strip() not in ("", "nan", "None"):
                    try:
                        old_sic_str = str(int(float(old_sic))).zfill(5)
                    except (ValueError, TypeError):
                        old_sic_str = str(old_sic).strip()
                    if old_sic_str != ch["sic"]:
                        add("INFO", i, "Investment Industry",
                            f"'{name}': SIC corrected '{old_sic_str}' → '{ch['sic']}' (CH)")
            elif not ch and pd.notna(old_sic) and str(old_sic).strip() not in ("", "nan", "None"):
                try:
                    new_df.at[i, "Investment Industry"] = str(int(float(old_sic))).zfill(5)
                except (ValueError, TypeError):
                    pass
                add("WARNING", i, "Investment Industry",
                    f"'{name}': SIC NOT verified — CH lookup failed")
            elif not ch:
                add("WARNING", i, "Investment Industry", f"'{name}': Missing SIC, CH lookup failed")

            # Postcode: always use CH registered office
            old_pc = str(new_df.at[i, "Investment Location"] or "").strip().upper()
            if ch and ch.get("postcode"):
                new_df.at[i, "Investment Location"] = ch["postcode"]
                if old_pc and old_pc != ch["postcode"].upper() and old_pc != "NONE":
                    add("INFO", i, "Investment Location",
                        f"'{name}': Location set to '{ch['postcode']}' (CH)")
            elif not ch and old_pc and UK_PC_RE.match(old_pc):
                pass  # keep submitted postcode
            elif not ch:
                add("WARNING", i, "Investment Location",
                    f"'{name}': No postcode, CH lookup failed")

        else:
            # Non-GBR: blank out Company Number, SIC, and Location
            # These fields are only meaningful for UK companies via Companies House
            new_df.at[i, "Company Number"] = None
            new_df.at[i, "Investment Industry"] = None
            new_df.at[i, "Investment Location"] = None

        # Investment Type
        inv_type = str(new_df.at[i, "Investment Type"] or "").strip()
        if inv_type and inv_type != "nan" and inv_type not in VALID_INV_TYPES:
            add("WARNING", i, "Investment Type",
                f"'{name}': '{inv_type}' not in BBB permitted values")

        # Investment Stage
        stage = str(new_df.at[i, "Investment Stage"] or "").strip()
        if stage and stage != "nan" and stage not in VALID_STAGES:
            add("WARNING", i, "Investment Stage",
                f"'{name}': '{stage}' not in BBB permitted values")

        # Gross = Unrealised Cost + Realised Cost
        gross = new_df.at[i, "Gross Investment Amount"]
        uc = new_df.at[i, "Unrealised Cost"]
        rc = new_df.at[i, "Realised Cost"]
        if pd.notna(gross) and pd.notna(uc) and pd.notna(rc):
            g, u, r = float(gross), float(uc), float(rc)
            if abs(g - (u + r)) > 0.02:
                add("ERROR", i, "Gross Investment Amount",
                    f"'{name}': Gross ({g:,.2f}) ≠ UC ({u:,.2f}) + RC ({r:,.2f})")

    # Save original per-row currency BEFORE standardisation
    # This is the company's reporting currency, used for turnover columns
    original_row_ccy = []
    for i2 in range(len(new_df)):
        c = str(new_df.at[i2, "Loan / Investment Denomination Currency"] or "").strip().upper()
        original_row_ccy.append(c if c and c != "NAN" else "")

    # =========================================================================
    # FUND DENOMINATION CURRENCY RULES
    # =========================================================================
    # Rule 1: If N-R values had mixed currency symbols → already flagged above
    # Rule 2: Determine fund denomination currency:
    #         a) Explicit override (fund_denomination_ccy param) always wins
    #         b) Currency detected from N-R value formatting
    #         c) Consistent column L
    #         d) Fall back to flag for manual resolution
    # =========================================================================

    raw_L_currencies = [str(c).strip().upper() for c in
                        new_df["Loan / Investment Denomination Currency"].dropna().unique()
                        if str(c).strip().upper() not in ("", "NAN")]

    # Helper: detect fund denomination currency from source file formatting
    def detect_fund_ccy_from_source(fpath, sheet_name):
        """
        Three-pass detection:
        1. openpyxl cell formats on financial columns (checks header rows 1-4)
        2. ALL number format codes in styles.xml (numFmts + DXF)
        3. DXF table styles in raw XML (for Excel Table formatting)
        """
        detected = set()

        def _classify_fmt(nf):
            """Return currency code from a number format string, or None."""
            if not nf or nf == 'General':
                return None
            if '€' in nf or '[$€' in nf: return "EUR"
            if '£' in nf or '[$£' in nf: return "GBP"
            if 'CHF' in nf: return "CHF"
            if 'kr' in nf.lower(): return "SEK"
            if '$' in nf and '[$€' not in nf and '[$£' not in nf:
                import re as _re
                if _re.search(r'\[\$\$|\$[#0]', nf):
                    return "USD"
            return None

        # --- Pass 1: openpyxl cell formats on financial columns ---
        try:
            wb_check = load_workbook(fpath, data_only=True)
            if sheet_name in wb_check.sheetnames:
                ws_check = wb_check[sheet_name]
                financial_headers = ["Gross Investment", "Unrealised Cost",
                                     "Unrealised Value", "Realised Cost", "Realised Value"]
                fin_cols = []
                # Check header rows 1-4 (different variants use different rows)
                for hdr_row in range(1, 5):
                    for c in range(1, min(ws_check.max_column + 1, 60)):
                        h = str(ws_check.cell(hdr_row, c).value or "")
                        if any(fh in h for fh in financial_headers) and "BBB" not in h:
                            fin_cols.append((c, hdr_row))

                openpyxl_ccys = set()
                for c, hdr_row in fin_cols:
                    data_start = hdr_row + 1
                    for r in range(data_start, min(ws_check.max_row + 1, data_start + 500)):
                        val = ws_check.cell(r, c).value
                        if val is None:
                            continue  # skip empty rows — residual formatting pollutes detection
                        nf = ws_check.cell(r, c).number_format or ""
                        ccy = _classify_fmt(nf)
                        if ccy:
                            openpyxl_ccys.add(ccy)

                if openpyxl_ccys:
                    detected = openpyxl_ccys
            wb_check.close()
        except Exception:
            pass

        if detected:
            return detected

        # --- Pass 2: ALL number format codes from styles.xml ---
        # Checks both numFmts (cell-level) and dxfs (table-level)
        try:
            import zipfile, re
            if not fpath.endswith('.xlsx'):
                return set()
            with zipfile.ZipFile(fpath) as z:
                with z.open('xl/styles.xml') as f:
                    styles_content = f.read().decode()

                # Extract ALL formatCode attributes from entire styles.xml
                all_fmts = re.findall(r'formatCode="([^"]*)"', styles_content)
                for fc in all_fmts:
                    fc_d = fc.replace('&quot;', '"').replace('&amp;', '&')
                    ccy = _classify_fmt(fc_d)
                    if ccy:
                        detected.add(ccy)
        except Exception:
            pass

        return detected

    if fund_denomination_ccy:
        # Explicit override or detected from N-R parsing
        fund_ccy = fund_denomination_ccy.upper()
    elif len(raw_L_currencies) == 1:
        # Column L is consistent → use it
        fund_ccy = raw_L_currencies[0]
    elif len(raw_L_currencies) > 1:
        # Column L is mixed → detect from source file formatting
        src_ccys = detect_fund_ccy_from_source(filepath, variant_sheet)
        if len(src_ccys) == 1:
            fund_ccy = list(src_ccys)[0]
            add("INFO", 0, "Currency",
                f"'{fund_label}': Column L mixed ({', '.join(raw_L_currencies)}) but "
                f"financial column formatting is all {fund_ccy} — using as fund denomination")
        else:
            fund_ccy = None
            extra = f" (detected: {', '.join(sorted(src_ccys))})" if src_ccys else " (no formatting detected)"
            add("ERROR", 0, "Currency",
                f"'{fund_label}': Column L shows MIXED CURRENCIES ({', '.join(raw_L_currencies)}){extra} "
                f"— go back to the fund manager. All financial values must be submitted "
                f"in a single fund denomination currency")
    else:
        # No currency info at all
        fund_ccy = None
        add("WARNING", 0, "Currency",
            f"'{fund_label}': No currency information found in column L")

    # Apply fund denomination currency to all rows
    if fund_ccy:
        changed = False
        for i2 in range(len(new_df)):
            old_val = str(new_df.at[i2, "Loan / Investment Denomination Currency"] or "").strip().upper()
            if old_val != fund_ccy:
                changed = True
            new_df.at[i2, "Loan / Investment Denomination Currency"] = fund_ccy
        if changed:
            add("INFO", 0, "Currency",
                f"'{fund_label}': Column L standardised to {fund_ccy}"
                + (f" (was: {', '.join(raw_L_currencies)})" if len(raw_L_currencies) > 1 else ""))

    # Remove blank rows
    mask = new_df["Enterprise Name"].notna() & (new_df["Enterprise Name"].astype(str).str.strip() != "")
    new_df = new_df[mask].reset_index(drop=True)

    # =========================================================================
    # DEDUPLICATE: Consolidate rows with same Enterprise Name + Company Number
    # =========================================================================
    # Financial fields: SUM across duplicates
    # Non-financial fields: keep from most recent Original Investment Date
    # Percentage of Equity: FLAG for manual review
    # =========================================================================

    SUM_COLS = [
        "Gross Investment Amount", "Unrealised Cost", "Unrealised Value",
        "Realised Cost", "Realised Value", "Other Income",
        "Cumulative Roundsize to Date",
    ]

    # Build dedup key: Enterprise Name + Company Number
    # If Company Number is blank, match on Enterprise Name alone
    # Only split if two rows have DIFFERENT non-blank Company Numbers
    def normalise_name(n):
        return str(n).strip().upper()

    def normalise_co(c):
        s = str(c).strip()
        if s in ("", "nan", "None", "0", "00000000"):
            return ""
        return s

    # Group by normalised Enterprise Name first
    name_groups = {}
    for idx in range(len(new_df)):
        n = normalise_name(new_df.at[idx, "Enterprise Name"])
        name_groups.setdefault(n, []).append(idx)

    # Within each name group, check if there are conflicting Company Numbers
    # If all non-blank Company Numbers are the same → one group
    # If different non-blank Company Numbers → split into subgroups
    key_groups = {}
    for name, indices in name_groups.items():
        co_nums = {}
        for idx in indices:
            co = normalise_co(new_df.at[idx, "Company Number"])
            if co:
                co_nums.setdefault(co, []).append(idx)
            else:
                co_nums.setdefault("", []).append(idx)

        non_blank_cos = {k: v for k, v in co_nums.items() if k != ""}

        if len(non_blank_cos) <= 1:
            # All same company (blank or one consistent number) → single group
            key_groups[f"{name}|MERGED"] = indices
        else:
            # Different company numbers under same name → keep separate
            for co, idxs in co_nums.items():
                key_groups[f"{name}|{co}"] = idxs

    duplicates_found = {k: v for k, v in key_groups.items() if len(v) > 1}

    if duplicates_found:
        deduped_rows = []
        deduped_orig_ccy = []
        processed = set()

        # Process in original row order
        for i in range(len(new_df)):
            if i in processed:
                continue

            # Find which group this row belongs to
            my_group = None
            for k, group in key_groups.items():
                if i in group:
                    my_group = group
                    break
            if my_group is None:
                continue

            processed.update(my_group)

            if len(my_group) == 1:
                # No duplicate — keep as-is
                deduped_rows.append(new_df.iloc[i].copy())
                deduped_orig_ccy.append(original_row_ccy[i])
            else:
                # Multiple rows for same company — consolidate
                group_df = new_df.iloc[my_group]
                company_name = new_df.at[my_group[0], "Enterprise Name"]

                # Find row with most recent Original Investment Date
                best_idx = my_group[0]
                best_date = None
                for gi in my_group:
                    d = new_df.at[gi, "Original Investment Date"]
                    if d is not None:
                        if isinstance(d, (datetime, date)):
                            d_cmp = d
                        elif isinstance(d, pd.Timestamp):
                            d_cmp = d.to_pydatetime()
                        else:
                            d_cmp = None
                        if d_cmp and (best_date is None or d_cmp > best_date):
                            best_date = d_cmp
                            best_idx = gi

                # Start with most recent row as base (non-financial fields)
                merged = new_df.iloc[best_idx].copy()

                # Sum financial columns
                for col in SUM_COLS:
                    vals = [float(new_df.at[gi, col]) for gi in my_group
                            if pd.notna(new_df.at[gi, col])]
                    merged[col] = sum(vals) if vals else None

                # Percentage of Equity: flag for manual review
                pct_vals = [new_df.at[gi, "Percentage of Equity"] for gi in my_group
                            if pd.notna(new_df.at[gi, "Percentage of Equity"])]
                if len(pct_vals) > 1:
                    merged["Percentage of Equity"] = pct_vals[-1]  # keep most recent as placeholder
                    add("FLAG", my_group[0], "Percentage of Equity",
                        f"'{company_name}': {len(my_group)} rows consolidated — "
                        f"equity percentages were {[round(float(v)*100 if float(v) <= 1 else float(v), 2) for v in pct_vals]}% "
                        f"— REVIEW MANUALLY")

                # Original Investment Date: keep the EARLIEST (first investment)
                dates = []
                for gi in my_group:
                    d = new_df.at[gi, "Original Investment Date"]
                    if d is not None and isinstance(d, (datetime, date, pd.Timestamp)):
                        dates.append(d)
                if dates:
                    merged["Original Investment Date"] = min(dates)

                # Company Number: use the first non-blank value
                for gi in my_group:
                    co = normalise_co(new_df.at[gi, "Company Number"])
                    if co:
                        merged["Company Number"] = co
                        break

                # Investment Reference: use company name
                merged["Investment Reference"] = company_name

                deduped_rows.append(merged)
                deduped_orig_ccy.append(original_row_ccy[best_idx])

                add("INFO", my_group[0], "Deduplication",
                    f"'{company_name}': {len(my_group)} rows consolidated into 1 "
                    f"(financial values summed, non-financial from most recent investment)")

        new_df = pd.DataFrame(deduped_rows, columns=NEW_COLS).reset_index(drop=True)
        original_row_ccy = deduped_orig_ccy

        total_dupes = sum(len(g) for g in duplicates_found.values())
        print(f"  Deduplicated: {total_dupes} rows → "
              f"{len(duplicates_found)} consolidated | "
              f"{len(new_df)} final rows")

    print(f"  Clean rows: {len(new_df)} | Issues: {len(issues)}")
    return new_df, issues, original_row_ccy, fund_ccy

# ============================================================================
# AUTO-DETECT FUND METADATA FROM FILE
# ============================================================================

def auto_detect_fund(filepath):
    """Detect fund name and denomination currency from a file.
    Returns (fund_label, denomination_currency) or None if not a valid template."""
    import subprocess

    work_path = filepath
    base = os.path.splitext(os.path.basename(filepath))[0]

    # Convert binary formats
    if filepath.endswith('.xlsb') or filepath.endswith('.xls'):
        if not LIBREOFFICE:
            return None
        subprocess.run([LIBREOFFICE, "--headless", "--convert-to", "xlsx",
                        filepath, "--outdir", TEMP_DIR], capture_output=True)
        work_path = f"{TEMP_DIR}/{base}.xlsx"

    # Try opening (fix corrupt files)
    try:
        wb = load_workbook(work_path, data_only=True)
        sheets = wb.sheetnames
        wb.close()
    except (ValueError, Exception):
        if not LIBREOFFICE:
            return None
        subprocess.run([LIBREOFFICE, "--headless", "--convert-to", "xlsx",
                        filepath, "--outdir", TEMP_DIR], capture_output=True)
        work_path = f"{TEMP_DIR}/{base}.xlsx"
        try:
            wb = load_workbook(work_path, data_only=True)
            sheets = wb.sheetnames
            wb.close()
        except Exception:
            return None

    fund_name = None
    denom_ccy = None

    # Check for Azava format first
    for sn in sheets:
        try:
            wb2 = load_workbook(work_path, data_only=True)
            ws = wb2[sn]
            hdrs = [str(ws.cell(1, c).value or "").strip() for c in range(1, min(ws.max_column+1, 30))]
            wb2.close()
            if "Total Invested" in hdrs and "Unrealized Value" in hdrs:
                df = pd.read_excel(work_path, sheet_name=sn, header=0)
                df = df.dropna(how='all')
                # Fund name: most common investor value
                if 'investors' in df.columns:
                    top = df['investors'].value_counts()
                    fund_name = top.index[0] if len(top) > 0 else base
                # Currency: should be uniform for Azava
                if 'Reporting Currency' in df.columns:
                    ccys = df['Reporting Currency'].dropna().unique()
                    denom_ccy = ccys[0] if len(ccys) == 1 else None
                return (fund_name or base, denom_ccy)
        except Exception:
            continue

    # V1: Portfolio Company Data
    if 'Portfolio Company Data' in sheets:
        try:
            df = pd.read_excel(work_path, sheet_name='Portfolio Company Data', header=1)
            fund_name = df['Fund Reference'].dropna().iloc[0] if df['Fund Reference'].notna().any() else None
            if not fund_name and 'Fund Manager Reference' in df.columns:
                fund_name = df['Fund Manager Reference'].dropna().iloc[0] if df['Fund Manager Reference'].notna().any() else None
        except Exception:
            pass
        return (fund_name or base, denom_ccy)

    # V2: BBI_Equity Template
    if 'BBI_Equity Template' in sheets:
        try:
            df = pd.read_excel(work_path, sheet_name='BBI_Equity Template', header=2)
            fund_name = df['Fund Reference'].dropna().iloc[0] if 'Fund Reference' in df.columns and df['Fund Reference'].notna().any() else None
        except Exception:
            pass
        return (fund_name or base, denom_ccy)

    # V5: Isomer Template
    if 'Isomer Template' in sheets:
        try:
            df = pd.read_excel(work_path, sheet_name='Isomer Template', header=0)
            fund_name = df['Fund Manager Reference'].dropna().iloc[0] if 'Fund Manager Reference' in df.columns and df['Fund Manager Reference'].notna().any() else None
        except Exception:
            pass
        return (fund_name or base, denom_ccy)

    # V3: Managed Funds Template
    if 'Managed Funds Template' in sheets:
        try:
            df = pd.read_excel(work_path, sheet_name='Managed Funds Template', header=0)
            fund_name = df['Fund Reference'].dropna().iloc[0] if 'Fund Reference' in df.columns and df['Fund Reference'].notna().any() else None
            if 'Loan / Investment Denomination Currency' in df.columns:
                ccys = df['Loan / Investment Denomination Currency'].dropna().unique()
                denom_ccy = ccys[0] if len(ccys) == 1 else None
        except Exception:
            pass
        return (fund_name or base, denom_ccy)

    return None


def scan_folder(folder_path):
    """Scan a folder for fund templates and auto-detect metadata."""
    VALID_EXT = {'.xlsx', '.xlsb', '.xls', '.xlsm'}
    files = []
    for f in sorted(os.listdir(folder_path)):
        ext = os.path.splitext(f)[1].lower()
        if ext in VALID_EXT and not f.startswith('~') and not f.startswith('.'):
            # Skip our own output files
            if f.startswith('BBB_Consolidated'):
                continue
            full = os.path.join(folder_path, f)
            files.append(full)

    print(f"\n  📁 Scanning: {folder_path}")
    print(f"  Found {len(files)} Excel files\n")

    funds = []
    for fp in files:
        fname = os.path.basename(fp)
        result = auto_detect_fund(fp)
        if result:
            label, ccy = result
            print(f"    ✓ {fname:45s} → {label} (ccy: {ccy or 'auto'})")
            funds.append((fp, label, ccy))
        else:
            print(f"    ✗ {fname:45s} → SKIPPED (unrecognised format)")
    return funds


# ============================================================================
# MAIN: PROCESS ALL FUNDS
# ============================================================================

import sys
from datetime import datetime

# Folder mode: pass folder path as argument, or use default
if len(sys.argv) > 1:
    INPUT_FOLDER = sys.argv[1]
else:
    INPUT_FOLDER = "/mnt/user-data/uploads"

# Output path with timestamp
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
if len(sys.argv) > 2:
    OUTPUT = sys.argv[2]
else:
    # Try input folder first, fall back to /home/claude
    out_dir = INPUT_FOLDER
    try:
        test_path = os.path.join(out_dir, ".write_test")
        with open(test_path, "w") as f: f.write("test")
        os.remove(test_path)
    except (PermissionError, OSError):
        out_dir = TEMP_DIR
    OUTPUT = os.path.join(out_dir, f"BBB_Consolidated_{timestamp}.xlsx")

# Optional config file: fund_config.json in input folder
# Format: {"filename.xlsx": {"label": "Fund Name", "currency": "EUR"}}
config = {}
config_path = os.path.join(INPUT_FOLDER, "fund_config.json")
if os.path.exists(config_path):
    with open(config_path) as f:
        config = json.load(f)
    print(f"  📋 Loaded config: {config_path} ({len(config)} overrides)")

print("=" * 65)
print("  BBB CONSOLIDATED REPORT — AUTOMATED")
print(f"  {datetime.now().strftime('%d %B %Y %H:%M')}")
print("=" * 65)

FUNDS = scan_folder(INPUT_FOLDER)

# Apply config overrides
for i, (fp, label, ccy) in enumerate(FUNDS):
    fname = os.path.basename(fp)
    if fname in config:
        c = config[fname]
        new_label = c.get("label", label)
        new_ccy = c.get("currency", ccy)
        if new_label != label or new_ccy != ccy:
            print(f"    ↳ Config override: {fname} → {new_label} ({new_ccy or 'auto'})")
            FUNDS[i] = (fp, new_label, new_ccy)

if not FUNDS:
    print("\n  ✗ No valid fund templates found. Exiting.")
    sys.exit(1)

print(f"\n  Processing {len(FUNDS)} funds...\n")
import time
_start = time.time()

fund_results = []  # list of (df, issues, orig_ccy, label, fund_ccy)
for filepath, label, denom_ccy in FUNDS:
    df, issues, orig_ccy, fund_ccy = process_fund(filepath, label, fund_denomination_ccy=denom_ccy)
    if len(df) > 0:
        fund_results.append((df, issues, orig_ccy, label, fund_ccy))
    else:
        print(f"  ⚠ {label}: 0 rows — skipped (blank template?)")

# Combine all funds
combined = pd.concat([r[0] for r in fund_results], ignore_index=True)
combined_orig_ccy = []
for r in fund_results:
    combined_orig_ccy.extend(r[2])

# Assign absolute row numbers and fund labels to issues
all_issues = []
offset = 0
for df, issues, orig_ccy, label, fund_ccy in fund_results:
    for iss in issues:
        iss["row_abs"] = iss["row"] + offset + 2
        iss["fund"] = label
        all_issues.append(iss)
    offset += len(df)

# Track fund boundaries for row colouring + currency info
fund_boundaries = []  # (start, end, label, fund_ccy)
running = 0
for df, _, _, label, fund_ccy in fund_results:
    fund_boundaries.append((running, running + len(df), label, fund_ccy))
    running += len(df)

sizes = " + ".join(str(len(r[0])) for r in fund_results)
print(f"\n{'='*65}")
print(f"  COMBINED: {len(combined)} rows ({sizes})")
print(f"  Total issues: {len(all_issues)}")
print(f"{'='*65}")

# ============================================================================
# BUILD CLEAN OUTPUT WORKBOOK
# ============================================================================

print("\nBuilding workbook...")
wb = Workbook()

# -- Sheet 1: Managed Funds Template --
ws = wb.active
ws.title = "Managed Funds Template"

hdr_font = Font(bold=True, size=10, name="Arial", color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor="1F4E79")
dfont = Font(size=10, name="Arial")
tb = Border(left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin"))

for ci, col_name in enumerate(NEW_COLS, 1):
    cell = ws.cell(row=1, column=ci, value=col_name)
    cell.font = hdr_font; cell.fill = hdr_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = tb

for ci, col_name in enumerate(NEW_COLS, 1):
    ws.column_dimensions[get_column_letter(ci)].width = min(max(len(col_name) + 2, 12), 30)
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(NEW_COLS))}1"

money_cols = {14,15,16,17,18,32,33,34}  # fund denomination ccy (BBB cols 19-23 left blank)
company_money_cols = {26, 29}  # Annual Turnover + Turnover YE — company-level, mixed ccy OK
bbb_cols = {19,20,21,22,23}  # BBB columns — left blank per LP instructions
dt_cols = {2, 10}
pct_cols = {9, 25}

def mfmt(ccy):
    return {"GBP": '_-£#,##0.00_-;-£#,##0.00_-;_-£"-"??_-;_-@_-',
            "EUR": '_-€#,##0.00_-;-€#,##0.00_-;_-€"-"??_-;_-@_-',
            "USD": '_-$#,##0.00_-;-$#,##0.00_-;_-$"-"??_-;_-@_-',
            "SEK": '#,##0.00 "kr"', "DKK": '#,##0.00 "kr"', "NOK": '#,##0.00 "kr"',
            "CHF": '_-"CHF "#,##0.00_-;-"CHF "#,##0.00_-;_-"CHF ""-"??_-;_-@_-',
            }.get(str(ccy).strip().upper(), '#,##0.00')

# Fund separator fills — cycle through colours
FUND_FILLS = [
    PatternFill("solid", fgColor="EBF5FB"),  # light blue
    PatternFill("solid", fgColor="FEF9E7"),  # light yellow
    PatternFill("solid", fgColor="E8F8F5"),  # light green
    PatternFill("solid", fgColor="FDEDEC"),  # light red
    PatternFill("solid", fgColor="F4ECF7"),  # light purple
    PatternFill("solid", fgColor="FDEBD0"),  # light orange
]

def get_fund_fill(row_idx):
    for fi, (start, end, label, fund_ccy) in enumerate(fund_boundaries):
        if start <= row_idx < end:
            return FUND_FILLS[fi % len(FUND_FILLS)]
    return PatternFill()

def get_fund_ccy(row_idx):
    """Return the resolved fund currency for this row, or None if mixed."""
    for fi, (start, end, label, fund_ccy) in enumerate(fund_boundaries):
        if start <= row_idx < end:
            return fund_ccy
    return None

for i in range(len(combined)):
    er = i + 2
    resolved_fund_ccy = get_fund_ccy(i)
    row_ccy = str(combined.at[i, "Loan / Investment Denomination Currency"] or "").strip().upper()

    # Always show currency symbol:
    # - Resolved fund currency → use that (consistent across fund)
    # - Unresolved → use per-row currency from column L (still flagged as ERROR)
    if resolved_fund_ccy:
        mf = mfmt(resolved_fund_ccy)
    elif row_ccy and row_ccy not in ("", "NAN"):
        mf = mfmt(row_ccy)
    else:
        mf = '#,##0.00'

    # Company-level turnover uses the original per-row currency
    co_ccy = combined_orig_ccy[i] if i < len(combined_orig_ccy) and combined_orig_ccy[i] else row_ccy
    co_mf = mfmt(co_ccy) if co_ccy and co_ccy not in ("", "NAN") else '#,##0.00'
    row_fill = get_fund_fill(i)

    for ci, cn in enumerate(NEW_COLS, 1):
        val = combined.at[i, cn]
        if pd.isna(val): val = None
        # Zero → blank for financial columns
        if ci in money_cols or ci in company_money_cols:
            if isinstance(val, (int, float)) and (val == 0 or abs(val) < 0.001):
                val = None
        # BBB columns always blank
        if ci in bbb_cols:
            val = None
        cell = ws.cell(row=er, column=ci, value=val)
        cell.font = dfont; cell.border = tb; cell.fill = row_fill
        if ci in money_cols: cell.number_format = mf
        elif ci in company_money_cols: cell.number_format = co_mf
        elif ci in dt_cols: cell.number_format = "DD/MM/YYYY"
        elif ci in pct_cols: cell.number_format = "0.00%"

# -- Sheet 2: BBI Ownership --
ws_bbi = wb.create_sheet("BBI Ownership")
bbi_headers = ["Fund Reference", "BBB Participation %"]
for ci, h in enumerate(bbi_headers, 1):
    c = ws_bbi.cell(row=1, column=ci, value=h)
    c.font = hdr_font; c.fill = hdr_fill; c.border = tb
ws_bbi.cell(row=2, column=1, value="(Complete once all funds consolidated)")
ws_bbi.column_dimensions["A"].width = 30; ws_bbi.column_dimensions["B"].width = 25
ws_bbi.freeze_panes = "A2"

# -- Sheet 3: Lists --
ws_lists = wb.create_sheet("Lists")
list_data = {
    "Investment Type": VALID_INV_TYPES,
    "Investment Stage": VALID_STAGES,
    "Exit Type": ["Trade Sale", "IPO", "Secondary Sale", "Write-off",
                   "Management Buyback", "Other"],
}
col_off = 1
for header, values in list_data.items():
    c = ws_lists.cell(row=1, column=col_off, value=header)
    c.font = hdr_font; c.fill = hdr_fill; c.border = tb
    for ri, v in enumerate(values, 2):
        ws_lists.cell(row=ri, column=col_off, value=v).border = tb
    ws_lists.column_dimensions[get_column_letter(col_off)].width = 28
    col_off += 1
ws_lists.freeze_panes = "A2"

# -- Sheet 4: Validation Report --
ws_v = wb.create_sheet("Validation Report")
vhdr_fill = PatternFill("solid", fgColor="C00000")
for ci, h in enumerate(["Severity", "Fund", "Row", "Column", "Issue"], 1):
    c = ws_v.cell(row=1, column=ci, value=h)
    c.font = hdr_font; c.fill = vhdr_fill
    c.alignment = Alignment(horizontal="center"); c.border = tb

sf = {"ERROR": Font(size=10, name="Arial", color="C00000"),
      "FLAG": Font(size=10, name="Arial", color="0070C0"),
      "WARNING": Font(size=10, name="Arial", color="BF8F00"),
      "INFO": Font(size=10, name="Arial", color="808080")}
so = {"ERROR": 0, "FLAG": 1, "WARNING": 2, "INFO": 3}

for ri, iss in enumerate(sorted(all_issues, key=lambda x: so.get(x["sev"], 4)), 2):
    ws_v.cell(row=ri, column=1, value=iss["sev"]).border = tb
    ws_v.cell(row=ri, column=2, value=iss["fund"]).border = tb
    ws_v.cell(row=ri, column=3, value=iss["row_abs"]).border = tb
    ws_v.cell(row=ri, column=4, value=iss["col"]).border = tb
    ws_v.cell(row=ri, column=5, value=iss["msg"]).border = tb
    f = sf.get(iss["sev"], sf["INFO"])
    for c in range(1, 6): ws_v.cell(row=ri, column=c).font = f

ws_v.column_dimensions["A"].width = 12
ws_v.column_dimensions["B"].width = 25
ws_v.column_dimensions["C"].width = 8
ws_v.column_dimensions["D"].width = 28
ws_v.column_dimensions["E"].width = 100
ws_v.freeze_panes = "A2"

# -- Sheet 5: New Companies --
# A "new company" = Original Investment Date falls within the reporting quarter
# Determine quarter from the most common Period End Date
ws_n = wb.create_sheet("New Companies")
ncfill = PatternFill("solid", fgColor="2E75B6")
nc_headers = ["Enterprise Name", "Fund Reference", "Investment Type",
              "Original Investment Date", "Gross Investment Amount",
              "Investment Country", "Investment Industry", "Row"]
for ci, h in enumerate(nc_headers, 1):
    c = ws_n.cell(row=1, column=ci, value=h)
    c.font = hdr_font; c.fill = ncfill; c.border = tb
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Find reporting quarter from Period End Date
period_dates = combined["Period End Date"].dropna().unique()
quarter_end = None
for pd_val in period_dates:
    if isinstance(pd_val, (datetime, date)):
        quarter_end = pd_val if isinstance(pd_val, date) else pd_val.date()
        break
    elif hasattr(pd_val, 'date'):  # pandas Timestamp
        quarter_end = pd_val.date()
        break

new_companies = []
if quarter_end:
    # Quarter start = 3 months before quarter end
    q_month = quarter_end.month
    if q_month in (1, 2, 3):
        quarter_start = date(quarter_end.year, 1, 1)
    elif q_month in (4, 5, 6):
        quarter_start = date(quarter_end.year, 4, 1)
    elif q_month in (7, 8, 9):
        quarter_start = date(quarter_end.year, 7, 1)
    else:
        quarter_start = date(quarter_end.year, 10, 1)

    # Convert to pd.Timestamp for safe comparison
    qs = pd.Timestamp(quarter_start)
    qe = pd.Timestamp(quarter_end)

    for i in range(len(combined)):
        oid = combined.at[i, "Original Investment Date"]
        try:
            oid_ts = pd.Timestamp(oid)
            if pd.notna(oid_ts) and qs <= oid_ts <= qe:
                new_companies.append(i)
        except (ValueError, TypeError):
            pass

if new_companies:
    for ri, idx in enumerate(new_companies, 2):
        fund_ref = combined.at[idx, "Fund Reference"]
        # Determine fund ccy for formatting
        row_fund_ccy = get_fund_ccy(idx)
        nc_mf = mfmt(row_fund_ccy) if row_fund_ccy else '#,##0.00'

        ws_n.cell(row=ri, column=1, value=combined.at[idx, "Enterprise Name"]).border = tb
        ws_n.cell(row=ri, column=2, value=fund_ref).border = tb
        ws_n.cell(row=ri, column=3, value=combined.at[idx, "Investment Type"]).border = tb
        oid_cell = ws_n.cell(row=ri, column=4, value=combined.at[idx, "Original Investment Date"])
        oid_cell.number_format = "DD/MM/YYYY"; oid_cell.border = tb
        gross_cell = ws_n.cell(row=ri, column=5, value=combined.at[idx, "Gross Investment Amount"])
        gross_cell.number_format = nc_mf; gross_cell.border = tb
        ws_n.cell(row=ri, column=6, value=combined.at[idx, "Investment Country"]).border = tb
        ws_n.cell(row=ri, column=7, value=combined.at[idx, "Investment Industry"]).border = tb
        ws_n.cell(row=ri, column=8, value=idx + 2).border = tb  # row in main sheet
    print(f"\n  📋 New Companies: {len(new_companies)} companies invested in "
          f"{quarter_start.strftime('%d/%m/%Y')} – {quarter_end.strftime('%d/%m/%Y')}")
else:
    ws_n.cell(row=2, column=1, value="No new companies identified this quarter")

for ci in range(1, len(nc_headers) + 1):
    ws_n.column_dimensions[get_column_letter(ci)].width = 25
ws_n.freeze_panes = "A2"
ws_n.auto_filter.ref = f"A1:{get_column_letter(len(nc_headers))}1"

# ============================================================================
# SAVE
# ============================================================================

wb.save(OUTPUT)

# ============================================================================
# SUMMARY
# ============================================================================

errors = [x for x in all_issues if x["sev"] == "ERROR"]
warnings = [x for x in all_issues if x["sev"] == "WARNING"]
flags = [x for x in all_issues if x["sev"] == "FLAG"]
infos = [x for x in all_issues if x["sev"] == "INFO"]

print(f"\n{'='*65}")
print(f"  CONSOLIDATED SUMMARY")
print(f"{'='*65}")
for df, _, _, label, _ in fund_results:
    print(f"  {label:30s} {len(df)} rows")
print(f"  {'TOTAL':30s} {len(combined)} rows")

# Per-fund error breakdown
for _, _, _, fund_name, _ in fund_results:
    fe = [x for x in errors if x["fund"] == fund_name]
    fw = [x for x in warnings if x["fund"] == fund_name]
    ff = [x for x in flags if x["fund"] == fund_name]
    fi = [x for x in infos if x["fund"] == fund_name]
    print(f"\n  {fund_name}:")
    if fe:
        ue = {}
        for e in fe: ue.setdefault(e["col"], []).append(e)
        for col, items in ue.items():
            print(f"    ✗ ERROR: {col} ({len(items)} rows)")
    if fw:
        uw = {}
        for w in fw:
            k = w["msg"].split("':")[1].strip()[:50] if "':" in w["msg"] else w["col"]
            uw.setdefault(k, []).append(w)
        for desc, items in uw.items():
            print(f"    ⚠ WARNING: {desc[:55]} ({len(items)}x)")
    if ff:
        for f in ff: print(f"    ⬤ FLAG: {f['msg'][:70]}")
    if fi:
        for info in fi[:3]: print(f"    ✓ AUTO-FIX: {info['msg'][:70]}")
        if len(fi) > 3: print(f"    ... +{len(fi)-3} more")

print(f"\n  → {OUTPUT}")
elapsed = time.time() - _start
print(f"  ⏱  Completed in {elapsed:.1f}s")
print(f"{'='*65}")
