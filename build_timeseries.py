#!/usr/bin/env python3
"""
BBB Time-Series Builder
=======================
Reads multiple quarterly BBB Consolidated xlsx files and stacks them
into a single JSON time-series for the Manager Dashboard.

Usage:
    python3 build_timeseries.py /path/to/quarterly_files/ [-o output.json]

Input:  Folder containing BBB_Consolidated_*.xlsx files (one per quarter)
Output: dashboard_data.json — time-series data for the React dashboard

The script:
1. Reads each quarterly file's "Managed Funds Template" sheet
2. Identifies the quarter from the Period End Date column
3. Matches companies across quarters by Enterprise Name + Company Number
4. Builds per-company time series of key metrics
5. Outputs structured JSON for the dashboard
"""

import os, sys, json, re, glob
from datetime import datetime
from collections import defaultdict
import warnings
warnings.filterwarnings("ignore")

try:
    import pandas as pd
    import numpy as np
    from openpyxl import load_workbook
except ImportError:
    print("Install required packages: pip install pandas openpyxl numpy --break-system-packages")
    sys.exit(1)


def quarter_label(dt):
    """Convert a date to 'Q1 2024' format."""
    if pd.isna(dt) or dt is None:
        return None
    if isinstance(dt, str):
        try:
            dt = pd.to_datetime(dt, dayfirst=True)
        except Exception:
            return None
    q = (dt.month - 1) // 3 + 1
    return f"Q{q} {dt.year}"


def safe_float(val):
    """Convert to float, returning None for blanks/errors."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    try:
        v = float(val)
        return v if v != 0 else None
    except (ValueError, TypeError):
        return None


def normalise_name(n):
    """Normalise company name for matching."""
    s = str(n).strip().upper()
    # Remove common suffixes
    for suffix in [" LIMITED", " LTD", " LTD.", " INC", " INC.", " GMBH",
                   " SAS", " S.L.", " S.R.L.", " AG", " PLC", " LP", " LLP"]:
        if s.endswith(suffix):
            s = s[:-len(suffix)].strip()
    return s


def normalise_co_num(c):
    """Normalise company number."""
    s = str(c).strip()
    if s in ("", "nan", "None", "0", "00000000"):
        return ""
    return s


def read_quarterly_file(filepath):
    """Read a BBB Consolidated xlsx and extract company data."""
    wb = load_workbook(filepath, data_only=True)
    ws = wb["Managed Funds Template"]

    # Read headers from row 1
    headers = {}
    for c in range(1, ws.max_column + 1):
        h = ws.cell(1, c).value
        if h:
            headers[str(h).strip()] = c

    required = ["Enterprise Name", "Fund Reference"]
    for r in required:
        if r not in headers:
            print(f"  ⚠ Skipping {os.path.basename(filepath)}: missing '{r}' column")
            wb.close()
            return []

    rows = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, headers.get("Enterprise Name", 5)).value
        if not name or str(name).strip() == "":
            continue

        # Extract key fields
        fund = ws.cell(r, headers.get("Fund Reference", 1)).value
        co_num = ws.cell(r, headers.get("Company Number", 6)).value
        country = ws.cell(r, headers.get("Investment Country", 7)).value
        industry = ws.cell(r, headers.get("Investment Industry", 8)).value
        inv_type = ws.cell(r, headers.get("Investment Type", 11)).value
        stage = ws.cell(r, headers.get("Investment Stage", 10)).value
        pct_equity = safe_float(ws.cell(r, headers.get("Percentage of Equity Held", 9)).value)
        inv_date = ws.cell(r, headers.get("Original Investment Date", 3)).value
        period_end = ws.cell(r, headers.get("Period End Date", 2)).value
        currency = ws.cell(r, headers.get("Loan / Investment Denomination Currency", 12)).value

        gross = safe_float(ws.cell(r, headers.get("Gross Investment Amount", 14)).value)
        uc = safe_float(ws.cell(r, headers.get("Unrealised Cost", 15)).value)
        uv = safe_float(ws.cell(r, headers.get("Unrealised Value", 16)).value)
        rc = safe_float(ws.cell(r, headers.get("Realised Cost", 17)).value)
        rv = safe_float(ws.cell(r, headers.get("Realised Value", 18)).value)
        turnover = safe_float(ws.cell(r, headers.get("Annual Turnover", 26)).value)
        ev = safe_float(ws.cell(r, headers.get("Enterprise Value", 33)).value)
        roundsize = safe_float(ws.cell(r, headers.get("Cumulative Roundsize to Date", 32)).value)

        # Calculate MOIC
        moic = None
        if uv is not None and uc is not None and uc > 0:
            moic = round(uv / uc, 2)

        quarter = quarter_label(period_end)

        rows.append({
            "name": str(name).strip(),
            "name_norm": normalise_name(name),
            "co_num": normalise_co_num(co_num),
            "fund": str(fund).strip() if fund else "",
            "country": str(country).strip() if country else "",
            "industry": str(industry).strip() if industry else "",
            "inv_type": str(inv_type).strip() if inv_type else "",
            "stage": str(stage).strip() if stage else "",
            "pct_equity": pct_equity,
            "inv_date": str(inv_date)[:10] if inv_date else None,
            "currency": str(currency).strip().upper() if currency else "",
            "quarter": quarter,
            "gross": gross,
            "unrealised_cost": uc,
            "unrealised_value": uv,
            "realised_cost": rc,
            "realised_value": rv,
            "turnover": turnover,
            "enterprise_value": ev,
            "roundsize": roundsize,
            "moic": moic,
        })

    wb.close()
    return rows


def build_company_id(row):
    """Build a stable company ID from name + company number."""
    name = row["name_norm"]
    co = row["co_num"]
    if co:
        return f"{name}|{co}"
    return name


def build_timeseries(folder, output_path):
    """Main pipeline: read quarterly files and build time-series JSON."""
    # Find all BBB consolidated files
    patterns = [
        os.path.join(folder, "BBB_Consolidated_*.xlsx"),
        os.path.join(folder, "BBB_*.xlsx"),
    ]
    files = []
    for p in patterns:
        files.extend(glob.glob(p))
    files = sorted(set(files))

    if not files:
        print(f"No BBB_Consolidated_*.xlsx files found in {folder}")
        sys.exit(1)

    print(f"Found {len(files)} quarterly file(s):")
    for f in files:
        print(f"  • {os.path.basename(f)}")

    # Read all files
    all_rows = []
    for f in files:
        print(f"\nReading {os.path.basename(f)}...")
        rows = read_quarterly_file(f)
        print(f"  → {len(rows)} companies")
        all_rows.extend(rows)

    if not all_rows:
        print("No data found.")
        sys.exit(1)

    # Build company registry — match across quarters
    companies = {}  # company_id → company info
    timeseries = defaultdict(list)  # company_id → [quarterly snapshots]

    for row in all_rows:
        cid = build_company_id(row)

        # Register company (keep most recent info)
        if cid not in companies:
            companies[cid] = {
                "id": cid,
                "name": row["name"],
                "co_num": row["co_num"],
                "fund": row["fund"],
                "country": row["country"],
                "industry": row["industry"],
                "inv_type": row["inv_type"],
                "inv_date": row["inv_date"],
                "currency": row["currency"],
            }
        else:
            # Update with latest non-blank values
            for k in ["fund", "country", "industry", "inv_type", "stage", "currency"]:
                if row.get(k) and row[k] not in ("", "nan", "None"):
                    companies[cid][k] = row[k]

        # Add quarterly snapshot
        timeseries[cid].append({
            "quarter": row["quarter"],
            "stage": row["stage"],
            "pct_equity": row["pct_equity"],
            "gross": row["gross"],
            "unrealised_cost": row["unrealised_cost"],
            "unrealised_value": row["unrealised_value"],
            "realised_cost": row["realised_cost"],
            "realised_value": row["realised_value"],
            "turnover": row["turnover"],
            "enterprise_value": row["enterprise_value"],
            "roundsize": row["roundsize"],
            "moic": row["moic"],
        })

    # Sort each company's timeseries by quarter
    def quarter_sort_key(q):
        if not q:
            return (0, 0)
        match = re.match(r'Q(\d) (\d{4})', q)
        if match:
            return (int(match.group(2)), int(match.group(1)))
        return (0, 0)

    for cid in timeseries:
        timeseries[cid].sort(key=lambda x: quarter_sort_key(x["quarter"]))

    # Compute derived metrics per company
    company_list = []
    for cid, info in companies.items():
        ts = timeseries[cid]
        latest = ts[-1] if ts else {}

        # Current stage
        info["current_stage"] = latest.get("stage", "")

        # Latest metrics
        info["latest_moic"] = latest.get("moic")
        info["latest_turnover"] = latest.get("turnover")
        info["latest_ev"] = latest.get("enterprise_value")
        info["latest_gross"] = latest.get("gross")
        info["latest_uv"] = latest.get("unrealised_value")
        info["latest_uc"] = latest.get("unrealised_cost")

        # Revenue growth (QoQ if 2+ quarters with turnover data)
        turnovers = [(s["quarter"], s["turnover"]) for s in ts if s["turnover"] is not None]
        if len(turnovers) >= 2:
            prev, curr = turnovers[-2][1], turnovers[-1][1]
            if prev and prev > 0:
                info["revenue_growth_pct"] = round((curr - prev) / prev * 100, 1)
            else:
                info["revenue_growth_pct"] = None
        else:
            info["revenue_growth_pct"] = None

        # Stage changes (for round progression)
        stages = [(s["quarter"], s["stage"]) for s in ts if s.get("stage")]
        stage_transitions = []
        for j in range(1, len(stages)):
            if stages[j][1] != stages[j-1][1]:
                stage_transitions.append({
                    "from": stages[j-1][1],
                    "to": stages[j][1],
                    "quarter": stages[j][0],
                })
        info["stage_transitions"] = stage_transitions

        # Quarters tracked
        info["quarters_tracked"] = len(ts)
        info["timeseries"] = ts

        company_list.append(info)

    # Sort by fund, then name
    company_list.sort(key=lambda x: (x["fund"], x["name"]))

    # Summary stats
    quarters = sorted(set(s["quarter"] for row in all_rows if row["quarter"]
                          for s in [row]),
                      key=quarter_sort_key)

    funds = sorted(set(c["fund"] for c in company_list if c["fund"]))

    output = {
        "generated_at": datetime.now().isoformat(),
        "quarters": quarters,
        "funds": funds,
        "total_companies": len(company_list),
        "companies": company_list,
    }

    with open(output_path, "w") as f:
        json.dump(output, f, indent=2, default=str)

    print(f"\n✅ Built time-series for {len(company_list)} companies across {len(quarters)} quarter(s)")
    print(f"   Funds: {', '.join(funds)}")
    print(f"   Quarters: {', '.join(quarters)}")
    print(f"   → {output_path}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 build_timeseries.py /path/to/quarterly_bbbs/ [-o output.json]")
        sys.exit(1)

    folder = sys.argv[1]
    output = "dashboard_data.json"

    if "-o" in sys.argv:
        idx = sys.argv.index("-o")
        if idx + 1 < len(sys.argv):
            output = sys.argv[idx + 1]

    build_timeseries(folder, output)
