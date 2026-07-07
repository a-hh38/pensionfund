import pandas as pd
from pathlib import Path
import re

# ==========================================
# CONFIG
# ==========================================

DOWNLOAD_DIR = Path("sbi_disclosures")
DOWNLOAD_DIR.mkdir(exist_ok=True)

INPUT_FOLDER = DOWNLOAD_DIR
TARGET_SHEETS =[
    "Scheme CG",
    "Scheme SG",
    "Scheme E - Tier I",
    "Scheme E - Tier II",
    "NPS Lite",
    "Corporate-CG Scheme",
    "APY",
    "Scheme  NPS TTS-II",
    "NPS-II COMPOSITE",
    "UPS CG",
    "UPSPOOLCG",
    "NPS JEEVAN SWARNA",
    "NPS AKSHAY DHARA",
    "NPS VATSALYA",
    "Scheme CG",
    "Scheme SG",
    "Scheme E Tier I",
    "Scheme E Tier II",
    "Scheme NPS Lite",
    "Schem Corp CG",
    "Scheme Atal Pension Yojana",
    "Scheme Tax Saver Tier II",
    "Scheme NPS-II COMPOSITE",
    "Scheme UPS CG",
    "Schme Jeevan Swarna",
    "Scheme Akshay Dhara",
    "Scheme NPS Vatsalya"
]

STOP_WORDS = [
    "debt",
    "government",
    "money market",
    "treps",
    "mutual fund",
    "asset backed",
    "alternative",
    "corporate bond",
    "corporate debt"
]

MONTH_MAP = {
    "january":1,"jan":1,
    "february":2,"feburary":2,"feb":2,
    "march":3,"mar":3,
    "april":4,"apr":4,
    "may":5,
    "june":6,"jun":6,
    "july":7,"jul":7,
    "august":8,"aug":8,
    "september":9,"sept":9,"sep":9,
    "october":10,"oct":10,
    "november":11,"nov":11,
    "december":12,"dec":12
}


# ==========================================
# DATE EXTRACTION
# ==========================================

def extract_report_date(filename):

    name = filename.lower()

    month = None
    month_text = None

    for m, num in MONTH_MAP.items():
        if m in name:
            month = num
            month_text = m
            break

    if month is None:
        return None

    year_match = re.search(r'20\d{2}', name)

    if year_match:
        return pd.Timestamp(
            int(year_match.group()),
            month,
            1
        )

    pos = name.find(month_text)

    window = name[pos:pos+15]

    yy_match = re.search(r'(\d{2})', window)

    if yy_match:

        yy = int(yy_match.group())

        if 20 <= yy <= 30:
            return pd.Timestamp(
                2000 + yy,
                month,
                1
            )

    return None

import requests
import re
from pathlib import Path
from datetime import datetime

JSON_URL = "https://www.sbipensionfunds.co.in/page-data/disclosures/portfolio-details/page-data.json"
BASE_URL = "https://www.sbipensionfunds.co.in"
DOWNLOAD_DIR = Path("sbi_disclosures")
DOWNLOAD_DIR.mkdir(exist_ok=True)

# ==========================
# DOWNLOAD JSON
# ==========================
def main(start_date=None, end_date=None):

    if start_date is None:
        start_date = pd.Timestamp(2025, 4, 1)

    if end_date is None:
        end_date = pd.Timestamp.today().replace(day=1)
    
    # Clean old downloads
    if DOWNLOAD_DIR.exists():

        for file in DOWNLOAD_DIR.rglob("*"):
            if file.is_file():
                try:
                    file.unlink()
                except:
                    pass

        for folder in sorted(DOWNLOAD_DIR.rglob("*"), reverse=True):
            if folder.is_dir():
                try:
                    folder.rmdir()
                except:
                    pass

    DOWNLOAD_DIR.mkdir(exist_ok=True)
    print("Downloading JSON...")
    data = requests.get(JSON_URL).json()
    master = []
    # ==========================
    # EXTRACT XLSX URLS
    # ==========================

    xlsx_urls = set()

    def extract_xlsx(obj):
        if isinstance(obj, dict):
            for v in obj.values():

                if isinstance(v, str) and v.lower().endswith(".xlsx"):

                    if v.startswith("/"):
                        xlsx_urls.add(BASE_URL + v)

                extract_xlsx(v)

        elif isinstance(obj, list):
            for item in obj:
                extract_xlsx(item)

    extract_xlsx(data)

    print(f"Found {len(xlsx_urls)} xlsx URLs")

    # ==========================
    # MONTH PARSER
    # ==========================

    month_map = {
        "january":1,"jan":1,
        "february":2,"feburary":2,"feb":2,
        "march":3,"mar":3,
        "april":4,"apr":4,
        "may":5,
        "june":6,"jun":6,
        "july":7,"jul":7,
        "august":8,"aug":8,
        "september":9,"sept":9,"sep":9,
        "october":10,"oct":10,
        "november":11,"nov":11,
        "december":12,"dec":12
    }

    def extract_month_year(filename):

        name = filename.lower()

        month_patterns = {
            "jan": 1, "january": 1,
            "feb": 2, "february": 2, "feburary": 2,
            "mar": 3, "march": 3,
            "apr": 4, "april": 4,
            "may": 5,
            "jun": 6, "june": 6,
            "jul": 7, "july": 7,
            "aug": 8, "august": 8,
            "sep": 9, "sept": 9, "september": 9,
            "oct": 10, "october": 10,
            "nov": 11, "november": 11,
            "dec": 12, "december": 12
        }

        month = None
        month_text = None

        for m, num in month_patterns.items():
            if m in name:
                month = num
                month_text = m
                break

        if month is None:
            return None

        # Look for explicit 2024/2025/2026 first

        year_match = re.search(r'20\d{2}', name)

        if year_match:
            return datetime(int(year_match.group()), month, 1)

        # ONLY look in a small window after the month name

        pos = name.find(month_text)

        window = name[pos:pos + 15]

        yy_match = re.search(r'(\d{2})', window)

        if yy_match:

            yy = int(yy_match.group(1))

            if 20 <= yy <= 30:
                return datetime(2000 + yy, month, 1)

        return None

    # ==========================
    # BUILD RECORDS
    # ==========================

    records = []

    for url in xlsx_urls:

        fname = Path(url).name

        dt = extract_month_year(fname)

        if dt:
            records.append({
                "date": dt,
                "file": fname,
                "url": url
            })

    print(f"\nMatched {len(records)} files to dates")

    # ==========================
    # DEDUPE BY MONTH
    # Prefer cleaner filename
    # ==========================

    monthly = {}

    for r in sorted(records, key=lambda x: len(x["file"])):

        key = r["date"].strftime("%Y-%m")

        if key not in monthly:
            monthly[key] = r

    records = list(monthly.values())
    records.sort(key=lambda x: x["date"])

    print("\nMonths found:")

    for r in records:
        print(r["date"].strftime("%Y-%m"), "->", r["file"])

    # ==========================
    # LAST 25 MONTHS
    # ==========================

    records = [
    r
    for r in records
    if start_date <= pd.Timestamp(r["date"]) <= end_date
    ]

    print("\nSelected months:")

    for r in records:
        print(r["date"].strftime("%Y-%m"))

    # ==========================
    # DOWNLOAD
    # ==========================

    print(f"\nDownloading {len(records)} files...\n")

    for r in records:

        month_folder = DOWNLOAD_DIR / r["date"].strftime("%Y-%m")
        month_folder.mkdir(exist_ok=True)

        outfile = month_folder / r["file"]

        if outfile.exists():
            print("Skipping:", r["file"])
            continue

        print("Downloading:", r["file"])

        resp = requests.get(r["url"], timeout=60)
        resp.raise_for_status()

        with open(outfile, "wb") as f:
            f.write(resp.content)

    print("\nDone.")
    # ==========================================
    # MAIN EXTRACTION
    # ==========================================

    master = []

    all_files = sorted(Path(INPUT_FOLDER).rglob("*.xlsx"))
    files = []

    for file in all_files:

        report_date = extract_report_date(file.name)

        if report_date is None:
            continue

        if start_date <= report_date <= end_date:
            files.append(file)

    print(f"Found {len(files)} files")

    for file in files:

        print(f"Processing: {file.name}")

        report_date = extract_report_date(file.name)

        try:

            xl = pd.ExcelFile(file)

            available_sheets = xl.sheet_names

            for sheet in TARGET_SHEETS:

                if sheet not in available_sheets:
                    continue

                raw = pd.read_excel(
                    file,
                    sheet_name=sheet,
                    header=None
                )

                header_row = None

                for idx in range(len(raw)):

                    row_text = " ".join(
                        raw.iloc[idx]
                        .fillna("")
                        .astype(str)
                    ).lower()

                    if "name of instruments" in row_text:

                        header_row = idx
                        break

                if header_row is None:
                    continue

                headers = (
                    raw.iloc[header_row]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .tolist()
                )

                data = raw.iloc[header_row+1:].copy()
                data.columns = headers

                first_col = data.columns[0]

                rows = []

                for _, row in data.iterrows():

                    val = str(row[first_col]).strip()

                    if val == "" or val.lower() == "nan":
                        continue

                    lower_val = val.lower()

                    # Stop once equity section ends
                    if any(
                        stop in lower_val
                        for stop in STOP_WORDS
                    ):
                        break

                    rows.append(row)

                if not rows:
                    continue

                equity_df = pd.DataFrame(rows)

                # -------------------------------------------------
                # Standardize column names
                # -------------------------------------------------

                rename_map = {}

                for col in equity_df.columns:

                    c = str(col).strip().lower()

                    if "instrument" in c:
                        rename_map[col] = "Particulars"

                    elif "isin" in c:
                        rename_map[col] = "ISIN"

                    elif "industry" in c:
                        rename_map[col] = "Industry"

                    elif "quantity" in c:
                        rename_map[col] = "Quantity"

                    elif "mkt" in c and "value" in c:
                        rename_map[col] = "Market Value"

                    elif "%" in c:
                        rename_map[col] = "% of Portfolio"

                equity_df = equity_df.rename(columns=rename_map)

                # -------------------------------------------------
                # Keep only required columns
                # -------------------------------------------------

                required = [
                    "Particulars",
                    "ISIN",
                    "Industry",
                    "Quantity",
                    "Market Value",
                    "% of Portfolio"
                ]

                for col in required:

                    if col not in equity_df.columns:
                        equity_df[col] = None

                equity_df = equity_df[required]

                # -------------------------------------------------
                # Filters
                # -------------------------------------------------

                equity_df["ISIN"] = (
                    equity_df["ISIN"]
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                equity_df = equity_df[
                    equity_df["ISIN"].str.startswith("INE", na=False)
                ]

                equity_df = equity_df[
                    ~equity_df["Particulars"]
                    .astype(str)
                    .str.upper()
                    .str.contains(
                        r"%|REIT|INVIT|BOND|DEBENTURE|NCD",
                        regex=True,
                        na=False
                    )
                ]

                # -------------------------------------------------
                # Add metadata
                # -------------------------------------------------

                equity_df.insert(
                    0,
                    "Date",
                    report_date.strftime("%b-%y")
                )

                equity_df.insert(
                    1,
                    "Fund",
                    "SBI"
                )

                equity_df.insert(
                    2,
                    "Scheme",
                    sheet
                )

                master.append(equity_df)
        except Exception as e:

            print(
                f"ERROR: {file.name}"
            )
            print(e)

    # ==========================================
    # FINAL OUTPUT
    # ==========================================
    if not master:
        print("No disclosures found.")
        return pd.DataFrame(
            columns=[
                "Date",
                "Fund",
                "Scheme",
                "Particulars",
                "ISIN",
                "Industry",
                "Quantity",
                "Market Value",
                "% of Portfolio"
            ]
        )

    final_df = pd.concat(master, ignore_index=True)
    final_df = final_df.drop_duplicates(
        subset=[
            "Date",
            "Scheme",
            "ISIN"
        ]
    )

    final_df = final_df[
        [
            "Date",
            "Fund",
            "Scheme",
            "Particulars",
            "ISIN",
            "Industry",
            "Quantity",
            "Market Value",
            "% of Portfolio"
        ]
    ]
    final_df = final_df.sort_values(
    [
        "Date",
        "Scheme",
        "Particulars"
    ]
    ).reset_index(drop=True)
    print(f"\nRows : {len(final_df):,}")
    for file in DOWNLOAD_DIR.rglob("*"):

        if file.is_file():
            try:
                file.unlink()
            except:
                pass

    for folder in sorted(
        DOWNLOAD_DIR.rglob("*"),
        reverse=True
    ):
        if folder.is_dir():
            try:
                folder.rmdir()
            except:
                pass

    return final_df

def run(start_date=None, end_date=None):
    return main(start_date, end_date)

if __name__ == "__main__":
    run()