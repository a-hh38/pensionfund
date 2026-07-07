import os
import re
import uuid
import requests
import pandas as pd
from pathlib import Path
from datetime import datetime

# ============================================================
# CONFIG
# ============================================================

DOWNLOAD_DIR = Path("lic_temp")
DOWNLOAD_DIR.mkdir(exist_ok=True)

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

EXCLUDED_SCHEMES = [
    "C-TIER I",
    "C-TIER II",
    "G-TIER I",
    "G-TIER II"
]

master_rows = []

seen_urls = set()

print("=" * 60)
print("LIC Portfolio Downloader")
print("=" * 60)

# ============================================================
# METADATA HELPERS
# ============================================================

def extract_metadata(filepath):
    """
    Extract scheme name and portfolio month from both
    2024 and 2025+ LIC formats.
    """

    try:

        raw = pd.read_excel(
            filepath,
            header=None,
            nrows=6
        )

        # Search top few rows for scheme/date text
        scheme_text = ""
        date_text = ""

        for value in raw.iloc[:, 0].dropna():

            text = str(value).strip()

            upper = text.upper()

            if (
                "SCHEME" in upper
                or "CENTRAL GOVT" in upper
                or "STATE GOVT" in upper
                or "CORPORATE" in upper
                or "UPS" in upper
                or "ATAL" in upper
                or "VATSALYA" in upper
                or "NPS LITE" in upper
                or "E - TIER I" in upper
                or "E - TIER II" in upper
            ):
                scheme_text = text

            if re.search(r"\d{2}-\d{2}-\d{4}", text):
                date_text = text

        year = None
        month = None

        m = re.search(
            r"(\d{2})-(\d{2})-(\d{4})",
            date_text
        )

        if m:
            month = int(m.group(2))
            year = int(m.group(3))

        return scheme_text.upper(), year, month

    except Exception:

        return "", None, None


def clean_scheme_name(text):

    text = str(text).upper()

    # Excluded schemes
    if "C-TIER I" in text:
        return "C_TIER_I"

    if "C-TIER II" in text:
        return "C_TIER_II"

    if "G-TIER I" in text:
        return "G_TIER_I"

    if "G-TIER II" in text:
        return "G_TIER_II"

    # Private Sector
    if "E - TIER II" in text or "E TIER II" in text:
        return "E_TIER_II"

    if "E - TIER I" in text or "E TIER I" in text:
        return "E_TIER_I"

    # Government
    if (
        "SCHEME : CG" in text
        or "SCHEME: CG" in text
        or "CENTRAL GOVT" in text
    ):
        return "CENTRAL_GOVT"

    if (
        "SCHEME : SG" in text
        or "SCHEME: SG" in text
        or "STATE GOVT" in text
    ):
        return "STATE_GOVT"

    # Other Schemes
    if "CORPORATE" in text:
        return "CORPORATE"

    if "UPS" in text:
        return "UPS"

    if "ATAL" in text:
        return "ATAL"

    if "VATSALYA" in text:
        return "VATSALYA"

    if "NPS LITE" in text:
        return "NPS_LITE"

    return "OTHER"

# ============================================================
# PORTFOLIO EXTRACTION
# ============================================================

def extract_portfolio(
    filepath,
    scheme,
    rpt_year,
    rpt_month
):

    try:

        raw = pd.read_excel(
            filepath,
            header=None
        )

        # ---------------------------------------------
        # Find header row automatically
        # ---------------------------------------------

        header_row = None

        for i in range(len(raw)):

            row = raw.iloc[i].astype(str).str.upper()

            text = " ".join(row)

            if (
                "PARTICULAR" in text
                and (
                    "ISIN" in text
                    or "QUANTITY" in text
                    or "MARKET VALUE" in text
                )
            ):
                header_row = i
                break

        if header_row is None:

            print("Header not found.")
            return

        # ---------------------------------------------
        # Read data below header
        # ---------------------------------------------
        headers = [
        "Particulars",
        "ISIN",
        "Industry",
        "Quantity",
        "Market Value",
        "% of Portfolio",
        "Ratings"
        ]

        df = pd.read_excel(
        filepath,
        skiprows=header_row + 1,
        header=None,
        usecols=range(7)
        )

        df.columns = headers

        # Ratings not required
        df = df.drop(columns=["Ratings"])        
        # Remove section headings
        df = df[
            ~df["Particulars"].astype(str).str.upper().isin([
                "GOVERNMENT SECURITIES AND RELATED INVESTMENTS",
                "EQUITY",
                "CORPORATE DEBT",
                "MONEY MARKET INSTRUMENTS",
                "REITS",
                "INVITS",
                "MUTUAL FUND UNITS"
            ])
        ]
        df = df[
        df["ISIN"].astype(str).str.match(
            r"^[A-Z]{2}[A-Z0-9]{10}$",
            na=False
        )
    ]
        # ---------------------------------------------
        # Clean rows
        # ---------------------------------------------

        df["Particulars"] = (
            df["Particulars"]
            .astype(str)
            .str.strip()
        )

        df = df[
            df["Particulars"] != ""
        ]

        df = df[
            df["Particulars"].str.upper() != "NAN"
        ]

        # Remove repeated headers
        df = df[
            ~df["Particulars"]
            .str.contains(
                "PARTICULAR",
                case=False,
                na=False
            )
        ]

        # Remove totals
        remove_words = [
            "TOTAL",
            "GRAND TOTAL",
            "SUB TOTAL",
            "NET CURRENT",
            "CURRENT ASSETS",
            "CURRENT LIABILITIES",
            "CASH",
            "MONEY MARKET",
            "MARGIN",
            "TREPS"
        ]

        pattern = "|".join(remove_words)

        df = df[
            ~df["Particulars"]
            .str.contains(
                pattern,
                case=False,
                na=False
            )
        ]

        # Remove rows without ISIN
        df = df[
            df["ISIN"]
            .notna()
        ]

        df["ISIN"] = (
            df["ISIN"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        df = df[df["ISIN"].str.match(r"^INE", na=False)]

        # Remove instruments whose name contains %
        df = df[
            ~df["Particulars"]
            .astype(str)
            .str.contains("%", regex=False, na=False)
        ]
        # Remove unwanted instrument names
        df = df[
            ~df["Particulars"]
            .astype(str)
            .str.upper()
            .str.contains(
                r"%|REIT|INVIT|BOND|DEBENTURE|NCD",
                regex=True,
                na=False
            )
        ]
        # ---------------------------------------------
        # Append to master
        # ---------------------------------------------

        month_name = datetime(
            rpt_year,
            rpt_month,
            1
        ).strftime("%b-%y")

        for _, row in df.iterrows():

            master_rows.append({

                "Date": month_name,
                "Fund": "LIC",
                "Scheme": scheme,

                "Particulars": row["Particulars"],
                "ISIN": row["ISIN"],
                "Industry": row["Industry"],
                "Quantity": row["Quantity"],
                "Market Value": row["Market Value"],
                "% of Portfolio": row["% of Portfolio"]

            })

        print(
            f"Extracted {len(df)} rows."
        )

    except Exception as e:

        print(e)

def main(start_date=None, end_date=None):

    global master_rows
    master_rows = []

    global seen_urls
    seen_urls = set()

    if start_date is None:
        start_date = datetime(2025, 4, 1)

    if end_date is None:
        end_date = datetime.today().replace(day=1)

    from dateutil.relativedelta import relativedelta

    current = start_date.replace(day=1)

    while current <= end_date:

        year = current.year
        month = current.month

        for page in range(1, 3):

            page_url = (
                "https://www.licpensionfund.in/"
                "public-disclosures/portfolio-disclosures"
                f"?year={year}&month={month}&page={page}"
            )

            print(f"\nChecking {year}-{month:02d} Page {page}")

            try:

                response = requests.get(
                    page_url,
                    headers=HEADERS,
                    timeout=130
                )

                if response.status_code != 200:
                    continue

                links = re.findall(
                    r'https://www\.licpensionfund\.in/resources/uploads/PageContentPdf/[^"\']+\.xlsx',
                    response.text,
                    flags=re.IGNORECASE
                )

                if not links:
                    break

                for url in links:

                    if url in seen_urls:
                        continue

                    seen_urls.add(url)

                    temp_file = DOWNLOAD_DIR / f"{uuid.uuid4().hex}.xlsx"

                    try:

                        file_response = requests.get(
                            url,
                            headers=HEADERS,
                            timeout=60
                        )

                        with open(temp_file, "wb") as f:
                            f.write(file_response.content)

                        scheme_text, rpt_year, rpt_month = extract_metadata(temp_file)

                        scheme = clean_scheme_name(scheme_text)

                        if scheme in [
                            "C_TIER_I",
                            "C_TIER_II",
                            "G_TIER_I",
                            "G_TIER_II"
                        ]:

                            temp_file.unlink(missing_ok=True)
                            print(f"Skipped : {scheme}")
                            continue

                        if rpt_year is None:
                            rpt_year = year

                        if rpt_month is None:
                            rpt_month = month

                        print(
                            f"Processing : "
                            f"{scheme} "
                            f"{rpt_month:02d}-{rpt_year}"
                        )

                        extract_portfolio(
                            temp_file,
                            scheme,
                            rpt_year,
                            rpt_month
                        )

                        temp_file.unlink(missing_ok=True)

                    except Exception as e:

                        print(f"Failed file : {url}")
                        print(e)

                        temp_file.unlink(missing_ok=True)

            except Exception as e:

                print(f"Error checking {page_url}")
                print(e)

        current += relativedelta(months=1)

    print("\nCreating master file...")

    master = pd.DataFrame(master_rows)

    if master.empty:

        print("No portfolio data extracted.")

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

    master = master.drop_duplicates(
        subset=[
            "Date",
            "Scheme",
            "ISIN"
        ]
    )

    master = master[
        master["ISIN"].notna()
    ]

    master = master[
        master["ISIN"]
        .astype(str)
        .str.len() >= 10
    ]

    master = master[
        ~master["Particulars"]
        .astype(str)
        .str.contains(
            "PARTICULAR|TOTAL|GRAND TOTAL|SUB TOTAL",
            case=False,
            na=False
        )
    ]

    master = master[
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

    master.reset_index(
        drop=True,
        inplace=True
    )

    for file in DOWNLOAD_DIR.glob("*.xlsx"):

        try:
            file.unlink()
        except:
            pass

    try:
        DOWNLOAD_DIR.rmdir()
    except:
        pass

    print(f"\nExtracted {len(master):,} holdings.")

    return master

def run(start_date=None, end_date=None):
    return main(start_date, end_date)

if __name__ == "__main__":
    run()