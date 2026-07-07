import os
import requests
import pandas as pd
from datetime import datetime
from openpyxl import Workbook, load_workbook
import re


# ==========================
# CONFIG
# ==========================

BASE_URL = "https://www.dsppension.com/uploads/upload_disclosure"
DOWNLOAD_FOLDER = "downloads_dsp"
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
MONTHS = [
    "January", "February", "March", "April",
    "May", "June", "July", "August",
    "September", "October", "November", "December"
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 Chrome/138.0 Safari/537.36"
    )
}


# ==========================
# DOWNLOAD LAST 13 MONTHS
# ==========================
def download_reports(start_date, end_date):

    downloaded_files = []

    current = start_date.replace(day=1)

    while current <= end_date:

        year = current.year
        month = current.month

        full_month = MONTHS[month - 1]
        short_month = current.strftime("%b")

        possible_files = [
            f"Monthly_Website_Portfolio_Form_{full_month}_{year}.xlsx",
            f"Monthly_Website_Portfolio_Form_{short_month}_{year}.xlsx",
        ]

        downloaded = False

        for filename in possible_files:

            url = f"{BASE_URL}/{filename}"

            print(f"Checking: {url}")

            try:

                r = requests.get(
                    url,
                    headers=HEADERS,
                    timeout=30
                )

                if (
                    r.status_code == 200
                    and r.content[:2] == b"PK"
                ):

                    save_path = os.path.join(
                        DOWNLOAD_FOLDER,
                        filename
                    )

                    with open(save_path, "wb") as f:
                        f.write(r.content)

                    print(f"Downloaded -> {filename}")

                    downloaded_files.append(save_path)

                    downloaded = True
                    break

            except Exception as e:

                print(e)

        if not downloaded:
            print(f"No report found for {full_month} {year}")

        if month == 12:
            current = current.replace(
                year=year + 1,
                month=1
            )
        else:
            current = current.replace(
                month=month + 1
            )

    return downloaded_files

# ==========================
# CREATE OUTPUT WORKBOOK
# ==========================
def extract_portfolio(file):

    all_dfs = []

    xls = pd.ExcelFile(file)

    REMOVE_ROWS = {
        "Equity Instruments",
        "Shares",
        "Subtotal",
        "Total",
        "GRAND TOTAL",
        "Bank FD",
        "Equity",
        "Equity Mutual Funds",
        "Equity Exchange Traded Funds",
        "Equity Oriented Mutual Fund Schemes",
        "Preference Shares",
        "Debt Instruments",
        "Bonds / NCD",
        "Private Corporate Bonds",
        "Government Securities",
        "Central Government Securities",
        "State Development Loans",
        "Treasury Bills",
        "Money Market Instruments",
        "Money Market Mutual Funds",
        "Money Market",
        "Overnight Funds",
        "Cash/Cash equivalent & Net Current Assets",
        "Alternate Investment Funds",
        "Real Estate Investment Trust Units",
        "Real Estate Investment Trusts",
        "Infrastructure Investment Trust Units",
        "Infrastructure Investment Trusts",
        "Mutual Funds",
        "NCDs",
        "Others",
        "Receivables",
        "Payables",
        "NAV Date",
        "Unit Outstanding",
        "Net Assets value",
        "Net Assets Value",
        "Credit Rating Exposure",
        "Modified Duration",
        "Average Maturity of Portfolio (in yrs)",
        "Yield to Maturity (%) (annualised) (at market price)",
        "Application Pending Allotment",
        "Net NPA",
    }

    BAD_PREFIXES = (
        "DSP Liquidity Fund",
        "Union Overnight Fund",
        "Union Liquid Fund",
        "HSBC Liquid Fund",
        "Mirae Asset Liquid Fund",
        "Nippon India Liquid Fund",
        "ICICI Prudential",
        "DSP Nifty",
        "Brookfield India",
        "Embassy Office Parks",
    )

    for sheet in xls.sheet_names:

        try:

            meta = pd.read_excel(xls, sheet_name=sheet, header=None)

            scheme_name = ""

            report_date = None

            for i in range(len(meta)):

                key = str(meta.iloc[i, 0]).strip()

                if key == "Name of the scheme":
                    scheme_name = str(meta.iloc[i, 1]).strip()

                elif "Portfolio Statement as on" in key:
                    report_date = pd.to_datetime(meta.iloc[i, 1]).strftime("%b-%y")

            if "PRIVATE LIMITED" in scheme_name:
                scheme_name = scheme_name.split("PRIVATE LIMITED", 1)[1].strip()

            scheme_name = scheme_name.lstrip("-: ").strip()

            print(f"\nSheet : {sheet}")
            print(f"Scheme: {scheme_name}")
            print(f"Date  : {report_date}")

            # Skip Scheme A, C and G
            scheme_upper = scheme_name.upper()

            if (
                "SCHEME A" in scheme_upper or
                "SCHEME C" in scheme_upper or
                "SCHEME G" in scheme_upper
            ):
                print(f"Skipping {scheme_name}")
                continue

            df = pd.read_excel(xls, sheet_name=sheet, header=5)

            df = df.dropna(how="all")

            if "Particulars" not in df.columns:
                continue

            df["Particulars"] = (
                df["Particulars"]
                .astype(str)
                .str.strip()
                .str.replace(r"\s+", " ", regex=True)
            )

                        # Remove blank rows
            df = df[
                (df["Particulars"].notna()) &
                (df["Particulars"] != "")
            ]

            # Remove known section headers
            df = df[
                ~df["Particulars"].isin(REMOVE_ROWS)
            ]

            # Remove known fund names / REITs
            df = df[
                ~df["Particulars"].str.startswith(
                    BAD_PREFIXES,
                    na=False
                )
            ]

            # Remove dates
            df = df[
                ~df["Particulars"].str.fullmatch(
                    r"\d{4}-\d{2}-\d{2}(?:\s00:00:00)?",
                    na=False
                )
            ]

            # Remove numbered notes
            df = df[
                ~df["Particulars"].str.match(
                    r"^\d+\.",
                    na=False
                )
            ]

            # Remove coupon/bond rows
            df = df[
                ~df["Particulars"].str.match(
                    r"^\d+(\.\d+)?%",
                    na=False
                )
            ]

            # Remove unwanted phrases
            BAD_CONTAINS = [
                "Money Market",
                "Liquid Fund",
                "Growth Option",
                "Direct Plan",
                "Direct - Growth",
                "Fixed Deposit",
                "Government Securities",
                "Infrastructure Investment",
                "Average Maturity",
                "Modified Duration",
                "Yield to Maturity",
                "Credit Rating Exposure",
                "Application Pending",
                "Unit Outstanding",
                "Net Assets",
                "Net asset values",
                "NAV Date",
                "Net NPA",
                "Out of above",
                "provision made",
                "Cash/Cash equivalent",
                "Treasury Bills",
                "State Development Loans",
                "Receivables",
                "Payables",
            ]

            pattern = "|".join(map(re.escape, BAD_CONTAINS))

            df = df[
                ~df["Particulars"].str.contains(
                    pattern,
                    case=False,
                    na=False
                )
            ]

            # Remove scheme names accidentally present
            df = df[
                ~df["Particulars"].str.contains(
                    r"SCHEME\s+[A-Z]",
                    case=False,
                    na=False
                )
            ]

            # Remove rows without a company name
            df = df[
                df["Particulars"].str.len() > 3
            ]

            # Drop unwanted columns
            DROP_COLS = [
                "Total Market Value",
                "Ratings",
                "Rating (if rated)",
            ]

            df = df.drop(
                columns=[c for c in DROP_COLS if c in df.columns],
                errors="ignore"
            )

            # Add metadata
            df.insert(0, "Date", report_date)
            df.insert(1, "Fund", "DSP")
            df.insert(2, "Scheme", scheme_name)

            df.reset_index(drop=True, inplace=True)

            print(f"Equity rows: {len(df)}")

            all_dfs.append(df)

        except Exception as e:
            print(f"Skipping sheet '{sheet}' -> {e}")

    xls.close()

    if all_dfs:

        final_df = pd.concat(all_dfs, ignore_index=True)

        BAD_PARTICULARS = [
            "Portfolio",
            "Scheme",
            "DSP NPS LONG TERM EQUITY FUND - TIER I",
            "DSP NPS LONG TERM EQUITY FUND",
            "NPS VATSALYA SCHEME",
            "NPS VATSALYA SCHEME - DIRECT",
            "NPS VATSALYA SCHEME - POP",
        ]

        final_df = final_df[
            ~final_df["Particulars"].isin(BAD_PARTICULARS)
        ]

        isin_col = next(
            (c for c in final_df.columns if "isin" in str(c).lower()),
            None
        )

        if isin_col:

            final_df[isin_col] = (
                final_df[isin_col]
                .astype(str)
                .str.strip()
            )

            final_df = final_df[
                final_df[isin_col].str.match(
                    r"^IN[A-Z0-9]{10}$",
                    na=False
                )
            ]
        # Standardize column names
        final_df.rename(
            columns={
                "ISIN No.": "ISIN",
                "ISIN NO.": "ISIN",
                "Isin No.": "ISIN",
                "Isin": "ISIN",
            },
            inplace=True
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

        return final_df

    return pd.DataFrame()

# ==========================
# MAIN
# ==========================
def main(start_date=None, end_date=None):

    if start_date is None:
        start_date = pd.Timestamp(2025, 4, 1)

    if end_date is None:
        end_date = pd.Timestamp.today().replace(day=1)

    print("=" * 60)
    print("DSP Pension Portfolio Downloader")
    print("=" * 60)

    # Clean download folder
    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

    for file in os.listdir(DOWNLOAD_FOLDER):

        path = os.path.join(
            DOWNLOAD_FOLDER,
            file
        )

        if os.path.isfile(path):

            try:
                os.remove(path)
            except:
                pass

    files = download_reports(
        start_date,
        end_date
    )

    if not files:

        print("\nNo files downloaded.")

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

    all_frames = []

    for file in sorted(files):

        try:

            df = extract_portfolio(file)

            if not df.empty:
                all_frames.append(df)

        except Exception as e:

            print(f"Error processing {os.path.basename(file)}")
            print(e)

        finally:

            try:
                os.remove(file)
            except:
                pass

    if not all_frames:

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

    final_df = pd.concat(
        all_frames,
        ignore_index=True
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

    final_df = final_df.drop_duplicates(
        subset=[
            "Date",
            "Scheme",
            "ISIN"
        ]
    )

    final_df = final_df.sort_values(
        [
            "Date",
            "Scheme",
            "Particulars"
        ]
    ).reset_index(drop=True)

    print("\n" + "=" * 60)
    print(f"Extracted {len(final_df):,} rows.")
    print("=" * 60)

    return final_df

def run(start_date=None, end_date=None):
    return main(start_date, end_date)

if __name__ == "__main__":
    run()