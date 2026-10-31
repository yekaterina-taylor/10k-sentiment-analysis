"""Retrieve NVIDIA (NVDA) Daily & Fiscal Year-End Stock Prices.

This script uses the yfinance library to download daily historical stock
prices for NVIDIA Corporation (ticker: NVDA) from 2017-01-01 to 2026-02-01.
It saves:
  1. Full daily trading data to 'NVDA_daily_prices_2017_2026.csv'.
  2. Fiscal year-end closing and adjusted closing prices for the 10
     balance sheet dates (FY2017–FY2026) to
     'NVDA_fiscal_year_end_prices.csv'.
"""

from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import yfinance as yf


# Fiscal year-end balance sheet dates specified for NVDA (FY2017–FY2026)
NVDA_BALANCE_SHEET_DATES: List[Tuple[str, str]] = [
    ("FY2017", "2017-01-29"),
    ("FY2018", "2018-01-28"),
    ("FY2019", "2019-01-27"),
    ("FY2020", "2020-01-26"),
    ("FY2021", "2021-01-31"),
    ("FY2022", "2022-01-30"),
    ("FY2023", "2023-01-29"),
    ("FY2024", "2024-01-28"),
    ("FY2025", "2025-01-26"),
    ("FY2026", "2026-01-25"),
]

DAILY_CSV_PATH: Path = Path("NVDA_daily_prices_2017_2026.csv")
FY_END_CSV_PATH: Path = Path("NVDA_fiscal_year_end_prices.csv")


def fetch_nvda_daily_data(
    ticker: str = "NVDA",
    start_date: str = "2017-01-01",
    end_date: str = "2026-02-01",
) -> pd.DataFrame:
    """Downloads daily price data for a ticker using yfinance.

    Args:
        ticker: Stock ticker symbol (default 'NVDA').
        start_date: Start date string formatted as YYYY-MM-DD.
        end_date: End date string formatted as YYYY-MM-DD.

    Returns:
        Pandas DataFrame containing daily stock prices and actions.

    Raises:
        RuntimeError: If yfinance returns an empty DataFrame.
    """
    print(
        f"[*] Downloading daily stock data for {ticker} from {start_date} "
        f"to {end_date}..."
    )
    df = yf.download(
        ticker,
        start=start_date,
        end=end_date,
        auto_adjust=False,
        actions=True,
    )
    if df.empty:
        raise RuntimeError(
            f"No price data returned by yfinance for ticker {ticker}."
        )

    print(f"[✓] Retrieved {len(df)} daily trading records.")
    return df


def save_daily_prices_csv(df: pd.DataFrame, output_path: Path) -> Path:
    """Saves full daily price data to CSV with clean column headers.

    Flattens any MultiIndex column tuples emitted by recent yfinance
    versions (e.g. ('Close', 'NVDA') -> 'Close') while preserving the Date
    index.

    Args:
        df: Daily price DataFrame from yfinance.
        output_path: Target Path for the CSV file.

    Returns:
        The resolved Path of the saved CSV.
    """
    export_df = df.copy()
    if isinstance(export_df.columns, pd.MultiIndex):
        # Flatten ('Close', 'NVDA') -> 'Close'
        export_df.columns = [col[0] for col in export_df.columns]

    export_df.to_csv(output_path, index=True)
    print(f"[✓] Saved daily stock prices to: {output_path.resolve()}")
    return output_path


def extract_fiscal_year_end_prices(
    df: pd.DataFrame,
    fiscal_dates: List[Tuple[str, str]],
) -> pd.DataFrame:
    """Extracts closing and adjusted closing prices for balance sheet dates.

    If a balance sheet date falls on a weekend or market holiday, the most
    recent trading day preceding that date is used.

    Args:
        df: Daily price DataFrame indexed by trading date.
        fiscal_dates: List of tuples (fiscal_year_label, date_string).

    Returns:
        DataFrame with columns: Fiscal Year, Balance Sheet Date,
        Actual Trading Date, Close, Adj Close.
    """
    if isinstance(df.columns, pd.MultiIndex):
        close_series = df["Close"]["NVDA"]
        adj_close_series = df["Adj Close"]["NVDA"]
    else:
        close_series = df["Close"]
        adj_close_series = df["Adj Close"]

    # Ensure index is datetime
    trading_dates = pd.to_datetime(close_series.index)

    rows: List[Dict[str, Optional[object]]] = []

    for fiscal_year, bs_date_str in fiscal_dates:
        bs_date = pd.to_datetime(bs_date_str)
        # Filter for all trading days on or before the balance sheet date
        valid_trading_days = trading_dates[trading_dates <= bs_date]

        if len(valid_trading_days) > 0:
            last_trading_day = valid_trading_days[-1]
            close_val = float(close_series.loc[last_trading_day])
            adj_close_val = float(adj_close_series.loc[last_trading_day])
            trading_date_str = last_trading_day.strftime("%Y-%m-%d")

            rows.append(
                {
                    "Fiscal Year": fiscal_year,
                    "Balance Sheet Date": bs_date_str,
                    "Actual Trading Date": trading_date_str,
                    "Close": close_val,
                    "Adj Close": adj_close_val,
                }
            )
        else:
            rows.append(
                {
                    "Fiscal Year": fiscal_year,
                    "Balance Sheet Date": bs_date_str,
                    "Actual Trading Date": "not found",
                    "Close": None,
                    "Adj Close": None,
                }
            )

    result_df = pd.DataFrame(rows)
    return result_df


def save_fiscal_year_end_csv(
    df: pd.DataFrame,
    output_path: Path,
) -> Path:
    """Saves fiscal year-end prices table to CSV.

    Args:
        df: DataFrame containing fiscal year-end price records.
        output_path: Target Path for the CSV file.

    Returns:
        The resolved Path of the saved CSV.
    """
    df.to_csv(output_path, index=False)
    print(f"[✓] Saved fiscal year-end prices to: {output_path.resolve()}")
    return output_path


def analyze_split_adjustment_status(df_fy: pd.DataFrame) -> None:
    """Analyzes and explains whether yfinance Close prices are split-adjusted.

    Checks FY2017 closing price:
      - If around $2-$3: Prices are retroactively split-adjusted for both
        the 4-for-1 split (July 2021) and 10-for-1 split (June 2024), a
        cumulative 40-to-1 split factor.
      - If around $110: Prices are unadjusted raw prices.

    Args:
        df_fy: DataFrame containing the 10 fiscal year-end prices.
    """
    row_2017 = df_fy[df_fy["Fiscal Year"] == "FY2017"]
    if row_2017.empty or row_2017["Close"].values[0] is None:
        print("[!] FY2017 record not available for split analysis.")
        return

    fy17_close = float(row_2017["Close"].values[0])
    fy17_adj = float(row_2017["Adj Close"].values[0])

    print("\n" + "=" * 76)
    print("           STOCK SPLIT ADJUSTMENT ANALYSIS (yfinance)")
    print("=" * 76)
    trading_dt = row_2017['Actual Trading Date'].values[0]
    print(f"  • FY2017 Actual Trading Date:  {trading_dt}")
    print(f"  • FY2017 Downloaded Close:     ${fy17_close:.4f}")
    print(f"  • FY2017 Downloaded Adj Close: ${fy17_adj:.4f}")

    if fy17_close < 10.0:
        unadj_approx = fy17_close * 40.0
        print("\n  [CONCLUSION]: 'Close' in yfinance IS SPLIT-ADJUSTED.")
        print(
            f"  Explanation: On January 27, 2017, NVDA originally closed at "
            f"~${unadj_approx:.2f}."
        )
        print(
            "  NVIDIA executed two major stock splits in this period:"
        )
        print("    1. 4-for-1 stock split on July 20, 2021")
        print("    2. 10-for-1 stock split on June 10, 2024")
        print(
            f"  Cumulative split multiplier: 4 x 10 = 40x. "
            f"Dividing ${unadj_approx:.2f} by 40 yields ~${fy17_close:.4f}."
        )
        print(
            "  Therefore, yfinance automatically applies retroactive split "
            "adjustments to 'Close'."
        )
        print(
            "  'Adj Close' additionally accounts for dividend reinvestment."
        )
    else:
        print("\n  [CONCLUSION]: 'Close' in yfinance IS NOT split-adjusted.")
    print("=" * 76)


def main() -> None:
    """Executes the daily and fiscal year-end stock price extraction."""
    # 1. Download full daily stock price series
    daily_df = fetch_nvda_daily_data(
        ticker="NVDA",
        start_date="2017-01-01",
        end_date="2026-02-01",
    )

    # 2. Save full daily data to CSV
    save_daily_prices_csv(daily_df, DAILY_CSV_PATH)

    # 3. Extract fiscal year-end prices
    fy_end_df = extract_fiscal_year_end_prices(
        daily_df,
        NVDA_BALANCE_SHEET_DATES,
    )

    # 4. Save fiscal year-end table to CSV
    save_fiscal_year_end_csv(fy_end_df, FY_END_CSV_PATH)

    # 5. Print the fiscal year-end prices table
    print("\n" + "=" * 76)
    print("       NVIDIA (NVDA) 10-YEAR FISCAL YEAR-END STOCK PRICES")
    print("=" * 76)
    print(
        f"{'Fiscal Year':<13} {'Balance Sheet Date':<20} "
        f"{'Actual Trading Date':<21} {'Close ($)':<12} {'Adj Close ($)':<12}"
    )
    print("-" * 76)
    for _, r in fy_end_df.iterrows():
        c_str = f"${r['Close']:.2f}" if pd.notnull(r["Close"]) else "not found"
        ac_str = (
            f"${r['Adj Close']:.2f}"
            if pd.notnull(r["Adj Close"])
            else "not found"
        )
        print(
            f"{r['Fiscal Year']:<13} {r['Balance Sheet Date']:<20} "
            f"{r['Actual Trading Date']:<21} {c_str:<12} {ac_str:<12}"
        )
    print("=" * 76)

    # 6. Analyze split adjustment status
    analyze_split_adjustment_status(fy_end_df)


if __name__ == "__main__":
    main()
