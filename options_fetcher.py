"""
options_fetcher.py

Download a recent closing price and a nearby options chain for one ticker
(SPY by default), print the contracts closest to that price, and save the
calls to a CSV you can open in Excel or Numbers.

Run it from this folder with the virtual environment active:

    python options_fetcher.py
"""

import warnings

# This note comes from a library yfinance uses on macOS. It is not an error
# in this script. The filter has to be set before that library is imported.
warnings.filterwarnings("ignore", message="urllib3 v2 only supports OpenSSL")

import pandas as pd
import yfinance as yf

# Change this string if you want a different symbol, such as "AAPL" or "QQQ".
TICKER = "AAPL"

# 0.05 means "5 percent." A strike is kept only if it sits between
# 95% and 105% of the last closing price.
STRIKE_BAND = 0.05

# These are the columns we print and write to the CSV file.
COLUMNS = [
    "contractSymbol",  # the exchange's short code for this exact contract
    "strike",          # the price the option is based on
    "lastPrice",       # the last traded price of the option itself
    "bid",             # the highest price a buyer is currently offering
    "ask",             # the lowest price a seller is currently asking
    "volume",          # how many contracts traded today (or the last session)
    "openInterest",    # how many contracts are still open
    "impliedVolatility",  # the market's expected movement, as a decimal
]


def near_the_money(contracts, last_close):
    """Keep strikes within 5% of the last close, then sort by volume."""
    # The lowest and highest strike we are willing to keep.
    low = last_close * (1 - STRIKE_BAND)
    high = last_close * (1 + STRIKE_BAND)

    # .copy() makes a new table so later sorting does not change the original.
    nearby = contracts[
        (contracts["strike"] >= low) & (contracts["strike"] <= high)
    ].copy()

    # Highest volume first. Contracts with a blank volume go to the bottom.
    nearby = nearby.sort_values("volume", ascending=False, na_position="last")

    # Volume and open interest are counts, so show them as whole numbers.
    # "Int64" still allows a blank when Yahoo did not report a value.
    for column in ("volume", "openInterest"):
        nearby[column] = nearby[column].astype("Int64")

    # Return only the columns we care about, in a friendly order.
    return nearby[COLUMNS]


def volume_above_open_interest(contracts):
    """Keep contracts where today's volume is higher than open interest.

    Open interest is how many contracts were already open. When volume is
    larger than that number, more contracts traded than were previously
    outstanding. Traders often read that as unusual buying.
    """
    # Drop rows Yahoo left blank. A blank volume or open interest cannot
    # be compared, so those contracts are left out of this list.
    comparable = contracts.dropna(subset=["volume", "openInterest"])

    unusual = comparable[comparable["volume"] > comparable["openInterest"]]

    # The table passed in is already sorted by volume. Keep that order.
    return unusual


def main():
    # Ticker() does not download anything yet. It just names the symbol
    # we will ask about in the next steps.
    stock = yf.Ticker(TICKER)

    # "5d" means the last five trading days. On a weekend the newest row
    # is still the most recent close, so this works while the market is shut.
    history = stock.history(period="5d")
    if history.empty:
        raise SystemExit(f"No price history came back for {TICKER}.")

    # iloc[-1] picks the last row. "Close" is the official end-of-day price.
    last_close = float(history["Close"].iloc[-1])
    last_date = history.index[-1].date()
    print(f"{TICKER} last close: ${last_close:.2f} on {last_date}")

    # .options is every expiration Yahoo has, soonest date first,
    # written like "2026-11-20".
    expirations = stock.options
    if not expirations:
        raise SystemExit(f"No option expirations came back for {TICKER}.")

    # Print only the first five dates. The list itself still has all of them.
    print("\nNext 5 expiration dates:")
    for expiration in expirations[:5]:
        print(f"  {expiration}")

    # The first date is the nearest expiration, so that is the chain we pull.
    nearest = expirations[0]
    print(f"\nUsing nearest expiration: {nearest}")

    # One download returns both sides of the chain:
    # calls (the right to buy) and puts (the right to sell).
    chain = stock.option_chain(nearest)
    calls = chain.calls
    puts = chain.puts

    filtered_calls = near_the_money(calls, last_close)
    filtered_puts = near_the_money(puts, last_close)

    # Tell pandas to print the whole table instead of hiding columns with "...".
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 200)
    pd.set_option("display.max_rows", None)

    low = last_close * (1 - STRIKE_BAND)
    high = last_close * (1 + STRIKE_BAND)
    print(
        f"\nStrikes between ${low:.2f} and ${high:.2f} "
        f"(within 5% of the last close)\n"
    )

    print("CALLS (highest volume first)")
    print(filtered_calls.to_string(index=False))

    print("\nPUTS (highest volume first)")
    print(filtered_puts.to_string(index=False))

    # Same nearby contracts, but only where volume beat open interest.
    unusual_calls = volume_above_open_interest(filtered_calls)
    unusual_puts = volume_above_open_interest(filtered_puts)

    print("\nUNUSUAL CALLS (volume greater than open interest)")
    if unusual_calls.empty:
        print("None")
    else:
        print(unusual_calls.to_string(index=False))

    print("\nUNUSUAL PUTS (volume greater than open interest)")
    if unusual_puts.empty:
        print("None")
    else:
        print(unusual_puts.to_string(index=False))

    # index=False leaves out pandas row numbers, so the file opens as a
    # normal spreadsheet with a header row.
    output_file = "spy_calls_weekend.csv"
    filtered_calls.to_csv(output_file, index=False)
    print(f"\nSaved {len(filtered_calls)} call contracts to {output_file}")


# This runs main() only when you start the file directly.
# Importing the file from another script will not download anything.
if __name__ == "__main__":
    main()
