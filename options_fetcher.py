"""
options_fetcher.py

Download a recent closing price and a nearby options chain for one ticker
(SPY by default), build an HTML report, and save the filtered calls to a
CSV you can open in Excel or Numbers.

By default the report is meant for Cursor's Simple Browser (not Chrome).
Use --browser if you want your system browser instead.

Run it from this folder with the virtual environment active:

    python3 options_fetcher.py
    python3 options_fetcher.py --serve
    python3 options_fetcher.py --browser
"""

import argparse
import html
import time
import webbrowser
import warnings
from datetime import datetime
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import Thread

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

# Self-contained page written here, then viewed in Cursor or a browser.
HTML_REPORT = "options_report.html"

# Used with --serve so Cursor Simple Browser can load http://127.0.0.1:...
REPORT_PORT = 8765

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


def _format_display_value(column, value):
    """Turn raw API numbers into readable text for the HTML tables."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "—"
    if hasattr(value, "item"):  # numpy / pandas scalar
        value = value.item()
    if column in ("volume", "openInterest"):
        return f"{int(value):,}"
    if column == "impliedVolatility":
        return f"{float(value) * 100:.2f}%"
    if column in ("strike", "lastPrice", "bid", "ask"):
        return f"{float(value):.2f}"
    return html.escape(str(value))


def _dataframe_to_html_table(dataframe, table_id, empty_message):
    """Build one sortable-looking table block (search filters rows in the browser)."""
    if dataframe.empty:
        return (
            f'<p class="empty" id="{table_id}">{html.escape(empty_message)}</p>'
        )

    headers = "".join(
        f"<th>{html.escape(col)}</th>" for col in dataframe.columns
    )
    rows = []
    for _, row in dataframe.iterrows():
        cells = "".join(
            f"<td>{_format_display_value(col, row[col])}</td>"
            for col in dataframe.columns
        )
        rows.append(f"<tr>{cells}</tr>")
    body = "\n".join(rows)
    count = len(dataframe)

    return f"""
    <div class="table-block" id="{table_id}-block">
      <label class="search" for="{table_id}-search">
        Search this table
        <input type="search" id="{table_id}-search" class="table-search"
               placeholder="Strike, symbol, volume…"
               data-table="{table_id}">
      </label>
      <p class="row-count">{count} contract{'s' if count != 1 else ''}</p>
      <div class="table-scroll">
        <table class="data-table" id="{table_id}">
          <thead><tr>{headers}</tr></thead>
          <tbody>
{body}
          </tbody>
        </table>
      </div>
    </div>
    """


def build_html_report(
    ticker,
    last_close,
    last_date,
    expirations_preview,
    nearest_expiration,
    strike_low,
    strike_high,
    filtered_calls,
    filtered_puts,
    unusual_calls,
    unusual_puts,
    generated_at,
):
    """Assemble one standalone HTML file with CSS and light client-side search."""
    calls_html = _dataframe_to_html_table(
        filtered_calls, "calls-table", "No call contracts in this strike range."
    )
    puts_html = _dataframe_to_html_table(
        filtered_puts, "puts-table", "No put contracts in this strike range."
    )
    unusual_calls_html = _dataframe_to_html_table(
        unusual_calls,
        "unusual-calls-table",
        "No calls with volume above open interest.",
    )
    unusual_puts_html = _dataframe_to_html_table(
        unusual_puts,
        "unusual-puts-table",
        "No puts with volume above open interest.",
    )

    expiration_tags = "".join(
        f'<span class="chip">{html.escape(d)}</span>'
        for d in expirations_preview
    )

    # f-string body is large; keep report styling in one place for easy edits.
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(ticker)} options · {html.escape(nearest_expiration)}</title>
  <style>
    :root {{
      --bg: #0f1419;
      --surface: #1a2332;
      --surface2: #243044;
      --text: #e7ecf3;
      --muted: #94a3b8;
      --accent: #3b82f6;
      --accent-dim: #2563eb;
      --call: #22c55e;
      --put: #f97316;
      --unusual: #a855f7;
      --border: #334155;
      --radius: 12px;
      --shadow: 0 8px 32px rgba(0,0,0,.35);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "SF Pro Text", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.5;
    }}
  a {{ color: var(--accent); text-decoration: none; }}
  a:hover {{ text-decoration: underline; }}
    .top {{
      background: linear-gradient(135deg, #1e3a5f 0%, var(--surface) 55%, #0f172a 100%);
      border-bottom: 1px solid var(--border);
      padding: 1.5rem 1.25rem 2rem;
    }}
    .wrap {{ max-width: 1200px; margin: 0 auto; padding: 0 1.25rem 3rem; }}
    h1 {{ margin: 0 0 .25rem; font-size: 1.75rem; font-weight: 700; letter-spacing: -0.02em; }}
    .subtitle {{ color: var(--muted); margin: 0; font-size: .95rem; }}
    nav {{
      display: flex;
      flex-wrap: wrap;
      gap: .5rem;
      margin-top: 1.25rem;
      position: sticky;
      top: 0;
      z-index: 10;
      background: rgba(15, 20, 25, .92);
      backdrop-filter: blur(10px);
      padding: .75rem 0;
      border-bottom: 1px solid transparent;
    }}
    nav.is-stuck {{ border-bottom-color: var(--border); }}
    nav a {{
      display: inline-block;
      padding: .45rem .85rem;
      border-radius: 999px;
      background: var(--surface2);
      color: var(--text);
      font-size: .875rem;
      font-weight: 500;
      border: 1px solid var(--border);
    }}
    nav a:hover {{ background: var(--accent-dim); border-color: var(--accent); text-decoration: none; }}
    .cards {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 1rem;
      margin: 1.5rem 0;
    }}
    .card {{
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 1rem 1.15rem;
      box-shadow: var(--shadow);
    }}
    .card .label {{ color: var(--muted); font-size: .8rem; text-transform: uppercase; letter-spacing: .06em; }}
    .card .value {{ font-size: 1.5rem; font-weight: 700; margin-top: .25rem; }}
    .chips {{ display: flex; flex-wrap: wrap; gap: .4rem; margin-top: .5rem; }}
    .chip {{
      font-size: .8rem;
      padding: .2rem .55rem;
      border-radius: 6px;
      background: var(--surface2);
      border: 1px solid var(--border);
      color: var(--muted);
    }}
    section {{
      margin-top: 2.5rem;
      scroll-margin-top: 4.5rem;
    }}
    section h2 {{
      font-size: 1.25rem;
      margin: 0 0 .35rem;
      display: flex;
      align-items: center;
      gap: .5rem;
    }}
    section h2 .badge {{
      font-size: .7rem;
      font-weight: 600;
      padding: .15rem .5rem;
      border-radius: 6px;
      text-transform: uppercase;
    }}
    .badge-call {{ background: rgba(34, 197, 94, .2); color: var(--call); }}
    .badge-put {{ background: rgba(249, 115, 22, .2); color: var(--put); }}
    .badge-unusual {{ background: rgba(168, 85, 247, .2); color: var(--unusual); }}
    section .lead {{ color: var(--muted); margin: 0 0 1rem; font-size: .9rem; }}
    .table-block {{
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 1rem;
      box-shadow: var(--shadow);
    }}
    .search {{
      display: block;
      font-size: .85rem;
      color: var(--muted);
      margin-bottom: .5rem;
    }}
    .table-search {{
      display: block;
      width: 100%;
      max-width: 320px;
      margin-top: .35rem;
      padding: .5rem .75rem;
      border-radius: 8px;
      border: 1px solid var(--border);
      background: var(--bg);
      color: var(--text);
      font-size: .95rem;
    }}
    .row-count {{ margin: 0 0 .75rem; font-size: .8rem; color: var(--muted); }}
    .table-scroll {{ overflow-x: auto; border-radius: 8px; }}
    table.data-table {{
      width: 100%;
      border-collapse: collapse;
      font-size: .85rem;
    }}
    table.data-table th {{
      position: sticky;
      top: 0;
      background: var(--surface2);
      text-align: left;
      padding: .65rem .75rem;
      border-bottom: 2px solid var(--border);
      white-space: nowrap;
      font-weight: 600;
    }}
    table.data-table td {{
      padding: .55rem .75rem;
      border-bottom: 1px solid var(--border);
    }}
    table.data-table tbody tr:hover {{ background: rgba(59, 130, 246, .08); }}
    table.data-table td:first-child {{ font-family: ui-monospace, monospace; font-size: .78rem; }}
    .empty {{ color: var(--muted); font-style: italic; padding: 1rem; }}
    footer {{
      margin-top: 3rem;
      padding-top: 1rem;
      border-top: 1px solid var(--border);
      color: var(--muted);
      font-size: .8rem;
    }}
  </style>
</head>
<body>
  <header class="top">
    <div class="wrap">
      <h1>{html.escape(ticker)} options chain</h1>
      <p class="subtitle">Nearest expiration · generated {html.escape(generated_at)}</p>
      <nav id="main-nav">
        <a href="#overview">Overview</a>
        <a href="#calls">Calls</a>
        <a href="#puts">Puts</a>
        <a href="#unusual-calls">Unusual calls</a>
        <a href="#unusual-puts">Unusual puts</a>
      </nav>
    </div>
  </header>
  <main class="wrap">
    <section id="overview">
      <h2>Overview</h2>
      <p class="lead">Last closing price and strike filter used for every table below.</p>
      <div class="cards">
        <div class="card">
          <div class="label">Last close</div>
          <div class="value">${last_close:.2f}</div>
          <div class="label" style="margin-top:.5rem">{html.escape(str(last_date))}</div>
        </div>
        <div class="card">
          <div class="label">Expiration in use</div>
          <div class="value" style="font-size:1.15rem">{html.escape(nearest_expiration)}</div>
        </div>
        <div class="card">
          <div class="label">Strike band (±5%)</div>
          <div class="value" style="font-size:1.15rem">${strike_low:.2f} – ${strike_high:.2f}</div>
        </div>
        <div class="card">
          <div class="label">Next expirations</div>
          <div class="chips">{expiration_tags}</div>
        </div>
      </div>
    </section>
    <section id="calls">
      <h2><span class="badge badge-call">Calls</span> Near-the-money calls</h2>
      <p class="lead">Sorted by volume (highest first). Strikes within 5% of the last close.</p>
      {calls_html}
    </section>
    <section id="puts">
      <h2><span class="badge badge-put">Puts</span> Near-the-money puts</h2>
      <p class="lead">Same strike filter as calls.</p>
      {puts_html}
    </section>
    <section id="unusual-calls">
      <h2><span class="badge badge-unusual">Signal</span> Unusual calls</h2>
      <p class="lead">Volume greater than open interest — often watched for heavy new positioning.</p>
      {unusual_calls_html}
    </section>
    <section id="unusual-puts">
      <h2><span class="badge badge-unusual">Signal</span> Unusual puts</h2>
      <p class="lead">Same volume vs. open interest filter on puts.</p>
      {unusual_puts_html}
    </section>
    <footer>
      Data from Yahoo Finance via yfinance. For education only — not investment advice.
      Re-run <code>python options_fetcher.py</code> to refresh.
    </footer>
  </main>
  <script>
    document.querySelectorAll(".table-search").forEach((input) => {{
      const tableId = input.dataset.table;
      const table = document.getElementById(tableId);
      if (!table) return;
      input.addEventListener("input", () => {{
        const q = input.value.trim().toLowerCase();
        table.querySelectorAll("tbody tr").forEach((row) => {{
          const text = row.textContent.toLowerCase();
          row.style.display = !q || text.includes(q) ? "" : "none";
        }});
      }});
    }});
    const nav = document.getElementById("main-nav");
    window.addEventListener("scroll", () => {{
      nav.classList.toggle("is-stuck", window.scrollY > 120);
    }}, {{ passive: true }});
  </script>
</body>
</html>
"""


def save_report(html_content, filename):
    """Write the report HTML file next to this script."""
    path = Path(__file__).resolve().parent / filename
    path.write_text(html_content, encoding="utf-8")
    return path


def start_report_server(directory, port):
    """Serve the project folder on localhost (needed for --serve)."""
    handler = partial(SimpleHTTPRequestHandler, directory=str(directory))
    server = ThreadingHTTPServer(("127.0.0.1", port), handler)
    thread = Thread(target=server.serve_forever, daemon=True)
    thread.start()
    return server, f"http://127.0.0.1:{port}"


def print_cursor_view_instructions(report_url):
    """Tell the user how to open the report inside Cursor."""
    print("\nView in Cursor (not Chrome):")
    print("  1. Press Cmd+Shift+P (Command Palette)")
    print('  2. Run: Simple Browser: Show')
    print(f"  3. Paste this URL and press Enter:\n     {report_url}")
    print("\nTip: With --serve, keep this terminal window open while you read the report.")


def present_report(report_path, use_system_browser=False, keep_serving=False):
    """Show how to open the report in Cursor, or optionally in Chrome/Safari."""
    project_dir = report_path.parent
    report_name = report_path.name

    server = None
    if keep_serving:
        server, base = start_report_server(project_dir, REPORT_PORT)
        report_url = f"{base}/{report_name}"
    else:
        report_url = report_path.as_uri()

    print_cursor_view_instructions(report_url)

    if use_system_browser:
        webbrowser.open(report_url)
        print("\nAlso opened in your system default browser (--browser).")

    if keep_serving and server is not None:
        print(f"\nServing {report_name} at {report_url}")
        print("Press Ctrl+C to stop the server.")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            server.shutdown()
            print("\nServer stopped.")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Fetch nearby options and build an HTML report."
    )
    parser.add_argument(
        "--browser",
        action="store_true",
        help="Also open the report in your system browser (Chrome/Safari).",
    )
    parser.add_argument(
        "--serve",
        action="store_true",
        help="Run a local web server so Cursor Simple Browser can use http://127.0.0.1:8765/...",
    )
    return parser.parse_args()


def main(args=None):
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

    # .options is every expiration Yahoo has, soonest date first,
    # written like "2026-11-20".
    expirations = stock.options
    if not expirations:
        raise SystemExit(f"No option expirations came back for {TICKER}.")

    # The first five dates are shown on the report overview card.
    expirations_preview = list(expirations[:5])

    # The first date is the nearest expiration, so that is the chain we pull.
    nearest = expirations[0]

    # One download returns both sides of the chain:
    # calls (the right to buy) and puts (the right to sell).
    chain = stock.option_chain(nearest)
    calls = chain.calls
    puts = chain.puts

    filtered_calls = near_the_money(calls, last_close)
    filtered_puts = near_the_money(puts, last_close)

    # Same nearby contracts, but only where volume beat open interest.
    unusual_calls = volume_above_open_interest(filtered_calls)
    unusual_puts = volume_above_open_interest(filtered_puts)

    low = last_close * (1 - STRIKE_BAND)
    high = last_close * (1 + STRIKE_BAND)
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    report_html = build_html_report(
        TICKER,
        last_close,
        last_date,
        expirations_preview,
        nearest,
        low,
        high,
        filtered_calls,
        filtered_puts,
        unusual_calls,
        unusual_puts,
        generated_at,
    )
    report_path = save_report(report_html, HTML_REPORT)

    # index=False leaves out pandas row numbers, so the file opens as a
    # normal spreadsheet with a header row.
    output_file = "spy_calls_weekend.csv"
    filtered_calls.to_csv(output_file, index=False)

    print(f"\n{TICKER} last close: ${last_close:.2f} on {last_date}")
    print(f"Expiration: {nearest} · strikes ${low:.2f}–${high:.2f}")
    print(f"Report file: {report_path}")
    print(f"Saved {len(filtered_calls)} call rows to {output_file}")

    present_report(
        report_path,
        use_system_browser=args.browser,
        keep_serving=args.serve,
    )


# This runs main() only when you start the file directly.
# Importing the file from another script will not download anything.
if __name__ == "__main__":
    main(parse_args())
