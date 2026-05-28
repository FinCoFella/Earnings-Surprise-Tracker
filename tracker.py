import os
from datetime import datetime
import requests
from dotenv import load_dotenv
import openpyxl
import matplotlib.pyplot as plt

load_dotenv("Miscellaneous/.env")

API_KEY = os.getenv("FMP_API_KEY")
BASE_URL = "https://financialmodelingprep.com/stable/earnings"


def fetch_earnings(symbol):
    url = f"{BASE_URL}?symbol={symbol}&apikey={API_KEY}"
    response = requests.get(url)
    response.raise_for_status()
    return response.json()


def date_to_quarter(period_end_str):
    try:
        date = datetime.strptime(period_end_str, "%Y-%m-%d")
        quarter_map = {3: "1Q", 6: "2Q", 9: "3Q", 12: "4Q"}
        return f"{quarter_map[date.month]}{str(date.year)[-2:]}"
    except (ValueError, KeyError):
        return "N/A"


def date_to_period_end(date_str):
    try:
        date = datetime.strptime(date_str, "%Y-%m-%d")
        month, year = date.month, date.year
        if month <= 2:
            return f"{year - 1}-12-31"
        elif month <= 5:
            return f"{year}-03-31"
        elif month <= 8:
            return f"{year}-06-30"
        elif month <= 11:
            return f"{year}-09-30"
        else:
            return f"{year}-12-31"
    except ValueError:
        return "N/A"


def calculate_surprise(actual, estimated):
    if actual is None or estimated is None or estimated == 0:
        return None
    return (actual - estimated) / abs(estimated) * 100


def display_results(data, symbol):
    if not data:
        print(f"No earnings data found for {symbol}.")
        return

    print(f"\nEarnings Surprise Tracker — {symbol}")
    print("-" * 93)
    print(f"{'Announced':^12} {'Period End':^12} {'Quarter':^9} {'EPS Actual':^12} {'EPS Est':^10} {'Delta':^10} {'Surprise %':^12} {'Flag':^6}")
    print("-" * 93)

    for record in data:
        actual = record.get("epsActual")
        estimated = record.get("epsEstimated")

        if actual is None or estimated is None:
            continue

        date = record.get("date", "N/A")
        period_end = date_to_period_end(date)
        quarter = date_to_quarter(period_end)
        surprise = calculate_surprise(actual, estimated)
        delta = actual - estimated

        actual_str = f"{actual:.2f}"
        estimated_str = f"{estimated:.2f}"
        delta_str = f"{delta:+.2f}"

        surprise_str = f"{surprise:+.1f}%"
        flag = "BEAT" if surprise > 10 else "MISS" if surprise < -10 else ""

        print(f"{date:^12} {period_end:^12} {quarter:^9} {actual_str:^12} {estimated_str:^10} {delta_str:^10} {surprise_str:^12} {flag:^6}")

    print("-" * 93)


def export_to_excel(data, symbol):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = symbol

    headers = ["Announced", "Period End", "Quarter", "EPS Actual", "EPS Est", "Delta", "Surprise %", "Flag"]
    ws.append(headers)

    for record in data:
        actual = record.get("epsActual")
        estimated = record.get("epsEstimated")

        if actual is None or estimated is None:
            continue

        date = record.get("date", "N/A")
        period_end = date_to_period_end(date)
        quarter = date_to_quarter(period_end)
        surprise = calculate_surprise(actual, estimated)
        delta = actual - estimated
        flag = "BEAT" if surprise > 10 else "MISS" if surprise < -10 else ""

        ws.append([date, period_end, quarter, actual, estimated, round(delta, 2), round(surprise, 1), flag])

    os.makedirs("Excel", exist_ok=True)
    filename = os.path.join("Excel", f"{symbol}_Earnings.xlsx")
    wb.save(filename)
    print(f"Exported to {filename}")


def export_to_chart(data, symbol):
    quarters = []
    surprises = []

    chart_start = datetime(2014, 12, 31)

    for record in reversed(data):
        actual = record.get("epsActual")
        estimated = record.get("epsEstimated")
        if actual is None or estimated is None:
            continue
        period_end = date_to_period_end(record.get("date", ""))
        try:
            if datetime.strptime(period_end, "%Y-%m-%d") < chart_start:
                continue
        except ValueError:
            continue
        quarters.append(date_to_quarter(period_end))
        surprises.append(calculate_surprise(actual, estimated))

    colors = ["green" if s >= 0 else "red" for s in surprises]
    x = list(range(len(quarters)))

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.bar(x, surprises, color=colors)
    ax.axhline(0, color="black", linewidth=0.8)
    ax.set_title(f"{symbol} — Earnings Surprise %")
    ax.set_xlabel("Quarter")
    ax.set_ylabel("Surprise %")
    ax.set_xticks(x)
    ax.set_xticklabels(
        [q if q.startswith("4Q") else "" for q in quarters],
        rotation=45,
        ha="right"
    )
    plt.tight_layout()

    os.makedirs("Excel", exist_ok=True)
    filename = os.path.join("Excel", f"{symbol}_Surprise.png")
    plt.savefig(filename, dpi=150)
    plt.close()
    print(f"Chart saved to {filename}")


def main():
    symbol = input("Enter ticker symbol: ").strip().upper()
    print(f"Fetching earnings data for {symbol}...")

    try:
        data = fetch_earnings(symbol)
        display_results(data, symbol)
        export_to_excel(data, symbol)
        export_to_chart(data, symbol)
    except requests.exceptions.HTTPError as e:
        print(f"API error: {e}")
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()
