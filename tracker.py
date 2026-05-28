import os
from datetime import datetime
import requests
from dotenv import load_dotenv
import openpyxl

load_dotenv()

API_KEY = os.getenv("FMP_API_KEY")
BASE_URL = "https://financialmodelingprep.com/stable/earnings"


def fetch_earnings(symbol):
    url = f"{BASE_URL}?symbol={symbol}&apikey={API_KEY}"
    response = requests.get(url)
    response.raise_for_status()
    return response.json()


def date_to_quarter(date_str):
    try:
        date = datetime.strptime(date_str, "%Y-%m-%d")
        month, year = date.month, date.year
        if month <= 2:
            return f"4Q{str(year - 1)[-2:]}"
        elif month <= 5:
            return f"1Q{str(year)[-2:]}"
        elif month <= 8:
            return f"2Q{str(year)[-2:]}"
        elif month <= 11:
            return f"3Q{str(year)[-2:]}"
        else:
            return f"4Q{str(year)[-2:]}"
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
    print("-" * 80)
    print(f"{'Date':^12} {'Quarter':^9} {'EPS Actual':^12} {'EPS Est':^10} {'Delta':^10} {'Surprise %':^12} {'Flag':^6}")
    print("-" * 80)

    for record in data:
        actual = record.get("epsActual")
        estimated = record.get("epsEstimated")

        if actual is None or estimated is None:
            continue

        date = record.get("date", "N/A")
        quarter = date_to_quarter(date)
        surprise = calculate_surprise(actual, estimated)
        delta = actual - estimated

        actual_str = f"{actual:.2f}"
        estimated_str = f"{estimated:.2f}"
        delta_str = f"{delta:+.2f}"

        surprise_str = f"{surprise:+.1f}%"
        flag = "BEAT" if surprise > 10 else "MISS" if surprise < -10 else ""

        print(f"{date:^12} {quarter:^9} {actual_str:^12} {estimated_str:^10} {delta_str:^10} {surprise_str:^12} {flag:^6}")

    print("-" * 80)


def export_to_excel(data, symbol):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = symbol

    headers = ["Date", "Quarter", "EPS Actual", "EPS Est", "Delta", "Surprise %", "Flag"]
    ws.append(headers)

    for record in data:
        actual = record.get("epsActual")
        estimated = record.get("epsEstimated")

        if actual is None or estimated is None:
            continue

        date = record.get("date", "N/A")
        quarter = date_to_quarter(date)
        surprise = calculate_surprise(actual, estimated)
        delta = actual - estimated
        flag = "BEAT" if surprise > 10 else "MISS" if surprise < -10 else ""

        ws.append([date, quarter, actual, estimated, round(delta, 2), round(surprise, 1), flag])

    os.makedirs("Excel", exist_ok=True)
    filename = os.path.join("Excel", f"{symbol}_Earnings.xlsx")
    wb.save(filename)
    print(f"Exported to {filename}")


def main():
    symbol = input("Enter ticker symbol: ").strip().upper()
    print(f"Fetching earnings data for {symbol}...")

    try:
        data = fetch_earnings(symbol)
        display_results(data, symbol)
        export_to_excel(data, symbol)
    except requests.exceptions.HTTPError as e:
        print(f"API error: {e}")
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()
