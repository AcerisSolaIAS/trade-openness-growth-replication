#!/usr/bin/env python3
"""
Download World Bank WDI data for trade openness - economic growth paper.
Indicators:
  - GDP per capita constant 2015 USD: NY.GDP.PCAP.KD
  - Trade (% of GDP): NE.TRD.GNFS.ZS
  - Gross capital formation (% of GDP): NE.GDI.TOTL.ZS
  - School enrollment, primary (% gross): SE.PRM.ENRR
  - Population growth (annual %): SP.POP.GROW
"""

from pathlib import Path
import requests
import pandas as pd
import time

REPO_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = REPO_ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

INDICATORS = {
    "NY.GDP.PCAP.KD": "gdp_per_capita",
    "NE.TRD.GNFS.ZS": "trade_gdp_pct",
    "NE.GDI.TOTL.ZS": "investment_gdp_pct",
    "SE.PRM.ENRR": "school_primary",
    "SP.POP.GROW": "pop_growth",
}

START_YEAR = 2000
END_YEAR = 2022


def fetch_indicator(indicator, max_retries=5):
    url = f"https://api.worldbank.org/v2/country/all/indicator/{indicator}"
    params = {
        "format": "json",
        "date": f"{START_YEAR}:{END_YEAR}",
        "per_page": 20000,
    }
    all_data = []
    page = 1
    while True:
        params["page"] = page
        for attempt in range(max_retries):
            try:
                r = requests.get(url, params=params, timeout=180)
                r.raise_for_status()
                break
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
                print(f"  Timeout/connection error page {page}, attempt {attempt+1}/{max_retries}: {e}")
                if attempt == max_retries - 1:
                    raise
                time.sleep(2 ** attempt)
        js = r.json()
        if len(js) < 2:
            break
        data = js[1]
        if not data:
            break
        all_data.extend(data)
        if len(data) < params["per_page"]:
            break
        page += 1
        time.sleep(0.5)
    return all_data


def main():
    records = []
    for wb_code, var_name in INDICATORS.items():
        print(f"Downloading {wb_code} ({var_name})...")
        raw = fetch_indicator(wb_code)
        for item in raw:
            cc = item.get("countryiso3code")
            if not cc:
                continue
            records.append({
                "country_code": cc,
                "country": item.get("country", {}).get("value") if isinstance(item.get("country"), dict) else None,
                "year": int(item.get("date")) if item.get("date") else None,
                "variable": var_name,
                "value": item.get("value"),
            })

    df = pd.DataFrame(records)
    df = df.dropna(subset=["value", "year"])
    df.to_csv(DATA_DIR / "wdi_long.csv", index=False)

    df_wide = df.pivot_table(
        index=["country_code", "country", "year"],
        columns="variable",
        values="value",
    ).reset_index()
    df_wide.to_csv(DATA_DIR / "wdi_trade_growth.csv", index=False)

    print("Saved wide data:")
    print(df_wide.head())
    print(f"\nShape: {df_wide.shape}")
    print("\nMissing values:")
    print(df_wide.isna().sum())


