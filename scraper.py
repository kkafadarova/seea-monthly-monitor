from __future__ import annotations

import hashlib
import io
import json
import re
import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://portal.seea.government.bg"
PAGE_URL = f"{BASE_URL}/bg/ByProducerAndEnergyObject"
SEARCH_URL = f"{PAGE_URL}/Search"
EXPORT_URL = f"{PAGE_URL}/ExportKendoTable"

DATA_DIR = Path("data")
CSV_PATH = DATA_DIR / "energy_objects.csv"
XLSX_PATH = DATA_DIR / "energy_objects.xlsx"
NEW_RECORDS_PATH = DATA_DIR / "new_records.xlsx"
STATUS_PATH = DATA_DIR / "last_run.json"
TIMEZONE = ZoneInfo("Europe/Sofia")


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def flatten_columns(columns: pd.Index) -> list[str]:
    result: list[str] = []
    used: dict[str, int] = {}
    for column in columns:
        if isinstance(column, tuple):
            parts = [clean_text(part) for part in column if clean_text(part) and not clean_text(part).startswith("Unnamed:")]
            name = " | ".join(dict.fromkeys(parts))
        else:
            name = clean_text(column)
        name = name or "column"
        count = used.get(name, 0)
        used[name] = count + 1
        if count:
            name = f"{name}_{count + 1}"
        result.append(name)
    return result


def find_search_query_id(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    for selector in (
        'input[name="searchQueryId"]',
        'input[id="searchQueryId"]',
        'input[name="SearchQueryId"]',
        'input[id="SearchQueryId"]',
    ):
        element = soup.select_one(selector)
        if element and element.get("value"):
            return clean_text(element["value"])

    for pattern in (
        r'["\']searchQueryId["\']\s*[:=]\s*["\']([^"\']+)["\']',
        r'searchQueryId=([0-9a-fA-F-]{20,})',
        r'value=["\']([0-9a-fA-F-]{36})["\']',
    ):
        match = re.search(pattern, html, flags=re.IGNORECASE)
        if match:
            return match.group(1)
    raise RuntimeError("Не беше намерен searchQueryId в началната страница.")


def read_exported_xls(content: bytes) -> pd.DataFrame:
    try:
        tables = pd.read_html(io.BytesIO(content))
        if tables:
            return max(tables, key=lambda frame: len(frame.index))
    except Exception:
        pass

    try:
        return pd.read_excel(io.BytesIO(content), engine="xlrd")
    except Exception as exc:
        raise RuntimeError("Export файлът не може да бъде прочетен като HTML таблица или .xls.") from exc


def fetch_all_records() -> pd.DataFrame:
    today = datetime.now(TIMEZONE).strftime("%d.%m.%Y")
    common_headers = {
        "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0 Safari/537.36",
        "Accept-Language": "bg-BG,bg;q=0.9,en;q=0.8",
        "Referer": PAGE_URL,
    }

    with requests.Session() as session:
        page_response = session.get(PAGE_URL, headers=common_headers, timeout=60)
        page_response.raise_for_status()
        search_query_id = find_search_query_id(page_response.text)

        search_params = {
            "searchQueryId": search_query_id,
            "OwnerName": "",
            "OwnerBulstat": "",
            "ObjectName": "",
            "ProvinceSourceDropDownChoice": "",
            "MunicipalitySourceDropDownChoice": "",
            "SettlementSourceDropDownChoice": "",
            "Ekatte": "",
            "CapacityFrom": "",
            "CapacityTo": "",
            "From": "",
            "To": today,
            "SourceTypeDropDownChoice": "",
        }

        search_headers = {**common_headers, "Accept": "text/html, */*; q=0.01", "X-Requested-With": "XMLHttpRequest"}
        search_response = session.get(SEARCH_URL, params=search_params, headers=search_headers, timeout=120)
        search_response.raise_for_status()

        export_headers = {**common_headers, "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8"}
        export_response = session.get(EXPORT_URL, params={"searchQueryId": search_query_id}, headers=export_headers, timeout=180)
        export_response.raise_for_status()
        if len(export_response.content) < 1000:
            raise RuntimeError(f"Export файлът е необичайно малък: {len(export_response.content)} bytes.")

    table = read_exported_xls(export_response.content).copy()
    table.columns = flatten_columns(table.columns)
    table = table.dropna(how="all")
    for column in table.columns:
        table[column] = table[column].map(clean_text)
    table = table[table.apply(lambda row: any(clean_text(value) for value in row), axis=1)]
    return table.reset_index(drop=True)


def make_record_id(row: pd.Series, business_columns: list[str]) -> str:
    canonical = "\x1f".join(clean_text(row.get(column, "")) for column in business_columns)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def save_results(current: pd.DataFrame) -> tuple[int, int]:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    now = datetime.now(TIMEZONE).isoformat(timespec="seconds")
    metadata_columns = {"record_id", "first_seen_at", "last_seen_at"}
    business_columns = [column for column in current.columns if column not in metadata_columns]

    current["record_id"] = current.apply(lambda row: make_record_id(row, business_columns), axis=1)
    current = current.drop_duplicates(subset=["record_id"]).copy()

    if CSV_PATH.exists():
        existing = pd.read_csv(CSV_PATH, dtype=str, keep_default_na=False)
        all_business_columns = list(dict.fromkeys([
            *[c for c in existing.columns if c not in metadata_columns],
            *business_columns,
        ]))
        for column in all_business_columns:
            if column not in existing.columns:
                existing[column] = ""
            if column not in current.columns:
                current[column] = ""

        existing_ids = set(existing["record_id"].astype(str))
        new_rows = current[~current["record_id"].isin(existing_ids)].copy()
        new_rows["first_seen_at"] = now
        new_rows["last_seen_at"] = now

        current_ids = set(current["record_id"])
        if "last_seen_at" not in existing.columns:
            existing["last_seen_at"] = ""
        existing.loc[existing["record_id"].isin(current_ids), "last_seen_at"] = now

        ordered_columns = ["record_id", *all_business_columns, "first_seen_at", "last_seen_at"]
        combined = pd.concat([
            existing.reindex(columns=ordered_columns, fill_value=""),
            new_rows.reindex(columns=ordered_columns, fill_value=""),
        ], ignore_index=True)
    else:
        new_rows = current.copy()
        new_rows["first_seen_at"] = now
        new_rows["last_seen_at"] = now
        ordered_columns = ["record_id", *business_columns, "first_seen_at", "last_seen_at"]
        combined = new_rows.reindex(columns=ordered_columns, fill_value="")

    combined.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")
    combined.to_excel(XLSX_PATH, index=False)
    new_rows.to_excel(NEW_RECORDS_PATH, index=False)

    status = {
        "run_at": now,
        "records_received_from_portal": int(len(current)),
        "new_records_added": int(len(new_rows)),
        "total_unique_records_saved": int(len(combined)),
    }
    STATUS_PATH.write_text(json.dumps(status, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(status, ensure_ascii=False, indent=2))
    return len(new_rows), len(combined)


def main() -> int:
    try:
        current = fetch_all_records()
        if current.empty:
            raise RuntimeError("Порталът върна празен Excel файл.")
        save_results(current)
        return 0
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
