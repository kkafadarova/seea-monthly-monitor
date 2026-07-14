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

DATA_DIR = Path("data")
CSV_PATH = DATA_DIR / "energy_objects.csv"
XLSX_PATH = DATA_DIR / "energy_objects.xlsx"
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
            parts = [
                clean_text(part)
                for part in column
                if clean_text(part) and not clean_text(part).startswith("Unnamed:")
            ]
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

    selectors = [
        'input[name="searchQueryId"]',
        'input[id="searchQueryId"]',
        'input[name="SearchQueryId"]',
        'input[id="SearchQueryId"]',
    ]

    for selector in selectors:
        element = soup.select_one(selector)
        if element and element.get("value"):
            return clean_text(element["value"])

    patterns = [
        r'["\']searchQueryId["\']\s*[:=]\s*["\']([^"\']+)["\']',
        r'searchQueryId=([0-9a-fA-F-]{20,})',
        r'value=["\']([0-9a-fA-F-]{36})["\']',
    ]

    for pattern in patterns:
        match = re.search(pattern, html, flags=re.IGNORECASE)
        if match:
            return match.group(1)

    raise RuntimeError(
        "Не беше намерен searchQueryId в началната страница. "
        "Вероятно порталът е променил HTML структурата."
    )


def fetch_table() -> pd.DataFrame:
    today = datetime.now(TIMEZONE).strftime("%d.%m.%Y")

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/149.0 Safari/537.36"
        ),
        "Accept": "text/html, */*; q=0.01",
        "Accept-Language": "bg-BG,bg;q=0.9,en;q=0.8",
        "Referer": PAGE_URL,
        "X-Requested-With": "XMLHttpRequest",
    }

    with requests.Session() as session:
        page_response = session.get(PAGE_URL, headers=headers, timeout=60)
        page_response.raise_for_status()

        search_query_id = find_search_query_id(page_response.text)

        params = {
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

        response = session.get(
            SEARCH_URL,
            params=params,
            headers=headers,
            timeout=120,
        )
        response.raise_for_status()

    tables = pd.read_html(io.StringIO(response.text))
    if not tables:
        raise RuntimeError("Search заявката не върна HTML таблица.")

    # Избираме таблицата с най-много редове — обикновено това е основният резултат.
    table = max(tables, key=lambda frame: len(frame.index)).copy()
    table.columns = flatten_columns(table.columns)

    table = table.dropna(how="all")
    for column in table.columns:
        table[column] = table[column].map(clean_text)

    # Премахва евентуално повторен header, попаднал като ред.
    if not table.empty:
        table = table[
            ~table.apply(
                lambda row: sum(
                    clean_text(row.iloc[i]) == clean_text(table.columns[i])
                    for i in range(min(len(row), len(table.columns)))
                ) >= max(2, len(table.columns) // 2),
                axis=1,
            )
        ]

    return table.reset_index(drop=True)


def make_record_id(row: pd.Series, business_columns: list[str]) -> str:
    canonical = "\x1f".join(clean_text(row.get(column, "")) for column in business_columns)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def save_results(current: pd.DataFrame) -> tuple[int, int]:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    now = datetime.now(TIMEZONE).isoformat(timespec="seconds")

    metadata_columns = {"record_id", "first_seen_at", "last_seen_at"}
    business_columns = [
        column for column in current.columns if column not in metadata_columns
    ]

    current["record_id"] = current.apply(
        lambda row: make_record_id(row, business_columns),
        axis=1,
    )
    current = current.drop_duplicates(subset=["record_id"]).copy()

    if CSV_PATH.exists():
        existing = pd.read_csv(CSV_PATH, dtype=str, keep_default_na=False)

        # Ако порталът добави нови колони, подравняваме старите и новите данни.
        all_business_columns = list(
            dict.fromkeys(
                [
                    *[c for c in existing.columns if c not in metadata_columns],
                    *business_columns,
                ]
            )
        )

        for column in all_business_columns:
            if column not in existing.columns:
                existing[column] = ""
            if column not in current.columns:
                current[column] = ""

        existing_ids = set(existing.get("record_id", pd.Series(dtype=str)).astype(str))
        new_rows = current[~current["record_id"].isin(existing_ids)].copy()

        new_rows["first_seen_at"] = now
        new_rows["last_seen_at"] = now

        current_ids = set(current["record_id"])
        if "last_seen_at" not in existing.columns:
            existing["last_seen_at"] = ""
        existing.loc[existing["record_id"].isin(current_ids), "last_seen_at"] = now

        ordered_columns = [
            "record_id",
            *all_business_columns,
            "first_seen_at",
            "last_seen_at",
        ]

        combined = pd.concat(
            [
                existing.reindex(columns=ordered_columns, fill_value=""),
                new_rows.reindex(columns=ordered_columns, fill_value=""),
            ],
            ignore_index=True,
        )
    else:
        new_rows = current.copy()
        new_rows["first_seen_at"] = now
        new_rows["last_seen_at"] = now

        ordered_columns = [
            "record_id",
            *business_columns,
            "first_seen_at",
            "last_seen_at",
        ]
        combined = new_rows.reindex(columns=ordered_columns, fill_value="")

    combined.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")
    combined.to_excel(XLSX_PATH, index=False)

    status = {
        "run_at": now,
        "records_received_from_portal": int(len(current)),
        "new_records_added": int(len(new_rows)),
        "total_unique_records_saved": int(len(combined)),
    }
    STATUS_PATH.write_text(
        json.dumps(status, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(json.dumps(status, ensure_ascii=False, indent=2))
    return len(new_rows), len(combined)


def main() -> int:
    try:
        current = fetch_table()
        if current.empty:
            raise RuntimeError("Порталът върна празна таблица; данните не са променени.")

        save_results(current)
        return 0
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
