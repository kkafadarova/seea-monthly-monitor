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


def find_initial_query_id(html: str) -> str:
    patterns = [
        r'["\']searchQueryId["\']\s*[:=]\s*["\']([0-9a-fA-F-]{36})["\']',
        r'searchQueryId=([0-9a-fA-F-]{36})',
        r'value=["\']([0-9a-fA-F-]{36})["\']',
    ]
    for pattern in patterns:
        match = re.search(pattern, html, flags=re.IGNORECASE)
        if match:
            return match.group(1)
    raise RuntimeError("Не беше намерен началният searchQueryId.")


def find_export_query_id(html: str) -> str:
    patterns = [
        r'ExportKendoTable\?searchQueryId=([0-9a-fA-F-]{36})',
        r'ExportKendoTable(?:&amp;|&|\?)searchQueryId(?:=|%3D)([0-9a-fA-F-]{36})',
        r'["\']searchQueryId["\']\s*[:=]\s*["\']([0-9a-fA-F-]{36})["\']',
    ]
    for pattern in patterns:
        match = re.search(pattern, html, flags=re.IGNORECASE)
        if match:
            return match.group(1)
    raise RuntimeError(
        "Не беше намерен export searchQueryId в отговора на Search заявката."
    )


def flatten_columns(columns: pd.Index) -> list[str]:
    names, used = [], {}
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
        names.append(f"{name}_{count + 1}" if count else name)
    return names


def read_export(content: bytes) -> pd.DataFrame:
    # Истински .xls (OLE/BIFF) започва с D0 CF 11 E0.
    if content.startswith(bytes.fromhex("D0CF11E0")):
        return pd.read_excel(io.BytesIO(content), engine="xlrd")

    # Някои сървъри изпращат HTML таблица с .xls разширение.
    text_start = content[:500].lower()
    if b"<html" in text_start or b"<table" in text_start:
        tables = pd.read_html(io.BytesIO(content))
        if tables:
            return max(tables, key=lambda frame: len(frame.index))

    raise RuntimeError(
        "Export endpoint-ът не върна валиден .xls или HTML таблица. "
        f"Получени са {len(content)} bytes; начало={content[:20]!r}"
    )


def fetch_all_records() -> pd.DataFrame:
    today = datetime.now(TIMEZONE).strftime("%d.%m.%Y")
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0 Safari/537.36"
        ),
        "Accept-Language": "bg-BG,bg;q=0.9,en;q=0.8",
        "Referer": PAGE_URL,
    }

    with requests.Session() as session:
        page = session.get(PAGE_URL, headers=headers, timeout=60)
        page.raise_for_status()
        initial_id = find_initial_query_id(page.text)

        params = {
            "searchQueryId": initial_id,
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

        search = session.get(
            SEARCH_URL,
            params=params,
            headers={
                **headers,
                "Accept": "text/html, */*; q=0.01",
                "X-Requested-With": "XMLHttpRequest",
            },
            timeout=120,
        )
        search.raise_for_status()

        # ВАЖНО: Search връща НОВ GUID за Excel export.
        # Той е различен от GUID-а, изпратен към Search.
        export_id = find_export_query_id(search.text)

        export = session.get(
            EXPORT_URL,
            params={"searchQueryId": export_id},
            headers={
                **headers,
                "Accept": (
                    "text/html,application/xhtml+xml,application/xml;q=0.9,"
                    "image/avif,image/webp,*/*;q=0.8"
                ),
            },
            timeout=180,
        )
        export.raise_for_status()

        print(
            json.dumps(
                {
                    "initial_search_query_id": initial_id,
                    "export_search_query_id": export_id,
                    "export_bytes": len(export.content),
                    "content_type": export.headers.get("Content-Type", ""),
                    "content_disposition": export.headers.get(
                        "Content-Disposition", ""
                    ),
                },
                ensure_ascii=False,
                indent=2,
            )
        )

    frame = read_export(export.content).copy()
    frame.columns = flatten_columns(frame.columns)
    frame = frame.dropna(how="all")

    for column in frame.columns:
        frame[column] = frame[column].map(clean_text)

    return frame.reset_index(drop=True)


def make_record_id(row: pd.Series, columns: list[str]) -> str:
    value = "\x1f".join(clean_text(row.get(column, "")) for column in columns)
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def save_results(current: pd.DataFrame) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    now = datetime.now(TIMEZONE).isoformat(timespec="seconds")
    metadata = {"record_id", "first_seen_at", "last_seen_at"}
    business_columns = [c for c in current.columns if c not in metadata]

    current["record_id"] = current.apply(
        lambda row: make_record_id(row, business_columns), axis=1
    )
    current = current.drop_duplicates(subset=["record_id"]).copy()

    if CSV_PATH.exists():
        existing = pd.read_csv(CSV_PATH, dtype=str, keep_default_na=False)
        all_business_columns = list(
            dict.fromkeys(
                [c for c in existing.columns if c not in metadata] + business_columns
            )
        )

        for column in all_business_columns:
            if column not in existing:
                existing[column] = ""
            if column not in current:
                current[column] = ""

        existing_ids = set(existing["record_id"].astype(str))
        new_rows = current[~current["record_id"].isin(existing_ids)].copy()
        new_rows["first_seen_at"] = now
        new_rows["last_seen_at"] = now

        if "last_seen_at" not in existing:
            existing["last_seen_at"] = ""
        existing.loc[
            existing["record_id"].isin(set(current["record_id"])), "last_seen_at"
        ] = now

        columns = [
            "record_id",
            *all_business_columns,
            "first_seen_at",
            "last_seen_at",
        ]
        combined = pd.concat(
            [
                existing.reindex(columns=columns, fill_value=""),
                new_rows.reindex(columns=columns, fill_value=""),
            ],
            ignore_index=True,
        )
    else:
        new_rows = current.copy()
        new_rows["first_seen_at"] = now
        new_rows["last_seen_at"] = now
        columns = [
            "record_id",
            *business_columns,
            "first_seen_at",
            "last_seen_at",
        ]
        combined = new_rows.reindex(columns=columns, fill_value="")

    combined.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")
    combined.to_excel(XLSX_PATH, index=False)
    new_rows.to_excel(NEW_RECORDS_PATH, index=False)

    status = {
        "run_at": now,
        "records_received_from_portal": len(current),
        "new_records_added": len(new_rows),
        "total_unique_records_saved": len(combined),
    }
    STATUS_PATH.write_text(
        json.dumps(status, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(status, ensure_ascii=False, indent=2))


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
