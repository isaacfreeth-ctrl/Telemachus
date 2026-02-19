#!/usr/bin/env python3
"""
Build a pre-indexed database of MEP meetings with interest representatives.

Downloads data from Integrity Watch (Transparency International EU), which
aggregates MEP meeting disclosures from the European Parliament.

Data source:
    https://www.integritywatch.eu/mepmeetings.php
    Raw JSON: https://integritywatch.eu/autoupdate_data_eu/mepmeetings/latest/mepmeetings.json

Run periodically (e.g. weekly via GitHub Actions) to keep up to date.

Usage:
    python build_mep_meetings_index.py

Output:
    mep_meetings_index.json.gz - Complete index of all MEP meetings
"""

import json
import gzip
import re
import sys
import requests
from datetime import datetime

SOURCE_URL = "https://www.integritywatch.eu/mepmeetings.php"
DATA_URL = "https://integritywatch.eu/autoupdate_data_eu/mepmeetings/latest/mepmeetings.json"


def normalize_date(date_str: str) -> str:
    """Normalize DD-MM-YYYY to DD/MM/YYYY."""
    if not date_str:
        return ""
    # Integrity Watch uses DD-MM-YYYY
    if re.match(r"\d{2}-\d{2}-\d{4}", date_str):
        parts = date_str.split("-")
        return f"{parts[0]}/{parts[1]}/{parts[2]}"
    return date_str


def build_index():
    print("=" * 60)
    print("MEP Meetings Index Builder (Integrity Watch)")
    print("=" * 60)

    print(f"\nDownloading meetings data from Integrity Watch...")
    response = requests.get(DATA_URL, timeout=60)
    response.raise_for_status()
    raw = response.json()
    print(f"  Downloaded {len(raw):,} raw records")

    meetings = []
    for item in raw:
        # Skip records with no MEP or lobbyist
        mep = item.get("mep", "").strip()
        lobbyists = item.get("lobbyists", "").strip()
        if not mep and not lobbyists:
            continue

        date = normalize_date(item.get("date", ""))

        # lobbyistsArray may list multiple orgs for one meeting
        lobbyists_array = item.get("lobbyistsArray", [])
        if not lobbyists_array and lobbyists:
            lobbyists_array = [lobbyists]

        meetings.append({
            "mep": mep,
            "epid": item.get("epid", ""),
            "date": date,
            "lobbyists": lobbyists,
            "lobbyists_array": lobbyists_array,
            "title": item.get("title", "").strip(),
            "dossier": item.get("dossier", "").strip(),
            "location": item.get("location", "").strip(),
            "group": item.get("group", "").strip(),
            "country": item.get("country", "").strip(),
            "committees": item.get("committees", []),
            "role": item.get("role", "").strip(),
        })

    print(f"  Kept {len(meetings):,} meetings after filtering")

    # Build metadata stats
    unique_meps = set(m["mep"] for m in meetings if m["mep"])
    unique_orgs = set()
    for m in meetings:
        for org in m["lobbyists_array"]:
            if org:
                unique_orgs.add(org)
    unique_groups = set(m["group"] for m in meetings if m["group"])
    unique_countries = set(m["country"] for m in meetings if m["country"])

    index = {
        "metadata": {
            "created": datetime.now().isoformat(),
            "source": SOURCE_URL,
            "data_url": DATA_URL,
            "total_meetings": len(meetings),
            "unique_meps": len(unique_meps),
            "unique_organisations": len(unique_orgs),
            "unique_groups": len(unique_groups),
            "unique_countries": len(unique_countries),
            "coverage": "2019-present",
        },
        "meetings": meetings,
    }

    output_path = "mep_meetings_index.json.gz"
    with gzip.open(output_path, "wt", encoding="utf-8") as f:
        json.dump(index, f)

    import os
    file_size = os.path.getsize(output_path)
    print(f"\nSaved to {output_path} ({file_size / 1024 / 1024:.1f} MB)")
    print(f"  {len(meetings):,} meetings")
    print(f"  {len(unique_meps):,} unique MEPs")
    print(f"  {len(unique_orgs):,} unique organisations")
    print(f"  {len(unique_groups):,} political groups")
    print(f"  {len(unique_countries):,} countries")

    return index


if __name__ == "__main__":
    build_index()
