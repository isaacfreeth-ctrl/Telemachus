#!/usr/bin/env python3
"""
Build a pre-indexed database of Dutch ministerial agenda appointments.

Data source: openlobby.nl (run by Open State Foundation), which scrapes
ministerial agendas from rijksoverheid.nl.

The resulting index can be searched instantly without any API calls.

Note: Dutch agenda data is patchy - not all appointments are published,
and quality varies across ministries.

Usage:
    python build_netherlands_index.py

Output:
    netherlands_agenda_index.json.gz - Compressed index of all appointments
"""

import json
import gzip
import re
import requests
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from html import unescape

OPENLOBBY_URL = "https://openlobby.nl/agenda/__data.json"


def strip_html(text):
    """Remove HTML tags and clean up text."""
    if not text:
        return ""
    text = re.sub(r'<[^>]+>', '', text)
    text = unescape(text)
    text = text.replace('\xa0', ' ')
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def parse_sveltekit_data(raw_json):
    """
    Parse SvelteKit dehydrated data format from openlobby.nl.

    SvelteKit serializes data as a flat array with struct dicts
    containing absolute index references (off by 1).

    Returns list of appointment dicts.
    """
    node = raw_json['nodes'][1]['data']
    values = node[1:]  # Strip the initial key map

    # values[0] = list of appointment indices
    appointment_indices = values[0]

    def resolve(idx):
        """Resolve a value, handling Date objects."""
        if idx < 0 or idx >= len(values):
            return None
        val = values[idx]
        if isinstance(val, list) and len(val) == 2 and val[0] == 'Date':
            return val[1]
        if isinstance(val, (dict, list)):
            return None  # Skip complex nested structures
        return val

    appointments = []
    for appt_start in appointment_indices:
        if appt_start < 1 or appt_start >= len(values):
            continue

        struct = values[appt_start - 1]
        if not isinstance(struct, dict):
            continue

        appt = {}
        for key, abs_idx in struct.items():
            # SvelteKit dehydrated format has absolute indices offset by 1
            real_idx = abs_idx - 1
            val = resolve(real_idx)
            if val is not None:
                appt[key] = val

        if appt.get('id') and appt.get('raw_text'):
            appointments.append(appt)

    return appointments


def build_index():
    """Main function to build the Netherlands agenda index."""

    print("=" * 60)
    print("Building Netherlands Ministerial Agenda Index")
    print("=" * 60)
    print()

    # Step 1: Fetch data from openlobby.nl
    print("Step 1: Fetching data from openlobby.nl...")
    print("  (This is a ~20MB download, may take a moment)")

    r = requests.get(OPENLOBBY_URL, timeout=120)
    r.raise_for_status()
    raw_data = r.json()
    print(f"  Downloaded {len(r.content) / 1024 / 1024:.1f} MB")

    # Step 2: Parse SvelteKit dehydrated format
    print("\nStep 2: Parsing SvelteKit data...")
    raw_appointments = parse_sveltekit_data(raw_data)
    print(f"  Parsed {len(raw_appointments)} appointments")

    # Step 3: Normalize into clean records
    print("\nStep 3: Normalizing records...")

    appointments = []
    seen_ids = set()

    for appt in raw_appointments:
        appt_id = appt.get('id')
        if appt_id in seen_ids:
            continue
        seen_ids.add(appt_id)

        # Clean the text description
        raw_text = strip_html(appt.get('raw_text', ''))
        if not raw_text:
            continue

        # Parse the date
        start_date = appt.get('start', '')
        if isinstance(start_date, str) and 'T' in start_date:
            start_date = start_date[:10]

        appointments.append({
            "id": appt_id,
            "date": start_date,
            "description": raw_text,
            "type": appt.get('type', ''),
            "location": appt.get('location', ''),
            "minister_surname": appt.get('minister_surname', ''),
            "minister_title": appt.get('minister_title', ''),
            "ministry": appt.get('minister_ministry', ''),
            "source_url": appt.get('url', ''),
        })

    print(f"  {len(raw_appointments)} -> {len(appointments)} after cleaning")

    # Step 4: Build search indexes
    print("\nStep 4: Building search indexes...")

    # Index by text content (for searching by organisation name)
    text_index = defaultdict(list)
    for i, appt in enumerate(appointments):
        text = f"{appt['description']} {appt.get('location', '')}".lower()
        for word in text.split():
            word = re.sub(r'[^\w]', '', word)
            if len(word) > 2:
                text_index[word].append(i)

    # Index by minister
    minister_index = defaultdict(list)
    for i, appt in enumerate(appointments):
        surname = appt.get('minister_surname', '').lower()
        if surname:
            minister_index[surname].append(i)

    index = {
        "metadata": {
            "created": datetime.now().isoformat(),
            "appointment_count": len(appointments),
            "source": "openlobby.nl (data from rijksoverheid.nl)",
            "coverage": "2023-present",
            "note": "Dutch ministerial agenda data is voluntary and may be incomplete.",
        },
        "appointments": appointments,
        "text_index": dict(text_index),
        "minister_index": dict(minister_index),
    }

    # Step 5: Save compressed
    output_path = Path(__file__).parent / "netherlands_agenda_index.json.gz"
    print(f"\nStep 5: Saving to {output_path}...")

    with gzip.open(output_path, "wt", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False)

    file_size_mb = output_path.stat().st_size / (1024 * 1024)
    print(f"  Saved ({file_size_mb:.1f} MB)")

    # Summary
    print()
    print("=" * 60)
    print("BUILD COMPLETE")
    print("=" * 60)
    print(f"  Total appointments: {len(appointments)}")
    print(f"  File size: {file_size_mb:.1f} MB")

    # Stats by minister
    minister_counts = defaultdict(int)
    for appt in appointments:
        surname = appt.get('minister_surname', 'Unknown')
        minister_counts[surname] += 1
    top_ministers = sorted(minister_counts.items(), key=lambda x: -x[1])[:10]
    print(f"\nTop ministers by appointment count:")
    for name, count in top_ministers:
        print(f"  {name}: {count}")

    # Sample
    print("\nSample appointments:")
    for appt in appointments[:5]:
        print(f"  [{appt['date']}] {appt['minister_surname']}: {appt['description'][:60]}")
        print(f"    Source: {appt['source_url'][:80]}")

    return index


if __name__ == "__main__":
    build_index()
