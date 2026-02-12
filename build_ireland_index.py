#!/usr/bin/env python3
"""
Build a pre-indexed database of all Irish lobbying returns.

Data source: lobbyieng.com (visualises data from Ireland's official
lobbying register at lobbying.ie).

The resulting index can be searched instantly without any API calls.

Usage:
    python build_ireland_index.py

Output:
    ireland_lobbying_index.json.gz - Compressed index of all lobbying returns
"""

import json
import gzip
import re
import time
import requests
from pathlib import Path
from datetime import datetime
from collections import defaultdict

LOBBYIENG_BASE = "https://www.lobbyieng.com"
REQUEST_DELAY = 0.5  # Be polite to the API


def get_build_id():
    """Get the current Next.js build ID from lobbyieng.com."""
    r = requests.get(LOBBYIENG_BASE, timeout=15)
    r.raise_for_status()
    match = re.search(r'"buildId":"([^"]+)"', r.text)
    if match:
        return match.group(1)
    raise RuntimeError("Could not find Next.js build ID on lobbyieng.com")


def get_all_lobbyist_names():
    """Get the complete list of registered lobbyist names."""
    r = requests.get(f"{LOBBYIENG_BASE}/api/lobbyists",
                     headers={"Accept": "application/json"}, timeout=30)
    r.raise_for_status()
    return r.json()  # List of name strings


def slugify(name):
    """Convert a lobbyist name to a URL slug."""
    slug = name.lower().strip()
    slug = re.sub(r'[^a-z0-9\s\'-]', '', slug)
    slug = re.sub(r'\s+', '-', slug)
    slug = re.sub(r'-+', '-', slug)
    return slug.strip('-')


def fetch_lobbyist_returns(build_id, slug, max_pages=50):
    """
    Fetch all lobbying returns for a given lobbyist.

    Uses the Next.js data route to get paginated results.
    Returns list of return records.
    """
    all_records = []
    page = 1

    while page <= max_pages:
        url = f"{LOBBYIENG_BASE}/_next/data/{build_id}/lobbyists/{slug}.json"
        params = {"slug": slug}
        if page > 1:
            params["page"] = page

        try:
            r = requests.get(url, params=params, timeout=30)
            if r.status_code == 404:
                break
            r.raise_for_status()

            data = r.json().get("pageProps", {}).get("lobbyistData", {})
            records = data.get("records", [])
            total = data.get("total", 0)
            page_size = data.get("pageSize", 10)

            if not records:
                break

            all_records.extend(records)

            # Check if we've got all pages
            if len(all_records) >= total:
                break

            page += 1
            time.sleep(REQUEST_DELAY)

        except Exception as e:
            print(f"    Error fetching page {page} for {slug}: {e}")
            break

    return all_records


def build_index():
    """Main function to build the Ireland lobbying index."""

    print("=" * 60)
    print("Building Ireland Lobbying Index")
    print("=" * 60)
    print()

    # Step 1: Get build ID
    print("Step 1: Getting lobbyieng.com build ID...")
    build_id = get_build_id()
    print(f"  Build ID: {build_id}")

    # Step 2: Get all lobbyist names
    print("\nStep 2: Getting all registered lobbyist names...")
    lobbyist_names = get_all_lobbyist_names()
    print(f"  Found {len(lobbyist_names)} registered lobbyists")

    # Step 3: Fetch returns for each lobbyist
    print(f"\nStep 3: Fetching returns for all lobbyists...")
    print(f"  (This will take a while - ~{len(lobbyist_names) * 0.5 / 60:.0f} minutes minimum)")

    all_returns = []
    errors = []
    empty_count = 0

    for i, name in enumerate(lobbyist_names):
        if i % 100 == 0 and i > 0:
            print(f"  Processed {i}/{len(lobbyist_names)} lobbyists "
                  f"({len(all_returns)} returns so far)...")

        slug = slugify(name)
        if not slug:
            continue

        try:
            records = fetch_lobbyist_returns(build_id, slug)
            if records:
                all_returns.extend(records)
            else:
                empty_count += 1
        except Exception as e:
            errors.append((name, str(e)))

        time.sleep(REQUEST_DELAY)

    print(f"  Total returns fetched: {len(all_returns)}")
    if errors:
        print(f"  Errors: {len(errors)}")
        for name, err in errors[:5]:
            print(f"    - {name}: {err}")
    print(f"  Lobbyists with no returns: {empty_count}")

    # Step 4: Normalize and deduplicate
    print("\nStep 4: Normalizing and deduplicating...")

    seen_ids = set()
    unique_returns = []

    for ret in all_returns:
        ret_id = ret.get("id")
        if ret_id and ret_id not in seen_ids:
            seen_ids.add(ret_id)

            # Flatten DPO entries for searchability
            officials = []
            for dpo in ret.get("dpo_entries", []):
                officials.append({
                    "name": dpo.get("person_name", ""),
                    "title": dpo.get("job_title", ""),
                    "body": dpo.get("public_body", ""),
                })

            # Parse lobbying methods
            methods = []
            for activity in ret.get("lobbying_activities", []):
                # Format: "|Meeting|2-5" or "|Email|1"
                parts = activity.strip("|").split("|")
                if parts:
                    methods.append(parts[0])

            unique_returns.append({
                "id": ret_id,
                "lobbyist": ret.get("lobbyist_name", ""),
                "date": ret.get("date_published", "")[:10],
                "subject": ret.get("specific_details", ""),
                "intended_result": ret.get("intended_results", ""),
                "officials": officials,
                "methods": list(set(methods)),
                "source_url": f"https://www.{ret.get('url', '')}",
            })

    print(f"  {len(all_returns)} -> {len(unique_returns)} after deduplication")

    # Step 5: Build search index
    print("\nStep 5: Building search indexes...")

    lobbyist_index = defaultdict(list)
    for i, ret in enumerate(unique_returns):
        text = ret["lobbyist"].lower()
        for word in text.split():
            word = re.sub(r'[^\w]', '', word)
            if len(word) > 1:
                lobbyist_index[word].append(i)

    index = {
        "metadata": {
            "created": datetime.now().isoformat(),
            "return_count": len(unique_returns),
            "lobbyists_processed": len(lobbyist_names),
            "errors": len(errors),
            "source": "lobbyieng.com (data from lobbying.ie)",
            "coverage": "2015-present",
        },
        "returns": unique_returns,
        "lobbyist_index": dict(lobbyist_index),
    }

    # Step 6: Save compressed
    output_path = Path(__file__).parent / "ireland_lobbying_index.json.gz"
    print(f"\nStep 6: Saving to {output_path}...")

    with gzip.open(output_path, "wt", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False)

    file_size_mb = output_path.stat().st_size / (1024 * 1024)
    print(f"  Saved ({file_size_mb:.1f} MB)")

    # Summary
    print()
    print("=" * 60)
    print("BUILD COMPLETE")
    print("=" * 60)
    print(f"  Total returns: {len(unique_returns)}")
    print(f"  Lobbyists: {len(lobbyist_names)}")
    print(f"  File size: {file_size_mb:.1f} MB")

    # Sample
    print("\nSample returns:")
    for ret in unique_returns[:3]:
        officials_str = ", ".join(o["name"] for o in ret["officials"][:3])
        print(f"  {ret['lobbyist']}: {ret['subject'][:50]}")
        print(f"    Officials: {officials_str}")
        print(f"    Source: {ret['source_url']}")

    return index


if __name__ == "__main__":
    build_index()
