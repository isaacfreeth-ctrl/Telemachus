#!/usr/bin/env python3
"""
Build a pre-indexed database of all UK ministerial meetings.

This script downloads ALL ministerial meeting CSVs from GOV.UK,
combines them into a single searchable JSON file, and saves it.

The resulting index can be searched instantly without any API calls.

Run this script periodically (e.g., weekly via GitHub Actions) to keep
the index up to date.

Usage:
    python build_uk_index.py
    
Output:
    uk_meetings_index.json - Complete index of all meetings
"""

import json
import csv
import io
import requests
from pathlib import Path
from datetime import datetime
from collections import defaultdict

# GOV.UK API endpoints
GOVUK_SEARCH_URL = "https://www.gov.uk/api/search.json"
GOVUK_CONTENT_URL = "https://www.gov.uk/api/content"


def discover_publications(publication_type="ministerial"):
    """Discover all transparency publications from GOV.UK."""
    
    if publication_type == "ministerial":
        query = "ministerial meetings transparency"
    else:
        query = "senior officials meetings transparency"
    
    publications = []
    start = 0
    page_size = 100
    
    while True:
        params = {
            "q": query,
            "filter_format": "transparency",
            "count": page_size,
            "start": start,
            "fields": "title,link,public_timestamp,organisations"
        }
        
        response = requests.get(GOVUK_SEARCH_URL, params=params, timeout=60)
        response.raise_for_status()
        data = response.json()
        
        results = data.get("results", [])
        if not results:
            break
            
        for r in results:
            # Filter for meetings publications
            title_lower = r.get("title", "").lower()
            if "meeting" in title_lower:
                if publication_type == "ministerial" and "senior official" not in title_lower:
                    publications.append({
                        "title": r.get("title", ""),
                        "link": r.get("link", ""),
                        "date": r.get("public_timestamp", "")[:10] if r.get("public_timestamp") else "",
                        "organisations": r.get("organisations", [])
                    })
                elif publication_type == "senior_officials" and "senior official" in title_lower:
                    publications.append({
                        "title": r.get("title", ""),
                        "link": r.get("link", ""),
                        "date": r.get("public_timestamp", "")[:10] if r.get("public_timestamp") else "",
                        "organisations": r.get("organisations", [])
                    })
        
        start += page_size
        if start >= data.get("total", 0):
            break
    
    return publications


def get_csv_urls_from_publication(pub_path):
    """Extract CSV URLs from a publication page."""
    
    if not pub_path.startswith("/"):
        pub_path = "/" + pub_path
    
    url = f"{GOVUK_CONTENT_URL}{pub_path}"
    
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        data = response.json()
    except:
        return []
    
    csv_urls = []
    
    # Extract department from path
    dept = pub_path.split("/")[2] if len(pub_path.split("/")) > 2 else "unknown"
    dept = dept.replace("-", " ").title()
    
    # Check documents
    for doc in data.get("details", {}).get("documents", []):
        if isinstance(doc, str) and ".csv" in doc.lower():
            # Extract URL from HTML
            import re
            urls = re.findall(r'href=["\']([^"\']*\.csv[^"\']*)["\']', doc, re.IGNORECASE)
            for u in urls:
                if u.startswith("/"):
                    u = "https://www.gov.uk" + u
                csv_urls.append((dept, u))
    
    # Check attachments
    for attachment in data.get("details", {}).get("attachments", []):
        url = attachment.get("url", "")
        if ".csv" in url.lower():
            if url.startswith("/"):
                url = "https://www.gov.uk" + url
            csv_urls.append((dept, url))
    
    return csv_urls


def download_and_parse_csv(url):
    """Download a CSV and parse all meetings from it."""
    
    meetings = []
    
    try:
        response = requests.get(url, timeout=60)
        response.raise_for_status()
        
        # Handle encoding
        content = response.content.decode('utf-8-sig', errors='replace')
        
        reader = csv.DictReader(io.StringIO(content))
        
        for row in reader:
            # Normalize column names
            row_lower = {k.lower().strip(): v for k, v in row.items() if k}
            
            # Extract fields with various column name variations
            minister = (
                row_lower.get("minister", "") or 
                row_lower.get("name", "") or 
                row_lower.get("official", "") or
                row_lower.get("minister's name", "") or
                ""
            ).strip()
            
            date = (
                row_lower.get("date", "") or 
                row_lower.get("date of meeting", "") or 
                row_lower.get("meeting date", "") or
                ""
            ).strip()
            
            org = (
                row_lower.get("organisation", "") or 
                row_lower.get("organizations", "") or 
                row_lower.get("name of organisation", "") or
                row_lower.get("name of external organisation", "") or
                row_lower.get("external organisation", "") or
                ""
            ).strip()
            
            purpose = (
                row_lower.get("purpose", "") or 
                row_lower.get("purpose of meeting", "") or
                ""
            ).strip()
            
            if org and (minister or date):
                meetings.append({
                    "minister": minister,
                    "date": date,
                    "organisation": org,
                    "purpose": purpose
                })
    
    except Exception as e:
        print(f"  Error processing {url}: {e}")
    
    return meetings


def build_index():
    """Build complete index of all UK ministerial meetings."""
    
    print("=" * 60)
    print("Building UK Ministerial Meetings Index")
    print("=" * 60)
    print(f"Started: {datetime.now().isoformat()}")
    print()
    
    # Step 1: Discover ministerial publications
    print("Step 1: Discovering ministerial publications...")
    ministerial_pubs = discover_publications("ministerial")
    print(f"  Found {len(ministerial_pubs)} ministerial publications")
    
    # Step 2: Discover senior officials publications
    print("Step 2: Discovering senior officials publications...")
    senior_pubs = discover_publications("senior_officials")
    print(f"  Found {len(senior_pubs)} senior officials publications")
    
    # Step 3: Get all CSV URLs
    print("Step 3: Extracting CSV URLs...")
    all_csvs = []
    
    for i, pub in enumerate(ministerial_pubs):
        if i % 50 == 0:
            print(f"  Processing ministerial publication {i}/{len(ministerial_pubs)}...")
        csv_urls = get_csv_urls_from_publication(pub["link"])
        for dept, url in csv_urls:
            all_csvs.append({"dept": dept, "url": url, "type": "ministerial"})
    
    for i, pub in enumerate(senior_pubs):
        if i % 50 == 0:
            print(f"  Processing senior officials publication {i}/{len(senior_pubs)}...")
        csv_urls = get_csv_urls_from_publication(pub["link"])
        for dept, url in csv_urls:
            all_csvs.append({"dept": dept, "url": url, "type": "senior_officials"})
    
    # Deduplicate
    seen_urls = set()
    unique_csvs = []
    for c in all_csvs:
        if c["url"] not in seen_urls:
            seen_urls.add(c["url"])
            unique_csvs.append(c)
    
    print(f"  Found {len(unique_csvs)} unique CSV files")
    
    # Step 4: Download and parse all CSVs
    print("Step 4: Downloading and parsing CSVs...")
    all_meetings = []
    
    for i, csv_info in enumerate(unique_csvs):
        if i % 50 == 0:
            print(f"  Processing CSV {i}/{len(unique_csvs)}...")
        
        meetings = download_and_parse_csv(csv_info["url"])
        
        for m in meetings:
            m["department"] = csv_info["dept"]
            m["meeting_type"] = csv_info["type"]
            all_meetings.append(m)
    
    print(f"  Parsed {len(all_meetings)} total meetings")
    
    # Step 5: Deduplicate meetings
    print("Step 5: Deduplicating...")
    seen = set()
    unique_meetings = []
    
    for m in all_meetings:
        key = (m["minister"], m["date"], m["organisation"])
        if key not in seen:
            seen.add(key)
            unique_meetings.append(m)
    
    print(f"  {len(unique_meetings)} unique meetings after deduplication")
    
    # Step 6: Build index
    print("Step 6: Building searchable index...")
    
    # Create organisation lookup for fast searching
    org_index = defaultdict(list)
    for i, m in enumerate(unique_meetings):
        org_lower = m["organisation"].lower()
        # Index by words
        for word in org_lower.split():
            if len(word) > 2:
                org_index[word].append(i)
    
    index = {
        "metadata": {
            "created": datetime.now().isoformat(),
            "meeting_count": len(unique_meetings),
            "ministerial_publications": len(ministerial_pubs),
            "senior_officials_publications": len(senior_pubs),
            "csv_files_processed": len(unique_csvs),
            "coverage": "2012-present"
        },
        "meetings": unique_meetings,
        "org_index": {k: v for k, v in org_index.items()}  # Convert defaultdict
    }
    
    # Step 7: Save
    output_path = Path(__file__).parent / "uk_meetings_index.json"
    print(f"Step 7: Saving to {output_path}...")
    
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False)
    
    file_size = output_path.stat().st_size / 1024 / 1024
    print(f"  Saved ({file_size:.1f} MB)")
    
    print()
    print("=" * 60)
    print("Index build complete!")
    print(f"  Total meetings: {len(unique_meetings)}")
    print(f"  File size: {file_size:.1f} MB")
    print("=" * 60)
    
    return index


if __name__ == "__main__":
    build_index()
