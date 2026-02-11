#!/usr/bin/env python3
"""
Build a pre-indexed database of all UK ministerial and senior officials meetings.

This script downloads ALL ministerial and senior officials meeting CSVs from GOV.UK,
combines them into a single searchable JSON file, and saves it.

The resulting index can be searched instantly without any API calls.

Data sources:
- GOV.UK transparency publications (ministerial meetings)
- GOV.UK transparency publications (senior officials meetings)

Run this script periodically to keep the index up to date:
    python build_uk_index.py

Output:
    uk_meetings_index.json - Complete index of all meetings with source URLs
"""

import json
import csv
import io
import re
import sys
import time
import requests
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from urllib.parse import quote, unquote

# GOV.UK API endpoints
GOVUK_SEARCH_URL = "https://www.gov.uk/api/search.json"
GOVUK_CONTENT_URL = "https://www.gov.uk/api/content"

# Rate limiting
REQUEST_DELAY = 0.3  # seconds between requests to be polite


def discover_publications(query, label=""):
    """
    Discover transparency publications from GOV.UK Search API.
    
    Pages through all results (GOV.UK returns max 1000 per query).
    Returns list of dicts with 'link' and 'title' keys.
    """
    publications = []
    start = 0
    page_size = 100
    
    print(f"  Discovering {label} publications...")
    
    while True:
        try:
            r = requests.get(GOVUK_SEARCH_URL, params={
                'filter_format': 'transparency',
                'q': query,
                'count': page_size,
                'start': start,
                'fields': 'link,title,public_timestamp'
            }, timeout=30)
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            print(f"    Search API error at offset {start}: {e}")
            break
        
        results = data.get('results', [])
        if not results:
            break
        
        for res in results:
            link = res.get('link', '')
            title = res.get('title', '')
            if link and title:
                publications.append({
                    'link': link,
                    'title': title,
                    'timestamp': res.get('public_timestamp', '')
                })
        
        start += page_size
        total = data.get('total', 0)
        
        if start >= total or start >= 1000:  # GOV.UK caps at 1000
            break
        
        time.sleep(REQUEST_DELAY)
    
    print(f"    Found {len(publications)} publications")
    return publications


def get_csv_urls_from_publication(pub_link):
    """
    Extract CSV download URLs from a publication's Content API response.
    
    Handles BOTH URL formats:
    - New: https://assets.publishing.service.gov.uk/media/XXXXX/filename.csv
    - Old: https://assets.publishing.service.gov.uk/government/uploads/system/uploads/attachment_data/file/XXXXX/filename.csv
    
    Returns list of tuples: (csv_url, publication_url)
    where publication_url is the human-readable GOV.UK page.
    """
    if not pub_link.startswith("/"):
        pub_link = "/" + pub_link
    
    content_url = f"{GOVUK_CONTENT_URL}{pub_link}"
    publication_url = f"https://www.gov.uk{pub_link}"
    
    try:
        r = requests.get(content_url, timeout=30)
        r.raise_for_status()
        data = r.json()
    except requests.exceptions.HTTPError as e:
        return [], f"HTTP {e.response.status_code} for {content_url}"
    except Exception as e:
        return [], f"Error fetching {content_url}: {e}"
    
    # Skip collection pages (they don't have CSVs directly)
    schema = data.get('schema_name', '')
    if schema == 'document_collection':
        return [], None  # Not an error, just a collection
    
    csv_urls = set()
    details = data.get('details', {})
    
    # Method 1: Check attachments (most reliable)
    for att in details.get('attachments', []):
        url = att.get('url', '')
        if '.csv' in url.lower():
            # Normalize URL
            if url.startswith('/'):
                url = f"https://www.gov.uk{url}"
            csv_urls.add(url)
    
    # Method 2: Check documents HTML blocks (fallback)
    for doc in details.get('documents', []):
        if not isinstance(doc, str):
            continue
        
        # Match BOTH old and new URL formats
        # New: https://assets.publishing.service.gov.uk/media/...csv
        # Old: https://assets.publishing.service.gov.uk/government/uploads/...csv
        found = re.findall(
            r'href="(https://assets\.publishing\.service\.gov\.uk/[^"]*\.csv[^"]*)"',
            doc, re.IGNORECASE
        )
        for url in found:
            # Clean URL-encoded characters but preserve the structure
            csv_urls.add(url)
        
        # Also check for relative URLs starting with /
        relative = re.findall(r'href="(/[^"]*\.csv[^"]*)"', doc, re.IGNORECASE)
        for url in relative:
            # Skip preview URLs
            if '/csv-preview/' in url:
                continue
            csv_urls.add(f"https://www.gov.uk{url}")
    
    return [(url, publication_url) for url in csv_urls], None


def download_and_parse_csv(url, department="Unknown", meeting_type="ministerial", source_url=""):
    """
    Download a CSV file and parse meeting records from it.
    
    Handles various CSV formats, encodings, and column name variations
    used across different government departments.
    
    Returns: (meetings_list, error_string_or_None)
    """
    meetings = []
    
    try:
        r = requests.get(url, timeout=60)
        r.raise_for_status()
    except requests.exceptions.HTTPError as e:
        return [], f"HTTP {e.response.status_code}"
    except Exception as e:
        return [], str(e)
    
    try:
        # Handle various encodings (BOM, latin-1, etc.)
        content = r.content.decode('utf-8-sig', errors='replace')
        
        reader = csv.DictReader(io.StringIO(content))
        
        for row in reader:
            # Normalize column names to lowercase
            row_lower = {k.lower().strip(): v.strip() if v else '' for k, v in row.items() if k}
            
            # Extract minister/official name (various column names across departments)
            minister = (
                row_lower.get("minister", "") or
                row_lower.get("minister's name", "") or
                row_lower.get("name", "") or
                row_lower.get("official", "") or
                row_lower.get("senior official", "") or
                row_lower.get("name of senior official", "") or
                ""
            ).strip()
            
            # Extract date
            date = (
                row_lower.get("date", "") or
                row_lower.get("date of meeting", "") or
                row_lower.get("meeting date", "") or
                ""
            ).strip()
            
            # Extract organisation
            # Column names vary across departments and years:
            #   "Name of Individual or Organisation" (2024+ standard)
            #   "Name of organisation or individual" (2022-2023)
            #   "Organisation" (some departments)
            #   "Name of External Organisation" (some older formats)
            org = (
                row_lower.get("name of individual or organisation", "") or
                row_lower.get("name of organisation or individual", "") or
                row_lower.get("organisation", "") or
                row_lower.get("organizations", "") or
                row_lower.get("organisations", "") or
                row_lower.get("name of organisation", "") or
                row_lower.get("name of external organisation", "") or
                row_lower.get("external organisation", "") or
                row_lower.get("organisation(s)", "") or
                ""
            ).strip()
            
            # Extract purpose
            purpose = (
                row_lower.get("purpose of meeting", "") or
                row_lower.get("purpose", "") or
                ""
            ).strip()
            
            # Skip empty/nil rows
            if not org or org.lower() in ('nil', 'nil return', 'n/a', ''):
                continue
            if not minister and not date:
                continue
            if minister.lower() in ('nil return', 'nil'):
                continue
            
            meetings.append({
                "minister": minister,
                "date": date,
                "organisation": org,
                "purpose": purpose,
                "department": department,
                "meeting_type": meeting_type,
                "source_url": source_url
            })
    
    except Exception as e:
        return meetings, f"Parse error: {e}"
    
    return meetings, None


def extract_department_from_publication(title, link):
    """
    Extract department name from a publication title or URL path.
    
    Common patterns:
    - "Cabinet Office: ministerial meetings..."
    - "DSIT ministerial meetings..."  
    - "Home Office's ministerial meetings..."
    """
    # Try to get department from title before the colon or keyword
    dept_patterns = [
        r'^(.+?):\s*ministerial',
        r'^(.+?):\s*senior official',
        r'^(.+?)\s+ministerial',
        r'^(.+?)\s+senior official',
        r"^(.+?)'s\s+ministerial",
        r"^(.+?)'s\s+senior official",
    ]
    
    for pattern in dept_patterns:
        m = re.match(pattern, title, re.IGNORECASE)
        if m:
            dept = m.group(1).strip()
            # Clean up common abbreviations
            dept = dept.rstrip(':').strip()
            if len(dept) > 3:  # Skip if too short (probably just "HMT" etc, keep those)
                return dept
            return dept
    
    # Fallback: extract from URL path
    parts = link.strip('/').split('/')
    if len(parts) >= 3:
        return parts[-1].split('-ministerial')[0].replace('-', ' ').title()[:50]
    
    return "Unknown"


def is_meetings_csv(url, title=""):
    """Check if a CSV URL is likely a meetings file (not gifts/travel/hospitality)."""
    url_lower = url.lower()
    title_lower = title.lower()
    
    # Exclude non-meeting CSVs
    exclude_terms = ['gift', 'hospitality', 'travel', 'overseas', 'expense']
    for term in exclude_terms:
        if term in url_lower:
            return False
    
    # Include if it has meeting indicators
    include_terms = ['meeting', 'transparency']
    for term in include_terms:
        if term in url_lower:
            return True
    
    # If URL doesn't clearly indicate, it might still be a meetings file
    # (some departments just use generic names)
    return True


def build_index():
    """Main function to build the complete UK meetings index."""
    
    print("=" * 60)
    print("Building UK Ministerial & Senior Officials Meetings Index")
    print("=" * 60)
    print()
    
    # ---- Step 1: Discover publications ----
    print("Step 1: Discovering publications...")
    
    ministerial_pubs = discover_publications(
        "ministerial meetings", 
        label="ministerial"
    )
    
    senior_pubs = discover_publications(
        "senior officials meetings transparency", 
        label="senior officials"
    )
    
    # Deduplicate by link
    all_pubs = {}
    for pub in ministerial_pubs:
        key = pub['link']
        if key not in all_pubs:
            all_pubs[key] = {**pub, 'type': 'ministerial'}
    
    for pub in senior_pubs:
        key = pub['link']
        if key not in all_pubs:
            all_pubs[key] = {**pub, 'type': 'senior_official'}
        else:
            # Already found as ministerial, might contain both
            all_pubs[key]['type'] = 'both'
    
    print(f"  Total unique publications: {len(all_pubs)}")
    print()
    
    # ---- Step 2: Extract CSV URLs from each publication ----
    print("Step 2: Extracting CSV URLs from publications...")
    
    all_csvs = []  # list of dicts: {url, pub_url, dept, type}
    content_errors = []  # (pub_title, error)
    pubs_without_csvs = 0
    
    for i, (link, pub) in enumerate(all_pubs.items()):
        if i % 50 == 0 and i > 0:
            print(f"  Processed {i}/{len(all_pubs)} publications...")
        
        dept = extract_department_from_publication(pub['title'], link)
        csv_results, error = get_csv_urls_from_publication(link)
        
        if error:
            content_errors.append((pub['title'], error))
            continue
        
        if not csv_results:
            pubs_without_csvs += 1
            continue
        
        for csv_url, pub_url in csv_results:
            # Filter to meetings CSVs only
            if is_meetings_csv(csv_url, pub.get('title', '')):
                all_csvs.append({
                    'url': csv_url,
                    'pub_url': pub_url,
                    'dept': dept,
                    'type': pub['type']
                })
        
        time.sleep(REQUEST_DELAY)
    
    # Deduplicate CSV URLs
    seen_urls = set()
    unique_csvs = []
    for csv_info in all_csvs:
        if csv_info['url'] not in seen_urls:
            seen_urls.add(csv_info['url'])
            unique_csvs.append(csv_info)
    
    print(f"  Found {len(unique_csvs)} unique meeting CSV files")
    if content_errors:
        print(f"  Content API errors: {len(content_errors)}")
        for title, err in content_errors[:5]:
            print(f"    - {title[:60]}: {err}")
        if len(content_errors) > 5:
            print(f"    ... and {len(content_errors) - 5} more")
    if pubs_without_csvs:
        print(f"  Publications with no CSVs: {pubs_without_csvs}")
    print()
    
    # ---- Step 3: Download and parse all CSVs ----
    print("Step 3: Downloading and parsing CSV files...")
    
    all_meetings = []
    download_errors = []  # (csv_url, error)
    
    for i, csv_info in enumerate(unique_csvs):
        if i % 25 == 0:
            print(f"  Processing CSV {i}/{len(unique_csvs)}...")
        
        meeting_type = "ministerial" if csv_info['type'] != 'senior_official' else "senior_official"
        
        meetings, error = download_and_parse_csv(
            url=csv_info['url'],
            department=csv_info['dept'],
            meeting_type=meeting_type,
            source_url=csv_info['pub_url']
        )
        
        if error:
            download_errors.append((csv_info['url'], error))
        
        all_meetings.extend(meetings)
        time.sleep(REQUEST_DELAY)
    
    print(f"  Parsed {len(all_meetings)} total meeting records")
    if download_errors:
        print(f"  Download/parse errors: {len(download_errors)}")
        for url, err in download_errors[:10]:
            filename = url.split('/')[-1][:60]
            print(f"    - {filename}: {err}")
        if len(download_errors) > 10:
            print(f"    ... and {len(download_errors) - 10} more")
    print()
    
    # ---- Step 4: Deduplicate meetings ----
    print("Step 4: Deduplicating meetings...")
    
    seen = set()
    unique_meetings = []
    
    for m in all_meetings:
        # Create a dedup key from core fields
        key = (
            m['minister'].lower().strip(),
            m['date'].strip(),
            m['organisation'].lower().strip(),
            m['purpose'].lower().strip()[:100]
        )
        
        if key not in seen:
            seen.add(key)
            unique_meetings.append(m)
    
    print(f"  {len(all_meetings)} -> {len(unique_meetings)} after deduplication")
    print()
    
    # ---- Step 5: Build searchable index ----
    print("Step 5: Building search indexes...")
    
    org_index = defaultdict(list)
    for i, m in enumerate(unique_meetings):
        text = m['organisation'].lower()
        for word in text.split():
            word = re.sub(r'[^\w]', '', word)
            if len(word) > 1:
                org_index[word].append(i)
    
    index = {
        "metadata": {
            "created": datetime.now().isoformat(),
            "meeting_count": len(unique_meetings),
            "publications_processed": len(all_pubs),
            "csv_files_processed": len(unique_csvs),
            "content_api_errors": len(content_errors),
            "download_errors": len(download_errors),
            "coverage": "2010-present"
        },
        "meetings": unique_meetings,
        "org_index": dict(org_index)
    }
    
    # ---- Step 6: Save (gzip compressed to stay under GitHub's 25MB limit) ----
    output_path = Path(__file__).parent / "uk_meetings_index.json.gz"
    print(f"Step 6: Saving to {output_path}...")
    
    import gzip
    with gzip.open(output_path, "wt", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False)
    
    file_size_mb = output_path.stat().st_size / (1024 * 1024)
    print(f"  Saved ({file_size_mb:.1f} MB)")
    
    # ---- Summary ----
    print()
    print("=" * 60)
    print("BUILD COMPLETE")
    print("=" * 60)
    print(f"  Total meetings: {len(unique_meetings)}")
    print(f"  Publications: {len(all_pubs)}")
    print(f"  CSV files: {len(unique_csvs)}")
    print(f"  File size: {file_size_mb:.1f} MB")
    
    if content_errors or download_errors:
        print()
        print("ERRORS SUMMARY:")
        if content_errors:
            print(f"  {len(content_errors)} publications could not be fetched from Content API")
        if download_errors:
            print(f"  {len(download_errors)} CSV files could not be downloaded or parsed")
    
    # Sample data
    print()
    print("Sample meetings:")
    for m in unique_meetings[:5]:
        print(f"  [{m['department']}] {m['minister']}: {m['organisation'][:40]} ({m['date']})")
        if m.get('source_url'):
            print(f"    Source: {m['source_url'][:80]}")
    
    return index


if __name__ == "__main__":
    build_index()
