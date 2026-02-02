"""
Jurisdiction modules for European Lobbying Tracker.
Imports search functions from the proven eu_lobbying_core module.
"""

import sys
from pathlib import Path

# Import from the core module
from eu_lobbying_core import (
    search_eu_register,
    fetch_eu_data,
    search_france_register,
    fetch_france_data,
    search_germany_register,
    fetch_germany_data,
    search_uk_ministerial_meetings,
    search_austria_register,
    search_catalonia_register,
    search_finland_register,
    search_slovenia_register,
)


def search_eu(search_term: str) -> dict:
    """Search EU register and fetch data."""
    results = search_eu_register(search_term)
    if not results:
        return None
    
    # Take best match
    best = results[0]
    eu_id = best.get("id")
    if not eu_id:
        return None
    
    return fetch_eu_data(eu_id)


def search_france(search_term: str) -> dict:
    """Search France HATVP and fetch data."""
    results = search_france_register(search_term)
    if not results:
        return None
    
    best = results[0]
    fr_id = best.get("id")
    if not fr_id:
        return None
    
    return fetch_france_data(fr_id)


def search_germany(search_term: str) -> dict:
    """Search Germany Lobbyregister and fetch data."""
    results = search_germany_register(search_term)
    if not results:
        return None
    
    best = results[0]
    reg_num = best.get("register_number")
    if not reg_num:
        return None
    
    return fetch_germany_data(reg_num)


def search_uk(search_term: str) -> dict:
    """Search UK ministerial meetings."""
    return search_uk_ministerial_meetings(search_term)


def search_austria(search_term: str) -> dict:
    """Search Austria lobbying register."""
    return search_austria_register(search_term)


def search_catalonia(search_term: str) -> dict:
    """Search Catalonia interest group register."""
    return search_catalonia_register(search_term)


def search_finland(search_term: str) -> dict:
    """Search Finland transparency register."""
    return search_finland_register(search_term)


def search_slovenia(search_term: str) -> dict:
    """Search Slovenia lobbying register."""
    return search_slovenia_register(search_term)


# Registry of all jurisdictions
JURISDICTIONS = {
    "eu": {
        "id": "eu",
        "name": "EU (European Commission)",
        "flag": "🇪🇺",
        "search_fn": search_eu,
        "has_financial_data": True,
        "note": "Via LobbyFacts.eu - includes Commission meetings",
        "default_enabled": True,
    },
    "france": {
        "id": "france", 
        "name": "France",
        "flag": "🇫🇷",
        "search_fn": search_france,
        "has_financial_data": True,
        "note": "Via HATVP - detailed activity disclosures",
        "default_enabled": True,
    },
    "germany": {
        "id": "germany",
        "name": "Germany", 
        "flag": "🇩🇪",
        "search_fn": search_germany,
        "has_financial_data": True,
        "note": "Via Bundestag Lobbyregister - cost ranges",
        "default_enabled": True,
    },
    "uk": {
        "id": "uk",
        "name": "United Kingdom",
        "flag": "🇬🇧", 
        "search_fn": search_uk,
        "has_financial_data": False,
        "note": "Ministerial meetings only - no expenditure tracking. Can be slow.",
        "default_enabled": False,
    },
    "austria": {
        "id": "austria",
        "name": "Austria",
        "flag": "🇦🇹",
        "search_fn": search_austria,
        "has_financial_data": False,
        "note": "Financial data only disclosed if >€100,000",
        "default_enabled": True,
    },
    "catalonia": {
        "id": "catalonia",
        "name": "Catalonia",
        "flag": "🏴󠁥󠁳󠁣󠁴󠁿",
        "search_fn": search_catalonia,
        "has_financial_data": True,
        "note": "Regional register - annual business volume",
        "default_enabled": True,
    },
    "finland": {
        "id": "finland",
        "name": "Finland",
        "flag": "🇫🇮",
        "search_fn": search_finland,
        "has_financial_data": False,
        "note": "Financial data available from July 2026",
        "default_enabled": True,
    },
    "slovenia": {
        "id": "slovenia",
        "name": "Slovenia",
        "flag": "🇸🇮",
        "search_fn": search_slovenia,
        "has_financial_data": False,
        "note": "Lists individual lobbyists, not companies. Search by lobbyist name or employer.",
        "default_enabled": True,
    },
}


def search_all(search_term: str, jurisdictions: list = None, skip_slow: bool = True) -> dict:
    """Search multiple jurisdictions."""
    if jurisdictions is None:
        jurisdictions = list(JURISDICTIONS.keys())
    
    results = {}
    for jur_id in jurisdictions:
        if jur_id not in JURISDICTIONS:
            continue
            
        jur = JURISDICTIONS[jur_id]
        
        if skip_slow and jur_id == "uk":
            results[jur_id] = None
            continue
        
        try:
            results[jur_id] = jur["search_fn"](search_term)
        except Exception as e:
            print(f"Error searching {jur['name']}: {e}")
            results[jur_id] = None
    
    return results
