"""
European Lobbying Tracker - Streamlit App
Uses the full detailed Excel export from the original script.
"""

import streamlit as st
import pandas as pd
from io import BytesIO
import time
import tempfile
import os

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
    create_excel_report,
)

# Import Boolean search helpers
from boolean_search import is_or_query, extract_or_terms, get_matching_term


# Jurisdiction config
JURISDICTIONS = {
    "eu": {
        "name": "EU (European Commission)",
        "flag": "🇪🇺",
        "note": "Via LobbyFacts.eu - includes Commission meetings",
        "default": True,
    },
    "france": {
        "name": "France",
        "flag": "🇫🇷",
        "note": "Via HATVP - detailed activity disclosures",
        "default": True,
    },
    "germany": {
        "name": "Germany", 
        "flag": "🇩🇪",
        "note": "Via Bundestag Lobbyregister - cost ranges",
        "default": True,
    },
    "uk": {
        "name": "UK (Ministers + Senior Officials)",
        "flag": "🇬🇧", 
        "note": "26,000+ meetings. Data most reliable from 2024 onwards.",
        "default": True,
    },
    "austria": {
        "name": "Austria",
        "flag": "🇦🇹",
        "note": "Financial data only if >€100,000",
        "default": True,
    },
    "catalonia": {
        "name": "Catalonia",
        "flag": "🏴󠁥󠁳󠁣󠁴󠁿",
        "note": "Regional register - annual business volume",
        "default": True,
    },
    "finland": {
        "name": "Finland",
        "flag": "🇫🇮",
        "note": "Financial data from July 2026",
        "default": True,
    },
    "slovenia": {
        "name": "Slovenia",
        "flag": "🇸🇮",
        "note": "Lists individual lobbyists - search by name/employer",
        "default": True,
    },
}


def run_search(search_term: str, selected: dict, progress_callback=None, uk_months_back=12):
    """Run searches and return data in format expected by create_excel_report.
    
    For Boolean OR queries (e.g. "shell OR bp"), fetches data for ALL matching
    entities and tags each with which search term it matched.
    """
    
    results = {
        "eu": None,
        "france": None, 
        "germany": None,
        "uk": None,
        "austria": None,
        "catalonia": None,
        "finland": None,
        "slovenia": None,
    }
    
    total = sum(selected.values())
    done = 0
    
    # Check if this is an OR query
    or_terms = extract_or_terms(search_term) if is_or_query(search_term) else [search_term]
    is_multi_entity = len(or_terms) > 1
    
    # EU - handle multiple entities for OR queries
    if selected.get("eu"):
        if progress_callback:
            progress_callback("🇪🇺 Searching EU register...", done/total)
        
        if is_multi_entity:
            # Fetch data for each OR term separately
            all_entities = []
            for term in or_terms:
                eu_matches = search_eu_register(term)
                if eu_matches:
                    eu_id = eu_matches[0].get("id")
                    if eu_id:
                        entity_data = fetch_eu_data(eu_id)
                        if entity_data:
                            entity_data["matched_term"] = term
                            entity_data["matched_name"] = eu_matches[0].get("name", term)
                            all_entities.append(entity_data)
            
            if all_entities:
                # Store as multiple_entities structure
                results["eu"] = {
                    "multiple_entities": all_entities,
                    "search_term": search_term,
                    "is_or_query": True
                }
        else:
            eu_matches = search_eu_register(search_term)
            if eu_matches:
                eu_id = eu_matches[0].get("id")
                if eu_id:
                    results["eu"] = fetch_eu_data(eu_id)
        done += 1
    
    # France - handle multiple entities for OR queries
    if selected.get("france"):
        if progress_callback:
            progress_callback("🇫🇷 Searching France (HATVP)...", done/total)
        
        if is_multi_entity:
            all_entities = []
            for term in or_terms:
                fr_matches = search_france_register(term)
                if fr_matches:
                    fr_id = fr_matches[0].get("id")
                    if fr_id:
                        entity_data = fetch_france_data(fr_id)
                        if entity_data:
                            entity_data["matched_term"] = term
                            entity_data["matched_name"] = fr_matches[0].get("name", term)
                            all_entities.append(entity_data)
            
            if all_entities:
                results["france"] = {
                    "multiple_entities": all_entities,
                    "search_term": search_term,
                    "is_or_query": True
                }
        else:
            fr_matches = search_france_register(search_term)
            if fr_matches:
                fr_id = fr_matches[0].get("id")
                if fr_id:
                    results["france"] = fetch_france_data(fr_id)
        done += 1
    
    # Germany - handle multiple entities for OR queries
    if selected.get("germany"):
        if progress_callback:
            progress_callback("🇩🇪 Searching Germany (Bundestag)...", done/total)
        
        if is_multi_entity:
            all_entities = []
            for term in or_terms:
                de_matches = search_germany_register(term)
                if de_matches:
                    reg_num = de_matches[0].get("register_number")
                    if reg_num:
                        entity_data = fetch_germany_data(reg_num)
                        if entity_data:
                            entity_data["matched_term"] = term
                            entity_data["matched_name"] = de_matches[0].get("name", term)
                            all_entities.append(entity_data)
            
            if all_entities:
                results["germany"] = {
                    "multiple_entities": all_entities,
                    "search_term": search_term,
                    "is_or_query": True
                }
        else:
            de_matches = search_germany_register(search_term)
            if de_matches:
                reg_num = de_matches[0].get("register_number")
                if reg_num:
                    results["germany"] = fetch_germany_data(reg_num)
        done += 1
    
    # UK - Uses pre-built index, already handles OR queries via boolean matching
    # Results already tagged with "organisation" field matching each term
    if selected.get("uk"):
        if progress_callback:
            progress_callback("🇬🇧 Searching UK meetings...", done/total)
        uk_result = search_uk_ministerial_meetings(search_term, months_back=uk_months_back)
        if uk_result and is_multi_entity:
            # Tag each meeting with which term it matched
            meetings = uk_result.get("meetings", [])
            for meeting in meetings:
                org = meeting.get("organisation", "")
                meeting["matched_term"] = get_matching_term(search_term, org)
            uk_result["is_or_query"] = True
        results["uk"] = uk_result
        done += 1
        done += 1
    
    # Austria
    if selected.get("austria"):
        if progress_callback:
            progress_callback("🇦🇹 Searching Austria...", done/total)
        results["austria"] = search_austria_register(search_term)
        done += 1
    
    # Catalonia
    if selected.get("catalonia"):
        if progress_callback:
            progress_callback("🏴󠁥󠁳󠁣󠁴󠁿 Searching Catalonia...", done/total)
        results["catalonia"] = search_catalonia_register(search_term)
        done += 1
    
    # Finland
    if selected.get("finland"):
        if progress_callback:
            progress_callback("🇫🇮 Searching Finland...", done/total)
        results["finland"] = search_finland_register(search_term)
        done += 1
    
    # Slovenia
    if selected.get("slovenia"):
        if progress_callback:
            progress_callback("🇸🇮 Searching Slovenia...", done/total)
        results["slovenia"] = search_slovenia_register(search_term)
        done += 1
    
    if progress_callback:
        progress_callback("✅ Complete!", 1.0)
    
    return results


def generate_full_excel(search_term: str, results: dict) -> BytesIO:
    """Generate the full detailed Excel report using the original function."""
    
    # Create temp file for the Excel
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name
    
    try:
        # Call the original comprehensive Excel creator
        create_excel_report(
            eu_data=results.get("eu"),
            fr_data=results.get("france"),
            de_data=results.get("germany"),
            ie_data=None,  # Ireland requires manual CSV
            uk_data=results.get("uk"),
            at_data=results.get("austria"),
            cat_data=results.get("catalonia"),
            fi_data=results.get("finland"),
            si_data=results.get("slovenia"),
            uk_officials_data=results.get("uk_officials"),
            output_path=tmp_path,
            org_name=search_term
        )
        
        # Read back into BytesIO
        with open(tmp_path, "rb") as f:
            buffer = BytesIO(f.read())
        buffer.seek(0)
        return buffer
    finally:
        # Clean up temp file
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


def display_summary(search_term: str, results: dict):
    """Display summary cards for each jurisdiction."""
    
    st.header(f"Results for: {search_term}")
    
    found = [k for k, v in results.items() if v is not None]
    not_found = [k for k, v in results.items() if v is None and k in JURISDICTIONS]
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Found in", f"{len(found)} registers")
    with col2:
        st.metric("Not found in", f"{len(not_found)} registers")
    
    st.markdown("---")
    
    # EU
    if results.get("eu"):
        data = results["eu"]
        with st.expander("🇪🇺 **EU (European Commission)** ✅", expanded=True):
            # Check if this is a multiple entities result (from OR query)
            if data.get("multiple_entities"):
                st.info(f"🔀 Found {len(data['multiple_entities'])} entities matching OR query")
                
                for entity in data["multiple_entities"]:
                    matched_term = entity.get("matched_term", "")
                    matched_name = entity.get("matched_name", matched_term)
                    
                    st.subheader(f"📌 {matched_name}")
                    st.caption(f"Matched term: \"{matched_term}\"")
                    
                    regs = entity.get("registrations", [])
                    meetings = entity.get("meetings", [])
                    latest = regs[-1] if regs else {}
                    
                    cols = st.columns(4)
                    with cols[0]:
                        min_c = int(latest.get('min', 0) or 0)
                        max_c = int(latest.get('max', 0) or 0)
                        if min_c or max_c:
                            st.metric("Lobbying Costs", f"€{min_c:,} - €{max_c:,}")
                        else:
                            st.metric("Lobbying Costs", "Not disclosed")
                    with cols[1]:
                        st.metric("Commission Meetings", len(meetings))
                    with cols[2]:
                        st.metric("Staff", f"{latest.get('members', 'N/A')} ({latest.get('members_fte', 'N/A')} FTE)")
                    with cols[3]:
                        st.metric("Data Snapshots", len(regs))
                    
                    st.caption(f"ID: {entity.get('org_id', 'N/A')} | HQ: {latest.get('head_country', 'N/A')}")
                    st.markdown("---")
            else:
                # Single entity result
                regs = data.get("registrations", [])
                meetings = data.get("meetings", [])
                latest = regs[-1] if regs else {}
                
                cols = st.columns(4)
                with cols[0]:
                    min_c = int(latest.get('min', 0) or 0)
                    max_c = int(latest.get('max', 0) or 0)
                    if min_c or max_c:
                        st.metric("Lobbying Costs", f"€{min_c:,} - €{max_c:,}")
                    else:
                        st.metric("Lobbying Costs", "Not disclosed")
                with cols[1]:
                    st.metric("Commission Meetings", len(meetings))
                with cols[2]:
                    st.metric("Staff", f"{latest.get('members', 'N/A')} ({latest.get('members_fte', 'N/A')} FTE)")
                with cols[3]:
                    st.metric("Data Snapshots", len(regs))
                
                st.caption(f"ID: {data.get('org_id', 'N/A')} | HQ: {latest.get('head_country', 'N/A')} | 📅 Data: {data.get('data_coverage', '2012-present')}")
    
    # France
    if results.get("france"):
        data = results["france"]
        with st.expander("🇫🇷 **France (HATVP)** ✅", expanded=True):
            # Check if this is a multiple entities result (from OR query)
            if data.get("multiple_entities"):
                st.info(f"🔀 Found {len(data['multiple_entities'])} entities matching OR query")
                
                for entity in data["multiple_entities"]:
                    matched_term = entity.get("matched_term", "")
                    matched_name = entity.get("matched_name", matched_term)
                    
                    st.subheader(f"📌 {matched_name}")
                    st.caption(f"Matched term: \"{matched_term}\"")
                    
                    info = entity.get("info", {})
                    exercises = entity.get("exercises", [])
                    activities = entity.get("activities", [])
                    latest = exercises[2] if len(exercises) > 2 else (exercises[0] if exercises else {})
                    
                    cols = st.columns(4)
                    with cols[0]:
                        st.metric("Lobbying Costs", latest.get("montant_depense", "Not disclosed"))
                    with cols[1]:
                        st.metric("Activities", len(activities))
                    with cols[2]:
                        st.metric("Staff", latest.get("nombre_salaries", "N/A"))
                    with cols[3]:
                        st.metric("Years of Data", len(exercises))
                    
                    st.caption(f"SIREN: {info.get('identifiant_national', 'N/A')} | City: {info.get('ville', 'N/A')}")
                    st.markdown("---")
            else:
                # Single entity result
                info = data.get("info", {})
                exercises = data.get("exercises", [])
                activities = data.get("activities", [])
                latest = exercises[2] if len(exercises) > 2 else (exercises[0] if exercises else {})
                
                cols = st.columns(4)
                with cols[0]:
                    st.metric("Lobbying Costs", latest.get("montant_depense", "Not disclosed"))
                with cols[1]:
                    st.metric("Activities", len(activities))
                with cols[2]:
                    st.metric("Staff", latest.get("nombre_salaries", "N/A"))
                with cols[3]:
                    st.metric("Years of Data", len(exercises))
                
                st.caption(f"SIREN: {info.get('identifiant_national', 'N/A')} | City: {info.get('ville', 'N/A')} | 📅 Data: {data.get('data_coverage', '2017-present')}")
    
    # Germany
    if results.get("germany"):
        data = results["germany"]
        with st.expander("🇩🇪 **Germany (Bundestag)** ✅", expanded=True):
            # Check if this is a multiple entities result (from OR query)
            if data.get("multiple_entities"):
                st.info(f"🔀 Found {len(data['multiple_entities'])} entities matching OR query")
                
                for entity in data["multiple_entities"]:
                    matched_term = entity.get("matched_term", "")
                    matched_name = entity.get("matched_name", matched_term)
                    
                    st.subheader(f"📌 {matched_name}")
                    st.caption(f"Matched term: \"{matched_term}\"")
                    
                    cols = st.columns(4)
                    with cols[0]:
                        min_e = entity.get('expenses_min', 0)
                        max_e = entity.get('expenses_max', 0)
                        if min_e or max_e:
                            st.metric("Lobbying Costs", f"€{min_e:,} - €{max_e:,}")
                        else:
                            st.metric("Lobbying Costs", "Not disclosed")
                    with cols[1]:
                        st.metric("Staff (FTE)", entity.get("employee_fte", "N/A"))
                    with cols[2]:
                        st.metric("Regulatory Projects", len(entity.get("legislative_projects", [])))
                    with cols[3]:
                        st.metric("Fields of Interest", len(entity.get("fields_of_interest", [])))
                    
                    st.caption(f"Reg: {entity.get('register_number', 'N/A')} | Berlin Office: {'Yes' if entity.get('berlin_office') else 'No'}")
                    st.markdown("---")
            else:
                # Single entity result
                cols = st.columns(4)
                with cols[0]:
                    min_e = data.get('expenses_min', 0)
                    max_e = data.get('expenses_max', 0)
                    if min_e or max_e:
                        st.metric("Lobbying Costs", f"€{min_e:,} - €{max_e:,}")
                    else:
                        st.metric("Lobbying Costs", "Not disclosed")
                with cols[1]:
                    st.metric("Staff (FTE)", data.get("employee_fte", "N/A"))
                with cols[2]:
                    st.metric("Regulatory Projects", len(data.get("legislative_projects", [])))
                with cols[3]:
                    st.metric("Fields of Interest", len(data.get("fields_of_interest", [])))
                
                st.caption(f"Reg: {data.get('register_number', 'N/A')} | Berlin Office: {'Yes' if data.get('berlin_office') else 'No'} | 📅 Data: {data.get('data_coverage', '2022-present')}")
    
    # UK Ministerial
    if results.get("uk"):
        data = results["uk"]
        with st.expander("🇬🇧 **UK Ministers** ✅", expanded=True):
            meetings = data.get("meetings", [])
            
            # Check if this is an OR query with tagged meetings
            if data.get("is_or_query") and meetings:
                # Group meetings by matched term
                by_term = {}
                for m in meetings:
                    term = m.get("matched_term", "Unknown")
                    if term not in by_term:
                        by_term[term] = []
                    by_term[term].append(m)
                
                st.info(f"🔀 Found meetings for {len(by_term)} organisations")
                
                cols = st.columns(3)
                with cols[0]:
                    st.metric("Total Meetings", len(meetings))
                with cols[1]:
                    st.metric("Departments Searched", len(data.get("departments_searched", [])))
                with cols[2]:
                    st.metric("Unique Ministers", len(data.get("by_minister", {})))
                
                st.markdown("**Meetings by matched organisation:**")
                for term, term_meetings in sorted(by_term.items(), key=lambda x: -len(x[1])):
                    st.write(f"• **{term}**: {len(term_meetings)} meetings")
            else:
                # Single term result
                cols = st.columns(3)
                with cols[0]:
                    st.metric("Ministerial Meetings", len(meetings))
                with cols[1]:
                    st.metric("Departments Searched", len(data.get("departments_searched", [])))
                with cols[2]:
                    by_minister = data.get("by_minister", {})
                    st.metric("Unique Ministers", len(by_minister))
            
            st.caption(f"📅 Data coverage: {data.get('data_coverage', '2024-present')}")
    
    # UK Senior Officials
    if results.get("uk_officials"):
        data = results["uk_officials"]
        with st.expander("🇬🇧 **UK Senior Officials** ✅", expanded=True):
            meetings = data.get("meetings", [])
            cols = st.columns(3)
            with cols[0]:
                st.metric("Meetings", len(meetings))
            with cols[1]:
                st.metric("Departments", len(data.get("by_department", {})))
            with cols[2]:
                st.metric("Unique Officials", len(data.get("by_official", {})))
            st.caption(f"📅 Data coverage: {data.get('data_coverage', 'Last year')}")
    
    # Austria
    if results.get("austria"):
        data = results["austria"]
        with st.expander("🇦🇹 **Austria** ✅", expanded=True):
            cols = st.columns(3)
            with cols[0]:
                st.metric("Register Entries", data.get("entry_count", 0))
            with cols[1]:
                st.metric("Categories", len(data.get("by_category", {})))
            with cols[2]:
                st.metric("Financial Data", "If >€100k")
            st.caption(f"📅 Data coverage: {data.get('data_coverage', '2013-present')}")
    
    # Catalonia
    if results.get("catalonia"):
        data = results["catalonia"]
        with st.expander("🏴󠁥󠁳󠁣󠁴󠁿 **Catalonia** ✅", expanded=True):
            cols = st.columns(3)
            with cols[0]:
                st.metric("Register Entries", data.get("entry_count", 0))
            with cols[1]:
                st.metric("Annual Volume", data.get("total_volume_formatted", "N/A"))
            with cols[2]:
                st.metric("Categories", len(data.get("by_category", {})))
            st.caption(f"📅 Data coverage: {data.get('data_coverage', '2016-present')}")
    
    # Finland
    if results.get("finland"):
        data = results["finland"]
        with st.expander("🇫🇮 **Finland** ✅", expanded=True):
            cols = st.columns(3)
            with cols[0]:
                st.metric("Register Entries", data.get("entry_count", 0))
            with cols[1]:
                st.metric("Activity Disclosures", data.get("total_activities", 0))
            with cols[2]:
                st.metric("Financial Data", "From July 2026")
            
            st.caption(f"📅 Data coverage: {data.get('data_coverage', '2024-present')}")
            
            # Show topics if available
            entries = data.get("entries", [])
            if entries and entries[0].get("topics"):
                st.write("**Topics:**", ", ".join(entries[0]["topics"][:5]))
    
    # Slovenia
    if results.get("slovenia"):
        data = results["slovenia"]
        with st.expander("🇸🇮 **Slovenia** ✅", expanded=True):
            cols = st.columns(3)
            with cols[0]:
                st.metric("Lobbyists Found", data.get("entry_count", 0))
            with cols[1]:
                st.metric("Total Registered", data.get("total_registered", 0))
            with cols[2]:
                top_fields = data.get("top_fields", [])
                st.metric("Top Field", top_fields[0][0] if top_fields else "N/A")
            
            # Show matched lobbyists
            entries = data.get("entries", [])
            if entries:
                st.write("**Matched Lobbyists:**")
                for e in entries[:3]:
                    company = f" ({e['company']})" if e.get('company') else ""
                    st.write(f"• {e['name']}{company}")
            
            st.caption(f"⚠️ Slovenia lists individual lobbyists, not companies | 📅 Data: {data.get('data_coverage', '2010-present')}")
    
    # Not found
    for jur_id in not_found:
        if jur_id in JURISDICTIONS:
            jur = JURISDICTIONS[jur_id]
            with st.expander(f"{jur['flag']} **{jur['name']}** ❌"):
                st.info("No matches found")


# =============================================================================
# MATCH PREVIEW FUNCTIONS
# =============================================================================

def preview_matches(search_term: str, selected: dict, progress_callback=None):
    """First stage: search all registers and return matches without fetching full data."""
    
    matches = {
        "eu": [],
        "france": [],
        "germany": [],
        "uk": None,  # UK uses index, returns full results directly
        "austria": None,
        "catalonia": None,
        "finland": None,
        "slovenia": None,
    }
    
    total = sum(selected.values())
    done = 0
    
    # EU - just get matches, don't fetch full data
    if selected.get("eu"):
        if progress_callback:
            progress_callback("🇪🇺 Searching EU register...", done/total)
        matches["eu"] = search_eu_register(search_term) or []
        done += 1
    
    # France - just get matches
    if selected.get("france"):
        if progress_callback:
            progress_callback("🇫🇷 Searching France (HATVP)...", done/total)
        matches["france"] = search_france_register(search_term) or []
        done += 1
    
    # Germany - just get matches
    if selected.get("germany"):
        if progress_callback:
            progress_callback("🇩🇪 Searching Germany (Bundestag)...", done/total)
        matches["germany"] = search_germany_register(search_term) or []
        done += 1
    
    # UK - returns full data from index (fast, no second stage needed)
    if selected.get("uk"):
        if progress_callback:
            progress_callback("🇬🇧 Searching UK meetings...", done/total)
        matches["uk"] = search_uk_ministerial_meetings(search_term, months_back=12)
        done += 1
    
    # Austria - returns full data
    if selected.get("austria"):
        if progress_callback:
            progress_callback("🇦🇹 Searching Austria...", done/total)
        matches["austria"] = search_austria_register(search_term)
        done += 1
    
    # Catalonia - returns full data
    if selected.get("catalonia"):
        if progress_callback:
            progress_callback("🏴󠁥󠁳󠁣󠁴󠁿 Searching Catalonia...", done/total)
        matches["catalonia"] = search_catalonia_register(search_term)
        done += 1
    
    # Finland - returns full data
    if selected.get("finland"):
        if progress_callback:
            progress_callback("🇫🇮 Searching Finland...", done/total)
        matches["finland"] = search_finland_register(search_term)
        done += 1
    
    # Slovenia - returns full data
    if selected.get("slovenia"):
        if progress_callback:
            progress_callback("🇸🇮 Searching Slovenia...", done/total)
        matches["slovenia"] = search_slovenia_register(search_term)
        done += 1
    
    if progress_callback:
        progress_callback("✅ Search complete!", 1.0)
    
    return matches


def fetch_selected_data(selections: dict, other_results: dict, progress_callback=None):
    """Second stage: fetch full data only for user-selected matches."""
    
    results = {
        "eu": None,
        "france": None,
        "germany": None,
        "uk": other_results.get("uk"),  # Already have full data
        "austria": other_results.get("austria"),
        "catalonia": other_results.get("catalonia"),
        "finland": other_results.get("finland"),
        "slovenia": other_results.get("slovenia"),
    }
    
    total = len(selections.get("eu", [])) + len(selections.get("france", [])) + len(selections.get("germany", []))
    if total == 0:
        return results
    
    done = 0
    
    # EU - fetch selected
    eu_selections = selections.get("eu", [])
    if eu_selections:
        if len(eu_selections) == 1:
            # Single selection
            if progress_callback:
                progress_callback(f"🇪🇺 Fetching {eu_selections[0]['name']}...", done/total)
            results["eu"] = fetch_eu_data(eu_selections[0]["id"])
            done += 1
        else:
            # Multiple selections - use multiple_entities format
            all_entities = []
            for sel in eu_selections:
                if progress_callback:
                    progress_callback(f"🇪🇺 Fetching {sel['name']}...", done/total)
                entity_data = fetch_eu_data(sel["id"])
                if entity_data:
                    entity_data["matched_term"] = sel["name"]
                    entity_data["matched_name"] = sel["name"]
                    all_entities.append(entity_data)
                done += 1
            
            if all_entities:
                results["eu"] = {
                    "multiple_entities": all_entities,
                    "search_term": ", ".join(s["name"] for s in eu_selections),
                    "is_or_query": True
                }
    
    # France - fetch selected
    fr_selections = selections.get("france", [])
    if fr_selections:
        if len(fr_selections) == 1:
            if progress_callback:
                progress_callback(f"🇫🇷 Fetching {fr_selections[0]['name']}...", done/total)
            results["france"] = fetch_france_data(fr_selections[0]["id"])
            done += 1
        else:
            all_entities = []
            for sel in fr_selections:
                if progress_callback:
                    progress_callback(f"🇫🇷 Fetching {sel['name']}...", done/total)
                entity_data = fetch_france_data(sel["id"])
                if entity_data:
                    entity_data["matched_term"] = sel["name"]
                    entity_data["matched_name"] = sel["name"]
                    all_entities.append(entity_data)
                done += 1
            
            if all_entities:
                results["france"] = {
                    "multiple_entities": all_entities,
                    "search_term": ", ".join(s["name"] for s in fr_selections),
                    "is_or_query": True
                }
    
    # Germany - fetch selected
    de_selections = selections.get("germany", [])
    if de_selections:
        if len(de_selections) == 1:
            if progress_callback:
                progress_callback(f"🇩🇪 Fetching {de_selections[0]['name']}...", done/total)
            results["germany"] = fetch_germany_data(de_selections[0]["register_number"])
            done += 1
        else:
            all_entities = []
            for sel in de_selections:
                if progress_callback:
                    progress_callback(f"🇩🇪 Fetching {sel['name']}...", done/total)
                entity_data = fetch_germany_data(sel["register_number"])
                if entity_data:
                    entity_data["matched_term"] = sel["name"]
                    entity_data["matched_name"] = sel["name"]
                    all_entities.append(entity_data)
                done += 1
            
            if all_entities:
                results["germany"] = {
                    "multiple_entities": all_entities,
                    "search_term": ", ".join(s["name"] for s in de_selections),
                    "is_or_query": True
                }
    
    if progress_callback:
        progress_callback("✅ Data fetch complete!", 1.0)
    
    return results


# =============================================================================
# STREAMLIT APP
# =============================================================================

st.set_page_config(
    page_title="European Lobbying Tracker",
    page_icon="🏛️",
    layout="wide"
)

st.title("🏛️ European Lobbying Tracker")
st.markdown("Search corporate lobbying records across European transparency registers and download comprehensive Excel reports.")

# Sidebar
st.sidebar.header("🌍 Jurisdictions")

selected = {}
for jur_id, jur in JURISDICTIONS.items():
    selected[jur_id] = st.sidebar.checkbox(
        f"{jur['flag']} {jur['name']}", 
        value=jur["default"],
        help=jur["note"]
    )

st.sidebar.markdown(f"**{sum(selected.values())}** of {len(JURISDICTIONS)} selected")

st.sidebar.header("⚙️ Options")

# Date filter for UK data
uk_date_filter = st.sidebar.selectbox(
    "🇬🇧 UK data range",
    options=["Last 12 months", "Last 6 months", "Last 3 months", "All available"],
    index=0,
    help="Filter UK meetings by date. Data is most reliable from 2024 onwards."
)

st.sidebar.caption("Other jurisdictions search all available data")

# Main search
col1, col2 = st.columns([3, 1])
with col1:
    search_term = st.text_input(
        "🔍 Company name", 
        placeholder="e.g. Google, OpenAI, Meta...",
        help="Supports Boolean search: AND, OR, NOT, quotes, parentheses"
    )
with col2:
    st.markdown("<br>", unsafe_allow_html=True)
    search_button = st.button("Search", type="primary", use_container_width=True)

# Boolean search help expander
with st.expander("🔎 Advanced Search Syntax"):
    st.markdown("""
    **Boolean operators** for precise searches:
    
    | Syntax | Example | Meaning |
    |--------|---------|---------|
    | `AND` | `shell AND bp` | Both terms must appear |
    | `OR` | `shell OR bp` | Either term matches |
    | `NOT` | `shell NOT gas` | Excludes matches |
    | `"quotes"` | `"big oil"` | Exact phrase match |
    | `(parens)` | `(shell OR bp) AND energy` | Grouping |
    
    **Examples:**
    - `palantir OR anduril` — Defense tech companies
    - `meta NOT facebook` — Meta but not containing "facebook"
    - `(google OR microsoft) AND ai` — Either company with "ai" in name
    - `"consulting group"` — Exact phrase match
    """)

# Initialize session state for two-stage search
if "matches" not in st.session_state:
    st.session_state.matches = None
if "search_term_used" not in st.session_state:
    st.session_state.search_term_used = None
if "final_results" not in st.session_state:
    st.session_state.final_results = None

# Run search - Stage 1: Find matches
if search_button and search_term:
    if not any(selected.values()):
        st.warning("Please select at least one jurisdiction.")
    else:
        st.markdown("---")
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        def update_progress(msg, pct):
            status_text.text(msg)
            progress_bar.progress(pct)
        
        # Stage 1: Get matches
        st.session_state.matches = preview_matches(search_term, selected, update_progress)
        st.session_state.search_term_used = search_term
        st.session_state.final_results = None  # Reset final results
        st.session_state.selected_jurisdictions = selected.copy()
        
        time.sleep(0.3)
        status_text.empty()
        progress_bar.empty()

# Show match selection UI if we have matches
if st.session_state.matches and st.session_state.search_term_used:
    matches = st.session_state.matches
    search_term_display = st.session_state.search_term_used
    
    st.header(f"🔍 Matches for: {search_term_display}")
    
    # Check if there are any matches requiring selection (EU, France, Germany)
    has_selectable = (
        len(matches.get("eu", [])) > 0 or 
        len(matches.get("france", [])) > 0 or 
        len(matches.get("germany", [])) > 0
    )
    
    if has_selectable:
        st.info("📋 **Select the organisations you want to include in the report.** For EU, France, and Germany, we found multiple possible matches - tick the ones you want.")
    
    # Create selection UI
    user_selections = {"eu": [], "france": [], "germany": []}
    
    # EU matches
    if matches.get("eu"):
        with st.expander(f"🇪🇺 **EU Register** - {len(matches['eu'])} matches found", expanded=True):
            for i, match in enumerate(matches["eu"][:15]):  # Limit to top 15
                col1, col2 = st.columns([1, 4])
                with col1:
                    checked = st.checkbox(
                        "Select",
                        value=(i == 0),  # Default select first match
                        key=f"eu_{match['id']}"
                    )
                with col2:
                    country = match.get('country', '')
                    st.markdown(f"**{match['name']}** {f'({country})' if country else ''}")
                
                if checked:
                    user_selections["eu"].append(match)
    
    # France matches
    if matches.get("france"):
        with st.expander(f"🇫🇷 **France (HATVP)** - {len(matches['france'])} matches found", expanded=True):
            for i, match in enumerate(matches["france"][:15]):
                col1, col2 = st.columns([1, 4])
                with col1:
                    checked = st.checkbox(
                        "Select",
                        value=(i == 0),
                        key=f"fr_{match['id']}"
                    )
                with col2:
                    city = match.get('city', '')
                    st.markdown(f"**{match['name']}** {f'({city})' if city else ''}")
                
                if checked:
                    user_selections["france"].append(match)
    
    # Germany matches
    if matches.get("germany"):
        with st.expander(f"🇩🇪 **Germany (Bundestag)** - {len(matches['germany'])} matches found", expanded=True):
            for i, match in enumerate(matches["germany"][:15]):
                col1, col2 = st.columns([1, 4])
                with col1:
                    checked = st.checkbox(
                        "Select",
                        value=(i == 0),
                        key=f"de_{match['register_number']}"
                    )
                with col2:
                    city = match.get('city', '')
                    st.markdown(f"**{match['name']}** {f'({city})' if city else ''}")
                
                if checked:
                    user_selections["germany"].append(match)
    
    # Show other jurisdictions that already have full data
    other_found = []
    if matches.get("uk") and matches["uk"].get("meetings"):
        other_found.append(f"🇬🇧 UK: {len(matches['uk']['meetings'])} meetings")
    if matches.get("austria") and matches["austria"].get("entries"):
        other_found.append(f"🇦🇹 Austria: {matches['austria']['entry_count']} entries")
    if matches.get("catalonia") and matches["catalonia"].get("entries"):
        other_found.append(f"🏴󠁥󠁳󠁣󠁴󠁿 Catalonia: {matches['catalonia']['entry_count']} entries")
    if matches.get("finland") and matches["finland"].get("entries"):
        other_found.append(f"🇫🇮 Finland: {matches['finland']['entry_count']} entries")
    if matches.get("slovenia") and matches["slovenia"].get("entries"):
        other_found.append(f"🇸🇮 Slovenia: {matches['slovenia']['entry_count']} lobbyists")
    
    if other_found:
        st.markdown("**Also found:**")
        st.markdown(" • ".join(other_found))
    
    # Fetch button
    st.markdown("---")
    total_selected = len(user_selections["eu"]) + len(user_selections["france"]) + len(user_selections["germany"])
    
    fetch_button = st.button(
        f"📥 Fetch Data for {total_selected} Selected Organisation(s)",
        type="primary",
        disabled=(total_selected == 0 and not other_found)
    )
    
    if fetch_button:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        def update_progress(msg, pct):
            status_text.text(msg)
            progress_bar.progress(pct)
        
        # Stage 2: Fetch full data for selections
        st.session_state.final_results = fetch_selected_data(
            user_selections, 
            matches,  # Pass through UK, Austria, etc.
            update_progress
        )
        
        time.sleep(0.3)
        status_text.empty()
        progress_bar.empty()
        
        st.rerun()

# Show final results and download
if st.session_state.final_results:
    results = st.session_state.final_results
    search_term_display = st.session_state.search_term_used
    
    # Show summary
    display_summary(search_term_display, results)
    
    # Export - THE FULL DETAILED REPORT
    st.markdown("---")
    st.header("📥 Download Full Report")
    st.markdown("Download the comprehensive Excel report with **all details** - meetings, activities, financial history, and more.")
    
    with st.spinner("Generating detailed Excel report..."):
        excel_buffer = generate_full_excel(search_term_display, results)
    
    st.download_button(
        label="📊 Download Comprehensive Excel Report",
        data=excel_buffer,
        file_name=f"{search_term_display.lower().replace(' ', '_')}_lobbying_full.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
    
    st.caption("The Excel file contains separate sheets for each jurisdiction with full details: all meetings, activities, financial history, fields of interest, and more.")
    
    # New search button
    if st.button("🔄 New Search"):
        st.session_state.matches = None
        st.session_state.search_term_used = None
        st.session_state.final_results = None
        st.rerun()

# Footer
st.markdown("---")
st.caption("Data: LobbyFacts.eu • HATVP • Bundestag • GOV.UK • lobbyreg.justiz.gv.at • transparenciacatalunya.cat • avoimuusrekisteri.fi")
