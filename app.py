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
    find_france_intermediaries,
    find_eu_intermediaries,
    search_germany_register,
    fetch_germany_data,
    search_uk_ministerial_meetings,
    search_austria_register,
    search_catalonia_register,
    search_finland_register,
    search_slovenia_register,
    search_ireland_register,
    fetch_ireland_data,
    search_ireland_lobbying,
    search_netherlands_register,
    fetch_netherlands_data,
    search_netherlands_agendas,
    search_ec_meetings_by_representative,
    search_ec_meetings_by_organisation,
    search_ec_meetings_by_topic,
    search_ec_meetings_by_cabinet,
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
    "ireland": {
        "name": "Ireland",
        "flag": "🇮🇪",
        "note": "Lobbying.ie - mandatory register since 2015. ~2,800 lobbyists.",
        "default": True,
    },
    "netherlands": {
        "name": "Netherlands",
        "flag": "🇳🇱",
        "note": "Ministerial agendas via openlobby.nl. Voluntary, data may be incomplete.",
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
        "ireland": None,
        "netherlands": None,
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
    
    # Ireland - Uses pre-built index
    if selected.get("ireland"):
        if progress_callback:
            progress_callback("🇮🇪 Searching Ireland lobbying register...", done/total)
        results["ireland"] = search_ireland_lobbying(search_term)
        done += 1
    
    # Netherlands - Uses pre-built index
    if selected.get("netherlands"):
        if progress_callback:
            progress_callback("🇳🇱 Searching Netherlands agendas...", done/total)
        results["netherlands"] = search_netherlands_agendas(search_term)
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
            ie_data=results.get("ireland"),
            uk_data=results.get("uk"),
            at_data=results.get("austria"),
            cat_data=results.get("catalonia"),
            fi_data=results.get("finland"),
            si_data=results.get("slovenia"),
            uk_officials_data=results.get("uk_officials"),
            nl_data=results.get("netherlands"),
            ec_data=results.get("ec_meetings"),
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
    
    # Check if this is a minister or topic mode search
    uk_data = results.get("uk")
    is_minister_mode = uk_data and uk_data.get("search_field") == "minister"
    is_topic_mode = uk_data and uk_data.get("search_field") == "topic"
    is_department_mode = uk_data and uk_data.get("search_field") == "department"
    not_found = []
    
    if is_minister_mode:
        st.header(f"Minister/Official: {search_term}")
        col1, col2 = st.columns(2)
        with col1:
            if uk_data and uk_data.get("meetings"):
                st.metric("UK Meetings", len(uk_data["meetings"]))
            else:
                st.metric("UK Meetings", 0)
        with col2:
            ec_data = results.get("ec_meetings")
            if ec_data and ec_data.get("meetings"):
                st.metric("EC Meetings", len(ec_data["meetings"]))
            else:
                st.metric("EC Meetings", 0)
    elif is_topic_mode:
        st.header(f"Topic: {search_term}")
        col1, col2 = st.columns(2)
        uk_topic_count = len(uk_data["meetings"]) if uk_data and uk_data.get("meetings") else 0
        ec_data = results.get("ec_meetings")
        ec_topic_count = len(ec_data["meetings"]) if ec_data and ec_data.get("meetings") else 0
        with col1:
            st.metric("UK Meetings", uk_topic_count)
        with col2:
            st.metric("EC Meetings", ec_topic_count)
        st.caption(f"Total: {uk_topic_count + ec_topic_count} meetings mentioning '{search_term}'")
    elif is_department_mode:
        st.header(f"Department/DG: {search_term}")
        col1, col2 = st.columns(2)
        uk_dept_count = len(uk_data["meetings"]) if uk_data and uk_data.get("meetings") else 0
        ec_data = results.get("ec_meetings")
        ec_dept_count = len(ec_data["meetings"]) if ec_data and ec_data.get("meetings") else 0
        with col1:
            st.metric("UK Meetings", uk_dept_count)
        with col2:
            st.metric("EC Meetings", ec_dept_count)
    else:
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
        is_minister_mode = data.get("search_field") == "minister"
        is_uk_topic_mode = data.get("search_field") == "topic"
        is_uk_dept_mode = data.get("search_field") == "department"
        
        if is_minister_mode:
            # MINISTER SEARCH MODE - show who this minister/official met
            with st.expander("🇬🇧 **UK Minister/Official Meetings** ✅", expanded=True):
                meetings = data.get("meetings", [])
                by_org = data.get("by_organisation", {})
                by_dept = data.get("by_department", {})
                by_minister = data.get("by_minister", {})
                
                cols = st.columns(4)
                with cols[0]:
                    st.metric("Total Meetings", len(meetings))
                with cols[1]:
                    st.metric("Unique Organisations Met", len(by_org))
                with cols[2]:
                    st.metric("Ministers/Officials Matched", len(by_minister))
                with cols[3]:
                    st.metric("Departments", len(by_dept))
                
                # Show which ministers/officials matched the search
                if by_minister:
                    st.markdown("**Matched ministers/officials:**")
                    for minister, count in list(by_minister.items())[:10]:
                        st.write(f"• **{minister}** ({count} meetings)")
                
                # Show top organisations met - the key insight for this search mode
                if by_org:
                    st.markdown("**Top organisations met:**")
                    for org, count in list(by_org.items())[:15]:
                        st.write(f"• {org}: {count} meetings")
                    if len(by_org) > 15:
                        st.caption(f"...and {len(by_org) - 15} more organisations. Download the Excel report for the full list.")
                
                st.caption(f"📅 Data coverage: {data.get('data_coverage', '2024-present')}")
        elif is_uk_topic_mode:
            # TOPIC SEARCH MODE - show who met about this topic
            with st.expander("🇬🇧 **UK Meetings on Topic** ✅", expanded=True):
                meetings = data.get("meetings", [])
                by_org = data.get("by_organisation", {})
                by_dept = data.get("by_department", {})
                by_minister = data.get("by_minister", {})
                
                cols = st.columns(4)
                with cols[0]:
                    st.metric("Meetings on Topic", len(meetings))
                with cols[1]:
                    st.metric("Organisations", len(by_org))
                with cols[2]:
                    st.metric("Ministers/Officials", len(by_minister))
                with cols[3]:
                    st.metric("Departments", len(by_dept))
                
                # Top organisations lobbying on this topic
                if by_org:
                    st.markdown("**Top organisations meeting on this topic:**")
                    for org, count in list(by_org.items())[:15]:
                        st.write(f"• {org}: {count} meetings")
                    if len(by_org) > 15:
                        st.caption(f"...and {len(by_org) - 15} more.")
                
                # Which ministers took these meetings
                if by_minister:
                    st.markdown("**Ministers/officials who took meetings:**")
                    for minister, count in list(by_minister.items())[:10]:
                        st.write(f"• **{minister}** ({count} meetings)")
                
                st.caption(f"📅 Data coverage: {data.get('data_coverage', '2024-present')}")
        elif is_uk_dept_mode:
            # DEPARTMENT SEARCH MODE - show what this department discussed
            with st.expander("🇬🇧 **UK Department Meetings** ✅", expanded=True):
                meetings = data.get("meetings", [])
                by_org = data.get("by_organisation", {})
                by_dept = data.get("by_department", {})
                by_minister = data.get("by_minister", {})
                
                cols = st.columns(4)
                with cols[0]:
                    st.metric("Meetings", len(meetings))
                with cols[1]:
                    st.metric("Organisations", len(by_org))
                with cols[2]:
                    st.metric("Ministers/Officials", len(by_minister))
                with cols[3]:
                    st.metric("Departments Matched", len(by_dept))
                
                if by_dept:
                    st.markdown("**Matched departments:**")
                    for dept, count in list(by_dept.items())[:10]:
                        st.write(f"• **{dept}** ({count} meetings)")
                
                if by_org:
                    st.markdown("**Top organisations met:**")
                    for org, count in list(by_org.items())[:15]:
                        st.write(f"• {org}: {count} meetings")
                    if len(by_org) > 15:
                        st.caption(f"...and {len(by_org) - 15} more.")
                
                if by_minister:
                    st.markdown("**Ministers/officials who took meetings:**")
                    for minister, count in list(by_minister.items())[:10]:
                        st.write(f"• **{minister}** ({count} meetings)")
                
                st.caption(f"📅 Data coverage: {data.get('data_coverage', '2024-present')}")
        else:
            # ORGANISATION SEARCH MODE (default) 
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
    
    # EC Meetings (minister/representative mode or topic mode)
    if results.get("ec_meetings"):
        data = results["ec_meetings"]
        ec_is_topic = data.get("search_field") == "topic"
        ec_is_dept = data.get("search_field") == "department"
        
        if ec_is_topic:
            expander_title = "🇪🇺 **EC Meetings on Topic** ✅"
        elif ec_is_dept:
            expander_title = "🇪🇺 **EC Cabinet/DG Meetings** ✅"
        else:
            expander_title = "🇪🇺 **EC Commissioner/Official Meetings** ✅"
        with st.expander(expander_title, expanded=True):
            meetings = data.get("meetings", [])
            by_org = data.get("by_organisation", {})
            by_rep = data.get("by_representative", {})
            by_cabinet = data.get("by_cabinet", {})
            
            cols = st.columns(4)
            with cols[0]:
                st.metric("Meetings" if (ec_is_topic or ec_is_dept) else "Total EC Meetings", len(meetings))
            with cols[1]:
                st.metric("Organisations", len(by_org))
            with cols[2]:
                st.metric("EC Representatives", len(by_rep))
            with cols[3]:
                st.metric("Cabinets/DGs", len(by_cabinet))
            
            if ec_is_topic or ec_is_dept:
                # Topic/Department mode: show who's lobbying
                if by_org:
                    st.markdown("**Top organisations meeting on this topic:**" if ec_is_topic else "**Top organisations met:**")
                    for org, count in list(by_org.items())[:15]:
                        st.write(f"• {org}: {count} meetings")
                    if len(by_org) > 15:
                        st.caption(f"...and {len(by_org) - 15} more.")
                
                if ec_is_dept and by_cabinet:
                    st.markdown("**Matched Cabinets/DGs:**")
                    for cab, count in list(by_cabinet.items())[:10]:
                        st.write(f"• **{cab}** ({count} meetings)")
                
                if by_rep:
                    st.markdown("**EC representatives who took meetings:**")
                    for rep, count in list(by_rep.items())[:10]:
                        st.write(f"• **{rep}** ({count} meetings)")
            else:
                # Minister mode: show matched representatives and who they met
                if by_rep:
                    st.markdown("**Matched EC representatives:**")
                    for rep, count in list(by_rep.items())[:10]:
                        st.write(f"• **{rep}** ({count} meetings)")
                
                if by_org:
                    st.markdown("**Top organisations met:**")
                    for org, count in list(by_org.items())[:15]:
                        st.write(f"• {org}: {count} meetings")
                    if len(by_org) > 15:
                        st.caption(f"...and {len(by_org) - 15} more. Download the report for the full list.")
            
            st.caption(f"📅 Data coverage: {data.get('data_coverage', '2014-present')} | Source: EC Open Data Portal")
    
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
    
    # Ireland
    if results.get("ireland"):
        data = results["ireland"]
        with st.expander("🇮🇪 **Ireland (Lobbying.ie)** ✅", expanded=True):
            returns = data.get("returns", [])
            cols = st.columns(3)
            with cols[0]:
                st.metric("Lobbying Returns", data.get("return_count", len(returns)))
            with cols[1]:
                st.metric("Lobbyists", len(data.get("by_lobbyist", {})))
            with cols[2]:
                st.metric("Public Bodies", len(data.get("by_official_body", {})))
            
            # Show top public bodies contacted
            top_bodies = list(data.get("by_official_body", {}).items())[:3]
            if top_bodies:
                st.write("**Top public bodies contacted:**")
                for body, count in top_bodies:
                    st.write(f"• {body}: {count}")
            
            st.caption(f"📅 Data coverage: {data.get('data_coverage', '2015-present')}")
    
    # Netherlands
    if results.get("netherlands"):
        data = results["netherlands"]
        with st.expander("🇳🇱 **Netherlands (Ministerial Agendas)** ✅", expanded=True):
            appointments = data.get("appointments", [])
            cols = st.columns(3)
            with cols[0]:
                st.metric("Agenda Appointments", data.get("appointment_count", len(appointments)))
            with cols[1]:
                st.metric("Ministers", len(data.get("by_minister", {})))
            with cols[2]:
                st.metric("Ministries", len(data.get("by_ministry", {})))
            
            # Show top ministers
            top_ministers = list(data.get("by_minister", {}).items())[:3]
            if top_ministers:
                st.write("**Top ministers:**")
                for minister, count in top_ministers:
                    st.write(f"• {minister}: {count} appointments")
            
            st.caption(f"⚠️ Data is voluntary and may be incomplete | 📅 {data.get('data_coverage', '2023-present')}")
    
    # Not found
    for jur_id in not_found:
        if jur_id in JURISDICTIONS:
            jur = JURISDICTIONS[jur_id]
            with st.expander(f"{jur['flag']} **{jur['name']}** ❌"):
                st.info("No matches found")


# =============================================================================
# MATCH PREVIEW FUNCTIONS
# =============================================================================

def preview_matches(search_term: str, selected: dict, progress_callback=None, uk_months_back=None, search_field="organisation"):
    """First stage: search all registers and return matches without fetching full data."""
    
    matches = {
        "eu": [],
        "france": [],
        "germany": [],
        "uk": None,  # UK uses index, returns full results directly
        "ec_meetings": None,  # EC meetings by representative (minister mode)
        "ireland": None,
        "netherlands": None,
        "austria": None,
        "catalonia": None,
        "finland": None,
        "slovenia": None,
    }
    
    total = sum(selected.values())
    done = 0
    
    # In minister search mode, only search UK meetings and EC meetings
    # In topic search mode, only search UK meetings and EC meetings (by subject)
    minister_mode = search_field == "minister"
    topic_mode = search_field == "topic"
    department_mode = search_field == "department"
    meetings_only_mode = minister_mode or topic_mode or department_mode
    
    # EU - just get matches, don't fetch full data
    if selected.get("eu") and not meetings_only_mode:
        if progress_callback:
            progress_callback("🇪🇺 Searching EU register...", done/total)
        matches["eu"] = search_eu_register(search_term) or []
        done += 1
    
    # EC Meetings - search by representative in minister mode, by topic in topic mode
    if minister_mode:
        if progress_callback:
            progress_callback("🇪🇺 Searching EC meetings by representative...", done/total)
        matches["ec_meetings"] = search_ec_meetings_by_representative(search_term, months_back=uk_months_back)
        done += 1
    elif topic_mode:
        if progress_callback:
            progress_callback("🇪🇺 Searching EC meetings by topic...", done/total)
        matches["ec_meetings"] = search_ec_meetings_by_topic(search_term, months_back=uk_months_back)
        done += 1
    elif department_mode:
        if progress_callback:
            progress_callback("🇪🇺 Searching EC meetings by Cabinet/DG...", done/total)
        matches["ec_meetings"] = search_ec_meetings_by_cabinet(search_term, months_back=uk_months_back)
        done += 1
    
    # France - just get matches
    if selected.get("france") and not meetings_only_mode:
        if progress_callback:
            progress_callback("🇫🇷 Searching France (HATVP)...", done/total)
        matches["france"] = search_france_register(search_term) or []
        done += 1
    
    # Germany - just get matches
    if selected.get("germany") and not meetings_only_mode:
        if progress_callback:
            progress_callback("🇩🇪 Searching Germany (Bundestag)...", done/total)
        matches["germany"] = search_germany_register(search_term) or []
        done += 1
    
    # UK - returns full data from index (fast, no second stage needed)
    if selected.get("uk"):
        if progress_callback:
            label = "🇬🇧 Searching UK meetings by minister..." if minister_mode else "🇬🇧 Searching UK meetings by topic..." if topic_mode else "🇬🇧 Searching UK meetings by department..." if department_mode else "🇬🇧 Searching UK meetings..."
            progress_callback(label, done/total)
        matches["uk"] = search_uk_ministerial_meetings(search_term, months_back=uk_months_back, search_field=search_field)
        done += 1
    
    # Ireland - uses index, returns full results directly
    if selected.get("ireland") and not meetings_only_mode:
        if progress_callback:
            progress_callback("🇮🇪 Searching Ireland lobbying register...", done/total)
        matches["ireland"] = search_ireland_lobbying(search_term)
        done += 1
    
    # Netherlands - uses index, returns full results directly
    if selected.get("netherlands") and not meetings_only_mode:
        if progress_callback:
            progress_callback("🇳🇱 Searching Netherlands agendas...", done/total)
        matches["netherlands"] = search_netherlands_agendas(search_term)
        done += 1
    
    # Austria - returns full data
    if selected.get("austria") and not meetings_only_mode:
        if progress_callback:
            progress_callback("🇦🇹 Searching Austria...", done/total)
        matches["austria"] = search_austria_register(search_term)
        done += 1
    
    # Catalonia - returns full data
    if selected.get("catalonia") and not meetings_only_mode:
        if progress_callback:
            progress_callback("🏴󠁥󠁳󠁣󠁴󠁿 Searching Catalonia...", done/total)
        matches["catalonia"] = search_catalonia_register(search_term)
        done += 1
    
    # Finland - returns full data
    if selected.get("finland") and not meetings_only_mode:
        if progress_callback:
            progress_callback("🇫🇮 Searching Finland...", done/total)
        matches["finland"] = search_finland_register(search_term)
        done += 1
    
    # Slovenia - returns full data
    if selected.get("slovenia") and not meetings_only_mode:
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
        "uk": None,
        "uk_officials": None,
        "ec_meetings": None,
        "ireland": None,
        "netherlands": None,
        "austria": None,
        "catalonia": None,
        "finland": None,
        "slovenia": None,
    }
    
    # Include index-based jurisdictions only if user selected them
    if selections.get("uk"):
        results["uk"] = other_results.get("uk")
        results["uk_officials"] = other_results.get("uk_officials")
    if selections.get("ec_meetings"):
        results["ec_meetings"] = other_results.get("ec_meetings")
    if selections.get("ireland"):
        results["ireland"] = other_results.get("ireland")
    if selections.get("netherlands"):
        results["netherlands"] = other_results.get("netherlands")
    if selections.get("austria"):
        results["austria"] = other_results.get("austria")
    if selections.get("catalonia"):
        results["catalonia"] = other_results.get("catalonia")
    if selections.get("finland"):
        results["finland"] = other_results.get("finland")
    if selections.get("slovenia"):
        results["slovenia"] = other_results.get("slovenia")
    
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
            try:
                results["eu"] = fetch_eu_data(eu_selections[0]["id"])
            except Exception as e:
                st.warning(f"⚠️ Could not fetch EU data for {eu_selections[0]['name']}: {e}")
                results["eu"] = None
            done += 1
        else:
            # Multiple selections - use multiple_entities format
            all_entities = []
            for sel in eu_selections:
                if progress_callback:
                    progress_callback(f"🇪🇺 Fetching {sel['name']}...", done/total)
                try:
                    entity_data = fetch_eu_data(sel["id"])
                    if entity_data:
                        entity_data["matched_term"] = sel["name"]
                        entity_data["matched_name"] = sel["name"]
                        all_entities.append(entity_data)
                except Exception as e:
                    st.warning(f"⚠️ Could not fetch EU data for {sel['name']}: {e}")
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
            try:
                results["france"] = fetch_france_data(fr_selections[0]["id"])
            except Exception as e:
                st.warning(f"⚠️ Could not fetch France data for {fr_selections[0]['name']}: {e}")
                results["france"] = None
            done += 1
        else:
            all_entities = []
            for sel in fr_selections:
                if progress_callback:
                    progress_callback(f"🇫🇷 Fetching {sel['name']}...", done/total)
                try:
                    entity_data = fetch_france_data(sel["id"])
                    if entity_data:
                        entity_data["matched_term"] = sel["name"]
                        entity_data["matched_name"] = sel["name"]
                        all_entities.append(entity_data)
                except Exception as e:
                    st.warning(f"⚠️ Could not fetch France data for {sel['name']}: {e}")
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
            try:
                results["germany"] = fetch_germany_data(de_selections[0]["register_number"])
            except Exception as e:
                st.warning(f"⚠️ Could not fetch Germany data for {de_selections[0]['name']}: {e}")
                results["germany"] = None
            done += 1
        else:
            all_entities = []
            for sel in de_selections:
                if progress_callback:
                    progress_callback(f"🇩🇪 Fetching {sel['name']}...", done/total)
                try:
                    entity_data = fetch_germany_data(sel["register_number"])
                    if entity_data:
                        entity_data["matched_term"] = sel["name"]
                        entity_data["matched_name"] = sel["name"]
                        all_entities.append(entity_data)
                except Exception as e:
                    st.warning(f"⚠️ Could not fetch Germany data for {sel['name']}: {e}")
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
st.markdown("Search corporate lobbying records across European transparency registers, or search UK meetings by minister/official name.")

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
    options=["All available", "Last 12 months", "Last 6 months", "Last 3 months"],
    index=0,
    help="Filter UK meetings by date. Data is most reliable from 2024 onwards."
)

st.sidebar.caption("Other jurisdictions search all available data")

# Search mode toggle
st.sidebar.markdown("---")
search_mode = st.sidebar.radio(
    "🔍 Search mode",
    options=["Organisation", "Minister / Official", "Department / DG", "Topic / Subject"],
    index=0,
    help="Search by company name, minister/commissioner name, government department/DG, or meeting topic keyword."
)
is_minister_search = search_mode == "Minister / Official"
is_topic_search = search_mode == "Topic / Subject"
is_department_search = search_mode == "Department / DG"

if is_minister_search:
    st.sidebar.caption("🇬🇧🇪🇺 Minister search uses UK meetings index and EC meetings data. Other registers are disabled.")
elif is_topic_search:
    st.sidebar.caption("🇬🇧🇪🇺 Topic search finds UK and EC meetings by subject/purpose keyword. Other registers are disabled.")
elif is_department_search:
    st.sidebar.caption("🇬🇧🇪🇺 Department search finds UK meetings by department and EC meetings by Cabinet/DG. Other registers are disabled.")

# Main search
col1, col2 = st.columns([3, 1])
with col1:
    if is_minister_search:
        search_term = st.text_input(
            "🔍 Minister, Commissioner, or official name",
            placeholder="e.g. Gareth Davies, Henna Virkkunen, Rachel Reeves...",
            help="Search UK ministerial meetings and EC Commissioner/cabinet/DG meetings by name"
        )
    elif is_topic_search:
        search_term = st.text_input(
            "🔍 Meeting topic or subject keyword",
            placeholder="e.g. AI Act, Green Deal, Digital Markets Act, Net Zero...",
            help="Search UK and EC meetings by topic/subject. Supports Boolean: 'AI AND regulation', 'Green Deal OR climate'"
        )
    elif is_department_search:
        search_term = st.text_input(
            "🔍 Department or DG name",
            placeholder="e.g. DCMS, Treasury, DG COMP, Trade, Digital...",
            help="Search UK meetings by department and EC meetings by Cabinet/DG name"
        )
    else:
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
        
        # Convert UK date filter to months_back
        uk_months_map = {
            "Last 12 months": 12,
            "Last 6 months": 6,
            "Last 3 months": 3,
            "All available": None,
        }
        uk_months = uk_months_map.get(uk_date_filter, None)
        
        # Stage 1: Get matches
        search_field = "minister" if is_minister_search else "topic" if is_topic_search else "department" if is_department_search else "organisation"
        st.session_state.matches = preview_matches(search_term, selected, update_progress, uk_months_back=uk_months, search_field=search_field)
        st.session_state.search_mode = search_mode
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
    
    # UK matches - show with checkbox to include/exclude
    if matches.get("uk") and matches["uk"].get("meetings"):
        uk_meetings = matches["uk"]["meetings"]
        uk_is_minister_mode = matches["uk"].get("search_field") == "minister"
        uk_is_topic_mode = matches["uk"].get("search_field") == "topic"
        uk_is_dept_mode = matches["uk"].get("search_field") == "department"
        
        if uk_is_minister_mode:
            # Minister search mode - show organisation breakdown
            by_org = matches["uk"].get("by_organisation", {})
            by_minister = matches["uk"].get("by_minister", {})
            
            expander_label = f"🇬🇧 **UK Minister/Official Meetings** - {len(uk_meetings)} meetings with {len(by_org)} organisations"
            with st.expander(expander_label, expanded=True):
                uk_include = st.checkbox(
                    f"Include UK meetings ({len(uk_meetings)} meetings)",
                    value=True,
                    key="uk_include"
                )
                
                # Show which ministers/officials matched
                if by_minister:
                    minister_summary = ", ".join(f"{name}: {count}" for name, count in list(by_minister.items())[:5])
                    st.caption(f"**Matched:** {minister_summary}")
                
                # Show top organisations met
                if by_org:
                    top_orgs = list(by_org.items())[:8]
                    org_summary = ", ".join(f"{org}: {count}" for org, count in top_orgs)
                    st.caption(f"**Top orgs met:** {org_summary}")
                    if len(by_org) > 8:
                        st.caption(f"...and {len(by_org) - 8} more")
                
                user_selections["uk"] = uk_include
        elif uk_is_topic_mode:
            # Topic search mode - show who's lobbying on this topic
            by_org = matches["uk"].get("by_organisation", {})
            by_minister = matches["uk"].get("by_minister", {})
            
            expander_label = f"🇬🇧 **UK Meetings on Topic** - {len(uk_meetings)} meetings, {len(by_org)} organisations"
            with st.expander(expander_label, expanded=True):
                uk_include = st.checkbox(
                    f"Include UK meetings ({len(uk_meetings)} meetings)",
                    value=True,
                    key="uk_include"
                )
                
                if by_org:
                    top_orgs = list(by_org.items())[:8]
                    org_summary = ", ".join(f"{org}: {count}" for org, count in top_orgs)
                    st.caption(f"**Top orgs:** {org_summary}")
                
                if by_minister:
                    minister_summary = ", ".join(f"{name}: {count}" for name, count in list(by_minister.items())[:5])
                    st.caption(f"**Ministers involved:** {minister_summary}")
                
                user_selections["uk"] = uk_include
        elif uk_is_dept_mode:
            # Department search mode
            by_org = matches["uk"].get("by_organisation", {})
            by_dept = matches["uk"].get("by_department", {})
            
            expander_label = f"🇬🇧 **UK Department Meetings** - {len(uk_meetings)} meetings, {len(by_org)} organisations"
            with st.expander(expander_label, expanded=True):
                uk_include = st.checkbox(
                    f"Include UK meetings ({len(uk_meetings)} meetings)",
                    value=True,
                    key="uk_include"
                )
                
                if by_dept:
                    dept_summary = ", ".join(f"{d}: {c}" for d, c in list(by_dept.items())[:5])
                    st.caption(f"**Matched departments:** {dept_summary}")
                
                if by_org:
                    top_orgs = list(by_org.items())[:8]
                    org_summary = ", ".join(f"{org}: {count}" for org, count in top_orgs)
                    st.caption(f"**Top orgs:** {org_summary}")
                
                user_selections["uk"] = uk_include
        else:
            # Organisation search mode (default)
            with st.expander(f"🇬🇧 **UK Ministerial Meetings** - {len(uk_meetings)} meetings found", expanded=True):
                uk_include = st.checkbox(
                    f"Include UK meetings ({len(uk_meetings)} meetings)",
                    value=True,
                    key="uk_include"
                )
                
                # Show breakdown by organisation if OR query
                by_org = {}
                for m in uk_meetings:
                    org = m.get("matched_term") or m.get("organisation", "Unknown")
                    by_org[org] = by_org.get(org, 0) + 1
                
                if len(by_org) > 1:
                    st.caption("Breakdown: " + ", ".join(f"{org}: {count}" for org, count in sorted(by_org.items(), key=lambda x: -x[1])[:5]))
                
                user_selections["uk"] = uk_include
    
    # EC Meetings matches (minister mode)
    if matches.get("ec_meetings") and matches["ec_meetings"].get("meetings"):
        ec_data = matches["ec_meetings"]
        ec_meetings = ec_data["meetings"]
        by_org = ec_data.get("by_organisation", {})
        by_rep = ec_data.get("by_representative", {})
        
        expander_label = f"🇪🇺 **EC Commissioner/Official Meetings** - {len(ec_meetings)} meetings with {len(by_org)} organisations"
        with st.expander(expander_label, expanded=True):
            ec_include = st.checkbox(
                f"Include EC meetings ({len(ec_meetings)} meetings)",
                value=True,
                key="ec_meetings_include"
            )
            
            # Show matched representatives
            if by_rep:
                rep_summary = ", ".join(f"{name}: {count}" for name, count in list(by_rep.items())[:5])
                st.caption(f"**Matched:** {rep_summary}")
            
            # Show top orgs
            if by_org:
                top_orgs = list(by_org.items())[:6]
                org_summary = ", ".join(f"{org}: {count}" for org, count in top_orgs)
                st.caption(f"**Top orgs met:** {org_summary}")
            
            user_selections["ec_meetings"] = ec_include
    
    # Ireland matches
    if matches.get("ireland") and matches["ireland"].get("returns"):
        ie_data = matches["ireland"]
        ie_count = ie_data.get("return_count", len(ie_data.get("returns", [])))
        with st.expander(f"🇮🇪 **Ireland (Lobbying.ie)** - {ie_count} returns found", expanded=True):
            ie_include = st.checkbox(
                f"Include Ireland ({ie_count} lobbying returns)",
                value=True,
                key="ireland_include"
            )
            
            # Show top lobbyists
            by_lobbyist = ie_data.get("by_lobbyist", {})
            if by_lobbyist:
                top = list(by_lobbyist.items())[:3]
                st.caption("Top lobbyists: " + ", ".join(f"{name}: {count}" for name, count in top))
            
            user_selections["ireland"] = ie_include
    
    # Netherlands matches
    if matches.get("netherlands") and matches["netherlands"].get("appointments"):
        nl_data = matches["netherlands"]
        nl_count = nl_data.get("appointment_count", len(nl_data.get("appointments", [])))
        with st.expander(f"🇳🇱 **Netherlands (Ministerial Agendas)** - {nl_count} appointments found", expanded=True):
            nl_include = st.checkbox(
                f"Include Netherlands ({nl_count} agenda appointments)",
                value=True,
                key="netherlands_include"
            )
            
            # Show breakdown by minister
            by_minister = nl_data.get("by_minister", {})
            if by_minister:
                top = list(by_minister.items())[:3]
                st.caption("Top ministers: " + ", ".join(f"{name}: {count}" for name, count in top))
            
            user_selections["netherlands"] = nl_include
    
    # Austria matches
    if matches.get("austria") and matches["austria"].get("entries"):
        austria_data = matches["austria"]
        with st.expander(f"🇦🇹 **Austria** - {austria_data['entry_count']} entries found", expanded=True):
            austria_include = st.checkbox(
                f"Include Austria ({austria_data['entry_count']} register entries)",
                value=True,
                key="austria_include"
            )
            user_selections["austria"] = austria_include
    
    # Catalonia matches
    if matches.get("catalonia") and matches["catalonia"].get("entries"):
        cat_data = matches["catalonia"]
        with st.expander(f"🏴󠁥󠁳󠁣󠁴󠁿 **Catalonia** - {cat_data['entry_count']} entries found", expanded=True):
            cat_include = st.checkbox(
                f"Include Catalonia ({cat_data['entry_count']} register entries)",
                value=True,
                key="catalonia_include"
            )
            user_selections["catalonia"] = cat_include
    
    # Finland matches
    if matches.get("finland") and matches["finland"].get("entries"):
        fin_data = matches["finland"]
        with st.expander(f"🇫🇮 **Finland** - {fin_data['entry_count']} entries found", expanded=True):
            fin_include = st.checkbox(
                f"Include Finland ({fin_data['entry_count']} register entries)",
                value=True,
                key="finland_include"
            )
            user_selections["finland"] = fin_include
    
    # Slovenia matches
    if matches.get("slovenia") and matches["slovenia"].get("entries"):
        slo_data = matches["slovenia"]
        with st.expander(f"🇸🇮 **Slovenia** - {slo_data['entry_count']} entries found", expanded=True):
            slo_include = st.checkbox(
                f"Include Slovenia ({slo_data['entry_count']} lobbyists)",
                value=True,
                key="slovenia_include"
            )
            user_selections["slovenia"] = slo_include
    
    # Show jurisdictions with no matches (skip in minister/topic mode since only UK+EC are searched)
    is_meetings_only_mode = st.session_state.get("search_mode") in ("Minister / Official", "Topic / Subject", "Department / DG")
    no_matches = []
    if not is_meetings_only_mode:
        if st.session_state.selected_jurisdictions.get("eu") and not matches.get("eu"):
            no_matches.append("🇪🇺 EU")
        if st.session_state.selected_jurisdictions.get("france") and not matches.get("france"):
            no_matches.append("🇫🇷 France")
        if st.session_state.selected_jurisdictions.get("germany") and not matches.get("germany"):
            no_matches.append("🇩🇪 Germany")
        if st.session_state.selected_jurisdictions.get("ireland") and not (matches.get("ireland") and matches["ireland"].get("returns")):
            no_matches.append("🇮🇪 Ireland")
        if st.session_state.selected_jurisdictions.get("netherlands") and not (matches.get("netherlands") and matches["netherlands"].get("appointments")):
            no_matches.append("🇳🇱 Netherlands")
        if st.session_state.selected_jurisdictions.get("austria") and not (matches.get("austria") and matches["austria"].get("entries")):
            no_matches.append("🇦🇹 Austria")
        if st.session_state.selected_jurisdictions.get("catalonia") and not (matches.get("catalonia") and matches["catalonia"].get("entries")):
            no_matches.append("🏴󠁥󠁳󠁣󠁴󠁿 Catalonia")
        if st.session_state.selected_jurisdictions.get("finland") and not (matches.get("finland") and matches["finland"].get("entries")):
            no_matches.append("🇫🇮 Finland")
        if st.session_state.selected_jurisdictions.get("slovenia") and not (matches.get("slovenia") and matches["slovenia"].get("entries")):
            no_matches.append("🇸🇮 Slovenia")
    if st.session_state.selected_jurisdictions.get("uk") and not (matches.get("uk") and matches["uk"].get("meetings")):
        no_matches.append("🇬🇧 UK")
    
    if no_matches:
        st.caption(f"No matches in: {', '.join(no_matches)}")
    
    # Fetch button
    st.markdown("---")
    
    # Count total selections
    total_orgs = len(user_selections.get("eu", [])) + len(user_selections.get("france", [])) + len(user_selections.get("germany", []))
    total_jurisdictions = total_orgs
    if user_selections.get("uk"):
        total_jurisdictions += 1
    if user_selections.get("ec_meetings"):
        total_jurisdictions += 1
    if user_selections.get("ireland"):
        total_jurisdictions += 1
    if user_selections.get("netherlands"):
        total_jurisdictions += 1
    if user_selections.get("austria"):
        total_jurisdictions += 1
    if user_selections.get("catalonia"):
        total_jurisdictions += 1
    if user_selections.get("finland"):
        total_jurisdictions += 1
    if user_selections.get("slovenia"):
        total_jurisdictions += 1
    
    fetch_button = st.button(
        f"📥 Generate Report",
        type="primary",
        disabled=(total_jurisdictions == 0)
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
    
    # Use appropriate filename based on search mode
    search_mode_label = st.session_state.get("search_mode", "Organisation")
    if search_mode_label == "Minister / Official":
        filename = f"{search_term_display.lower().replace(' ', '_')}_minister_meetings.xlsx"
    elif search_mode_label == "Topic / Subject":
        filename = f"{search_term_display.lower().replace(' ', '_')}_topic_meetings.xlsx"
    elif search_mode_label == "Department / DG":
        filename = f"{search_term_display.lower().replace(' ', '_')}_department_meetings.xlsx"
    else:
        filename = f"{search_term_display.lower().replace(' ', '_')}_lobbying_full.xlsx"
    
    st.download_button(
        label="📊 Download Comprehensive Excel Report",
        data=excel_buffer,
        file_name=filename,
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
st.caption("Data: LobbyFacts.eu • HATVP • Bundestag • GOV.UK • lobbying.ie • openlobby.nl • lobbyreg.justiz.gv.at • transparenciacatalunya.cat • avoimuusrekisteri.fi")
