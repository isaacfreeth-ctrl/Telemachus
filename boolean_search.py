"""Boolean search helpers for the lobbying tracker."""

def is_boolean_query(term):
    if not term:
        return False
    upper = term.upper()
    return ' OR ' in upper or ' AND ' in upper or ' NOT ' in upper or '"' in term

def is_or_query(term):
    """Check if the search term contains an OR operator."""
    if not term:
        return False
    return ' OR ' in term.upper()

def extract_or_terms(query):
    """Extract individual terms from an OR query. Returns list of terms."""
    if not query:
        return []
    terms = [t.strip().strip('"') for t in query.upper().split(' OR ')]
    # Return original-case versions by splitting the original query
    parts = []
    remaining = query
    for t in terms:
        # Find this term in the remaining string (case-insensitive)
        idx = remaining.upper().find(t)
        if idx >= 0:
            original = remaining[idx:idx+len(t)].strip()
            parts.append(original)
    return parts if parts else [query]

def get_matching_term(query, text):
    """Given an OR query and text, return which OR term matched."""
    if not text or not query:
        return query
    text_lower = text.lower()
    terms = extract_or_terms(query)
    for term in terms:
        if term.lower() in text_lower:
            return term
    return terms[0] if terms else query

def boolean_match(query, text):
    if not text:
        return False
    text_lower = text.lower()
    query_upper = query.upper()
    
    if ' OR ' in query_upper:
        terms = [t.strip().strip('"').lower() for t in query.split(' OR ') if t.strip()]
        return any(t in text_lower for t in terms)
    elif ' AND ' in query_upper:
        terms = [t.strip().strip('"').lower() for t in query.split(' AND ') if t.strip()]
        return all(t in text_lower for t in terms)
    elif ' NOT ' in query_upper:
        parts = query.upper().split(' NOT ', 1)
        include = parts[0].strip().strip('"').lower()
        exclude = parts[1].strip().strip('"').lower()
        return include in text_lower and exclude not in text_lower
    else:
        return query.lower().strip('"') in text_lower

def parse_boolean_query(query):
    return query
