"""Minimal stub for boolean search functions."""

def is_boolean_query(term):
    if not term:
        return False
    upper = term.upper()
    return ' OR ' in upper or ' AND ' in upper or ' NOT ' in upper or '"' in term

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
