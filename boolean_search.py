"""Boolean search helpers for the lobbying tracker."""

import re

# ---------------------------------------------------------------------------
# Gap 7: topic-match robustness.
#
# Narrow word-order / vocabulary assumptions missed relevant rows. Topic matching
# should be word-order-insensitive (token-set, not substring), aware of common
# variant spellings, and must NOT let a bare "x" token false-match SpaceX/Entity X.
# ---------------------------------------------------------------------------

# Variant spellings that should match interchangeably. Each canonical maps to the
# set of surface forms that mean the same thing for topic matching.
TOPIC_VARIANTS = {
    "tiktok": ["tiktok", "tik tok"],
    "twitter": ["twitter", "x corp", "x corporation", "twitter/x"],
    "online safety": ["online safety", "safety online", "children's safety online",
                       "child safety online", "online harms", "internet safety"],
}

# Tokens that are too ambiguous to match on their own. A bare "x" must not match
# "SpaceX", "Entity X", etc. — it only counts when an adjacent qualifier disambiguates.
AMBIGUOUS_BARE_TOKENS = {"x"}


def _tokenize(text):
    """Lowercase word tokens (alphanumerics), order-independent."""
    return re.findall(r"[a-z0-9']+", (text or "").lower())


def expand_variants(term):
    """
    Return the list of surface forms to test for a topic term, including known
    variant spellings. Always includes the original term.
    """
    low = (term or "").strip().lower()
    forms = {low}
    for canonical, variants in TOPIC_VARIANTS.items():
        vset = {v.lower() for v in variants}
        if low == canonical or low in vset:
            forms.update(vset)
            forms.add(canonical)
    return [f for f in forms if f]


def _phrase_in_tokens(phrase_tokens, text_tokens):
    """
    Order-insensitive containment: every token of the phrase must appear in the
    text's token multiset. Single-token ambiguous terms (bare "x") never match here
    on their own — they require the multi-word variant forms instead.
    """
    if not phrase_tokens:
        return False
    if len(phrase_tokens) == 1 and phrase_tokens[0] in AMBIGUOUS_BARE_TOKENS:
        return False
    text_set = set(text_tokens)
    return all(tok in text_set for tok in phrase_tokens)


def topic_match(term, text):
    """
    Word-order-insensitive, variant-aware topic match (gap 7).

    Matches "online safety" against "children's safety online" and "safety online";
    matches "tiktok" against "tik tok"; matches "twitter" against "x corp". A bare
    "x" query will NOT match SpaceX / Entity X because the single ambiguous token is
    rejected unless a disambiguating variant form is supplied.
    """
    if not term or not text:
        return False
    text_tokens = _tokenize(text)
    for form in expand_variants(term):
        if _phrase_in_tokens(_tokenize(form), text_tokens):
            return True
    return False


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
