"""
Boolean Search Parser for Lobbying Tracker

Supports:
- AND: "shell AND bp" - both terms must match
- OR: "shell OR bp" - either term matches  
- NOT: "shell NOT gas" - first term but not second
- Parentheses: "(shell OR bp) AND climate"
- Quoted phrases: '"big oil"' - exact phrase match
- Implicit AND: "shell bp" treated as "shell AND bp"

Case-insensitive matching against organisation/registrant names.
"""

import re
from typing import Callable, List, Union


class BooleanSearchParser:
    """
    Parses and evaluates Boolean search expressions.
    
    Grammar:
        expression := term ((AND | OR) term)*
        term := NOT? factor
        factor := (expression) | quoted_string | word
    """
    
    def __init__(self, query: str):
        self.original_query = query
        self.tokens = self._tokenize(query)
        self.pos = 0
        
    def _tokenize(self, query: str) -> List[str]:
        """
        Tokenize the query into operators, parentheses, quoted strings, and words.
        """
        tokens = []
        i = 0
        query = query.strip()
        
        while i < len(query):
            # Skip whitespace
            if query[i].isspace():
                i += 1
                continue
                
            # Parentheses
            if query[i] in '()':
                tokens.append(query[i])
                i += 1
                continue
                
            # Quoted string
            if query[i] == '"':
                end = query.find('"', i + 1)
                if end == -1:
                    # Unclosed quote - take rest of string
                    tokens.append(query[i+1:])
                    break
                tokens.append(query[i:end+1])  # Include quotes
                i = end + 1
                continue
                
            # Word or operator
            word = ""
            while i < len(query) and not query[i].isspace() and query[i] not in '()"':
                word += query[i]
                i += 1
                
            if word:
                # Check for operators (case-insensitive)
                upper = word.upper()
                if upper in ('AND', 'OR', 'NOT', '&&', '||', '!'):
                    # Normalize operators
                    if upper == '&&':
                        tokens.append('AND')
                    elif upper == '||':
                        tokens.append('OR')
                    elif upper == '!':
                        tokens.append('NOT')
                    else:
                        tokens.append(upper)
                else:
                    tokens.append(word)
                    
        return tokens
    
    def parse(self) -> 'BooleanNode':
        """Parse the tokenized query into an AST."""
        if not self.tokens:
            return LiteralNode("")
        
        result = self._parse_expression()
        
        # Handle implicit AND for remaining tokens
        while self.pos < len(self.tokens):
            token = self.tokens[self.pos]
            if token not in ('AND', 'OR', 'NOT', '(', ')'):
                # Implicit AND
                right = self._parse_term()
                result = AndNode(result, right)
            else:
                break
                
        return result
    
    def _parse_expression(self) -> 'BooleanNode':
        """Parse: term ((AND | OR | NOT) term)*"""
        left = self._parse_term()
        
        while self.pos < len(self.tokens):
            token = self.tokens[self.pos]
            
            if token == 'AND':
                self.pos += 1
                right = self._parse_term()
                left = AndNode(left, right)
            elif token == 'OR':
                self.pos += 1
                right = self._parse_term()
                left = OrNode(left, right)
            elif token == 'NOT':
                # "A NOT B" means "A AND NOT B"
                self.pos += 1
                right = self._parse_factor()  # Get the term after NOT
                left = AndNode(left, NotNode(right))
            elif token not in (')') and token not in ('AND', 'OR', 'NOT'):
                # Implicit AND for adjacent terms
                right = self._parse_term()
                left = AndNode(left, right)
            else:
                break
                
        return left
    
    def _parse_term(self) -> 'BooleanNode':
        """Parse: NOT? factor"""
        if self.pos < len(self.tokens) and self.tokens[self.pos] == 'NOT':
            self.pos += 1
            factor = self._parse_factor()
            return NotNode(factor)
        return self._parse_factor()
    
    def _parse_factor(self) -> 'BooleanNode':
        """Parse: (expression) | quoted_string | word"""
        if self.pos >= len(self.tokens):
            return LiteralNode("")
            
        token = self.tokens[self.pos]
        
        if token == '(':
            self.pos += 1
            expr = self._parse_expression()
            # Consume closing paren
            if self.pos < len(self.tokens) and self.tokens[self.pos] == ')':
                self.pos += 1
            return expr
        elif token.startswith('"') and token.endswith('"'):
            # Quoted phrase - strip quotes
            self.pos += 1
            return LiteralNode(token[1:-1], exact_phrase=True)
        elif token in ('AND', 'OR', 'NOT', ')'):
            # Shouldn't happen in well-formed query
            return LiteralNode("")
        else:
            self.pos += 1
            return LiteralNode(token)


class BooleanNode:
    """Base class for AST nodes."""
    
    def evaluate(self, text: str) -> bool:
        """Evaluate if this node matches the given text."""
        raise NotImplementedError
        
    def __repr__(self):
        return f"{self.__class__.__name__}()"


class LiteralNode(BooleanNode):
    """Matches a literal string (substring or exact phrase)."""
    
    def __init__(self, value: str, exact_phrase: bool = False):
        self.value = value.lower()
        self.exact_phrase = exact_phrase
        
    def evaluate(self, text: str) -> bool:
        if not self.value:
            return True  # Empty matches everything
        text_lower = text.lower()
        if self.exact_phrase:
            # Word boundary matching for phrases
            pattern = r'\b' + re.escape(self.value) + r'\b'
            return bool(re.search(pattern, text_lower))
        return self.value in text_lower
        
    def __repr__(self):
        return f"Literal({self.value!r}, exact={self.exact_phrase})"


class AndNode(BooleanNode):
    """Logical AND of two nodes."""
    
    def __init__(self, left: BooleanNode, right: BooleanNode):
        self.left = left
        self.right = right
        
    def evaluate(self, text: str) -> bool:
        return self.left.evaluate(text) and self.right.evaluate(text)
        
    def __repr__(self):
        return f"AND({self.left}, {self.right})"


class OrNode(BooleanNode):
    """Logical OR of two nodes."""
    
    def __init__(self, left: BooleanNode, right: BooleanNode):
        self.left = left
        self.right = right
        
    def evaluate(self, text: str) -> bool:
        return self.left.evaluate(text) or self.right.evaluate(text)
        
    def __repr__(self):
        return f"OR({self.left}, {self.right})"


class NotNode(BooleanNode):
    """Logical NOT of a node."""
    
    def __init__(self, child: BooleanNode):
        self.child = child
        
    def evaluate(self, text: str) -> bool:
        return not self.child.evaluate(text)
        
    def __repr__(self):
        return f"NOT({self.child})"


def parse_boolean_query(query: str) -> BooleanNode:
    """
    Parse a Boolean search query into an evaluatable AST.
    
    Examples:
        parse_boolean_query("shell")
        parse_boolean_query("shell AND bp")
        parse_boolean_query("shell OR bp")
        parse_boolean_query("shell NOT gas")
        parse_boolean_query("(shell OR bp) AND climate")
        parse_boolean_query('"big oil"')
    """
    parser = BooleanSearchParser(query)
    return parser.parse()


def boolean_match(query: str, text: str) -> bool:
    """
    Check if text matches a Boolean query.
    
    Args:
        query: Boolean search expression
        text: Text to match against (e.g., organisation name)
        
    Returns:
        True if the text matches the query
        
    Examples:
        >>> boolean_match("shell", "Shell plc")
        True
        >>> boolean_match("shell AND bp", "Shell plc")
        False
        >>> boolean_match("shell OR bp", "BP plc")
        True
        >>> boolean_match("shell NOT gas", "Shell Gas")
        False
        >>> boolean_match('"shell plc"', "Shell plc trading")
        True
    """
    ast = parse_boolean_query(query)
    return ast.evaluate(text)


def is_boolean_query(query: str) -> bool:
    """
    Check if a query uses Boolean operators.
    
    Returns True if the query contains AND, OR, NOT, parentheses, or quotes.
    This helps determine if we should use the Boolean parser or simple matching.
    """
    # Normalize to uppercase for checking
    upper = query.upper()
    
    # Check for explicit operators
    if ' AND ' in upper or ' OR ' in upper or ' NOT ' in upper:
        return True
        
    # Check for symbols
    if '(' in query or ')' in query:
        return True
        
    # Check for quoted phrases
    if '"' in query:
        return True
        
    # Check for symbolic operators
    if ' && ' in query or ' || ' in query or query.startswith('!'):
        return True
        
    return False


def is_or_query(query: str) -> bool:
    """Check if query is a simple OR query (term1 OR term2 OR ...)."""
    upper = query.upper()
    return ' OR ' in upper or ' || ' in upper


def extract_or_terms(query: str) -> List[str]:
    """
    Extract individual search terms from an OR query.
    
    For "shell OR bp OR exxon", returns ["shell", "bp", "exxon"]
    For "shell AND bp", returns ["shell AND bp"] (not a pure OR query)
    For "(shell OR bp) AND energy", returns ["shell", "bp"] (extracts OR terms)
    
    This is useful for fetching data for each entity separately in an OR search.
    """
    upper = query.upper()
    
    # Check if it's a pure OR query (no AND, no parentheses)
    if ' AND ' in upper or '(' in query:
        # Complex query - try to extract top-level OR terms
        # For now, if it has AND, treat as single search
        if ' AND ' in upper:
            return [query]
        # Has parentheses but no AND - might be "(shell OR bp)"
        # Extract terms from inside
        import re
        # Try to find OR-separated terms
        inner = query.strip('()')
        if ' OR ' in inner.upper():
            parts = re.split(r'\s+OR\s+', inner, flags=re.IGNORECASE)
            return [p.strip().strip('"') for p in parts if p.strip()]
    
    # Simple OR query
    if ' OR ' in upper or ' || ' in upper:
        import re
        parts = re.split(r'\s+(?:OR|\|\|)\s+', query, flags=re.IGNORECASE)
        return [p.strip().strip('"') for p in parts if p.strip()]
    
    # Single term
    return [query]


def get_matching_term(query: str, text: str) -> str:
    """
    For an OR query, return which specific term matched the text.
    
    For "shell OR bp" matching "Shell plc", returns "shell"
    For "shell OR bp" matching "BP plc", returns "bp"
    
    Returns the original query if not an OR query or no specific match found.
    """
    terms = extract_or_terms(query)
    
    if len(terms) <= 1:
        return query
    
    text_lower = text.lower()
    for term in terms:
        term_lower = term.lower().strip('"')
        if term_lower in text_lower:
            return term
    
    # No specific match found, return first matching term
    return query


# Convenience function for filtering lists
def filter_by_boolean_query(query: str, items: list, key_func: Callable = None) -> list:
    """
    Filter a list of items using a Boolean query.
    
    Args:
        query: Boolean search expression
        items: List of items to filter
        key_func: Function to extract searchable text from each item.
                  If None, items are converted to strings.
                  
    Returns:
        Filtered list of items matching the query
        
    Examples:
        >>> orgs = [{"name": "Shell plc"}, {"name": "BP plc"}]
        >>> filter_by_boolean_query("shell", orgs, lambda x: x["name"])
        [{"name": "Shell plc"}]
    """
    ast = parse_boolean_query(query)
    
    if key_func is None:
        key_func = str
        
    return [item for item in items if ast.evaluate(key_func(item))]


if __name__ == "__main__":
    # Test cases
    print("Testing Boolean Search Parser")
    print("=" * 50)
    
    test_cases = [
        # (query, text, expected)
        ("shell", "Shell plc", True),
        ("shell", "BP plc", False),
        ("shell AND plc", "Shell plc", True),
        ("shell AND bp", "Shell plc", False),
        ("shell OR bp", "Shell plc", True),
        ("shell OR bp", "BP plc", True),
        ("shell OR bp", "Exxon", False),
        ("shell NOT gas", "Shell Energy", True),
        ("shell NOT gas", "Shell Gas Trading", False),
        ("(shell OR bp) AND energy", "Shell Energy", True),
        ("(shell OR bp) AND energy", "BP Energy", True),
        ("(shell OR bp) AND energy", "Shell plc", False),
        ('"shell plc"', "Shell plc trading", True),
        ('"shell plc"', "Shell Exploration plc", False),
        ("palantir", "Palantir Technologies UK", True),
        ("palantir NOT uk", "Palantir Technologies UK", False),
        ("palantir NOT uk", "Palantir Technologies", True),
        # Implicit AND
        ("shell bp", "Shell BP Joint Venture", True),
        ("shell bp", "Shell plc", False),
        # Multiple ORs
        ("shell OR bp OR exxon", "Exxon Mobil", True),
        # Complex nesting
        ("(shell OR bp) AND (energy OR gas)", "Shell Gas", True),
        ("(shell OR bp) AND (energy OR gas)", "Shell plc", False),
    ]
    
    passed = 0
    failed = 0
    
    for query, text, expected in test_cases:
        result = boolean_match(query, text)
        status = "✓" if result == expected else "✗"
        if result == expected:
            passed += 1
        else:
            failed += 1
        print(f"{status} Query: {query!r:40} Text: {text!r:30} Expected: {expected}, Got: {result}")
    
    print("=" * 50)
    print(f"Results: {passed} passed, {failed} failed")
    
    # Show AST for complex query
    print("\nAST for '(shell OR bp) AND energy':")
    ast = parse_boolean_query("(shell OR bp) AND energy")
    print(f"  {ast}")
