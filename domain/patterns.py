import re
from typing import Dict, Pattern 

PLACEHOLDER_PATTERNS : Dict[str, Pattern[str]] = {
    # {{ customer.name }}
    "text": re.compile(r"\{\{\s*([^{}]+)\s*\}\}"),

    # $<$ clause_name $>$
    "clause_block": re.compile(r"\$<\$([^$]+)\$>\$"),

    # <$ clause_name $>
    "single_clause_block": re.compile(r"<\$\s*([^$<>]+?)\s*\$>"),

    # <<placeholder>>
    "angle": re.compile(r"<<([^<>]+)>>"),

    # example hash (#) functions doctiger: #146#function_call#146#
    "function": re.compile(r"#146#(.+?)#146#"),
}