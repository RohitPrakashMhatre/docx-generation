import io
import re
import uuid
import hashlib
from datetime import datetime
from docx import Document

from domain.patterns import PLACEHOLDER_PATTERNS

def file_hash(file_bytes: bytes) -> str:
    return hashlib.sha256(file_bytes).hexdigest()

def normalize_placeholder_text(text: str) -> str:
    # Removes newlines, tabs, multiple spaces
    return re.sub(r"\s+", "", text)

def detect_placeholders_with_spans(text: str):
    results = []

    for ptype, regex in PLACEHOLDER_PATTERNS.items():
        for m in regex.finditer(text):
            results.append({
                "type": ptype,
                "raw": m.group(0),
                "key": m.group(1),
                "start": m.start(),
                "end": m.end()
            })

    return results

def parse_paragraphs(doc: Document):
    placeholders = []

    for p_idx, para in enumerate(doc.paragraphs):
        raw_text = para.text
        if not raw_text.strip():
            continue

        normalized_text = normalize_placeholder_text(raw_text)
        detected = detect_placeholders_with_spans(normalized_text)

        for d in detected:
            placeholders.append({
                "id": f"ph_{uuid.uuid4().hex[:8]}",
                "key": d["key"],
                "syntax": d["raw"],
                "type": d["type"],
                "scope": "body",
                "location": {
                    "paragraph_index": p_idx
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                }
            })

    return placeholders


def parse_tables(doc: Document):
    placeholders = []

    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):

                cell_text = "".join(p.text for p in cell.paragraphs)
                normalized = normalize_placeholder_text(cell_text)
                detected = detect_placeholders_with_spans(normalized)

                for d in detected:
                    placeholders.append({
                        "id": f"ph_{uuid.uuid4().hex[:8]}",
                        "key": d["key"],
                        "syntax": d["raw"],
                        "type": d["type"],
                        "scope": "table",
                        "location": {
                            "table": t_idx,
                            "row": r_idx,
                            "cell": c_idx
                        },
                        "rules": {
                            "preserve_style": True,
                            "allow_multiline": True
                        }
                    })

    return placeholders
