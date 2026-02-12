from flask import Flask, request, jsonify
from docx import Document
import re
import uuid
import hashlib
from datetime import datetime
import time
import io
import uuid
from collections import defaultdict


trace_id = uuid.uuid4().hex[:8]


app = Flask(__name__)


# propstrength__application_booking__c.propstrength__total_basic_sales_price__c


PLACEHOLDER_PATTERNS = {
    "text": re.compile(r"\{\{\s*([^{}]+)\s*\}\}"),
    "clause_block": re.compile(r"\$<\$([^$]+)\$>\$"),
    "single_clause_block": re.compile(r"<\$\s*([^$<>]+?)\s*\$>"),
    "angle": re.compile(r"<<([^<>]+)>>"),
    "function": re.compile(r"#146#(.+?)#146#")
}

def file_hash(file_bytes):
    return hashlib.sha256(file_bytes).hexdigest()

def detect_placeholders_with_spans(text):
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


def normalize_placeholder_text(text: str) -> str:
    # Removes newlines, tabs, multiple spaces
    return re.sub(r"\s+", "", text)


def parse_paragraphs(doc):
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
                "syntax": d["raw"],  # normalized syntax
                "type": d["type"],
                "scope": "body",
                "location": {
                    "paragraph_index": p_idx
                    # ❌ char_start / char_end intentionally ignored
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                }
            })

    return placeholders

def iter_block_items(parent):
    from docx.table import Table, _Cell
    from docx.text.paragraph import Paragraph

    if isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        parent_elm = parent.element.body

    for child in parent_elm.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, parent)
        elif child.tag.endswith('}tbl'):
            yield Table(child, parent)


def parse_tables(doc):
    placeholders = []

    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):

                cell_text = "".join(
                    p.text for p in cell.paragraphs
                )

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

def detect_placeholders_in_paragraph(para_text):
    found = []
    for ptype, regex in PLACEHOLDER_PATTERNS.items():
        for match in regex.finditer(para_text):
            found.append({
                "type": ptype,
                "raw": match.group(0),
                "key": match.group(1),
                "start": match.start(),
                "end": match.end()
            })
    return found


@app.route("/api/templates/upload", methods=["POST"])
def upload_template():
    file = request.files.get("file")
    template_name = request.form.get("template_name")

    if not file or not template_name:
        return jsonify({"error": "file and template_name required"}), 400

    file_bytes = file.read()
    if not file_bytes:
        return jsonify({"error": "Empty file uploaded"}), 400

    try:
        doc = Document(io.BytesIO(file_bytes))
    except Exception as e:
        return jsonify({"error": "Invalid DOCX file"}), 400

    placeholders = []
    placeholders.extend(parse_paragraphs(doc))
    placeholders.extend(parse_tables(doc))

    template_json = {
        "template_id": template_name,
        "meta": {
            "docx_hash": file_hash(file_bytes),
            "created_at": datetime.utcnow().isoformat(),
            "paragraph_count": len(doc.paragraphs),
            "table_count": len(doc.tables),
            "placeholder_count": len(placeholders)
        },
        "placeholders": placeholders
    }

    # 🔒 SAVE ONLY HERE (Mongo recommended)
    # db.templates.insert_one(template_json)

    return jsonify({
        "template_id": template_name,
        "placeholders_found": len(placeholders),
        "docx_hash": template_json["meta"]["docx_hash"],
        "status": "parsed",
        "patterns": template_json
    }), 200

def get_target_paragraph(doc, ph, trace_id):
    loc = ph["location"]

    try:
        if ph["scope"] == "body":
            para = doc.paragraphs[loc["paragraph_index"]]

        elif ph["scope"] == "table":
            table = doc.tables[loc["table_index"]]
            row = table.rows[loc["row_index"]]
            cell = row.cells[loc["cell_index"]]
            para = cell.paragraphs[loc["paragraph_index"]]

        else:
            raise ValueError("Invalid scope")

        print(f"[{trace_id}] PH={ph['id']} KEY={ph['key']}")
        print(f"[{trace_id}] Paragraph text BEFORE: {para.text}")

        return para

    except Exception as e:
        print(f"[{trace_id}] ❌ get_target_paragraph failed → {e}")
        raise

def replace_in_run(run, placeholder, value, trace_id):
    if placeholder not in run.text:
        print(f"[{trace_id}] ⚠ Placeholder NOT in run")
        return False

    print(f"[{trace_id}] ✅ Replacing in run: '{placeholder}' → '{value}'")
    run.text = run.text.replace(placeholder, value)
    return True

def replace_across_runs(paragraph, placeholder, value, trace_id):
    full_text = "".join(r.text for r in paragraph.runs)

    #print(f"[{trace_id}] Full merged text: {full_text}")

    if placeholder not in full_text:
        print(f"[{trace_id}] ❌ Placeholder NOT in merged text")
        return False

    new_text = full_text.replace(placeholder, value)
    print(f"[{trace_id}] ✅ Replaced merged text")

    # Clear all runs
    for r in paragraph.runs:
        r.text = ""

    paragraph.runs[0].text = new_text
    return True


def replace_multiline_run(run, placeholder, value):
    if placeholder not in run.text:
        return

    lines = value.split("\n")
    run.text = run.text.replace(placeholder, lines[0])

    para = run._parent
    for line in lines[1:]:
        new_run = para.add_run("\n" + line)
        new_run.bold = run.bold
        new_run.italic = run.italic
        new_run.font.name = run.font.name
        new_run.font.size = run.font.size


def replace_in_paragraph(paragraph, old, new, trace_id):
    full_text = "".join(run.text for run in paragraph.runs)

    if old not in full_text:
        return False

    print(f"[{trace_id}] 🔁 Paragraph fallback replace: {old}")

    new_text = full_text.replace(old, new)

    # preserve style of first run
    first_run = paragraph.runs[0]
    for run in paragraph.runs[1:]:
        run.text = ""

    first_run.text = new_text
    return True


def replace_in_paragraph_preserve_style(paragraph, old, new):
    full_text = "".join(run.text for run in paragraph.runs)

    if old not in full_text:
        return False

    new_text = full_text.replace(old, new)

    # clear runs
    for run in paragraph.runs:
        run.text = ""

    # put everything in first run (safe)
    paragraph.runs[0].text = new_text
    return True

def replace_in_table_cell(cell, placeholder, value, trace_id):
    """
    Safely replaces placeholder text even if it spans
    multiple paragraphs / runs inside a table cell.
    """

    # 1. Merge all paragraph texts
    original_text = "\n".join(p.text for p in cell.paragraphs)
    normalized_original = normalize_placeholder_text(original_text)
    normalized_ph = normalize_placeholder_text(placeholder)

    if normalized_ph not in normalized_original:
        print(f"[{trace_id}] CELL: placeholder not found")
        return False

    # 2. Replace in merged text
    replaced_text = original_text.replace(placeholder, value)

    print(f"[{trace_id}] CELL REPLACED: {placeholder} → {value}")

    # 3. Clear existing paragraphs
    for p in cell.paragraphs:
        p.clear()

    # 4. Rebuild cell content
    lines = replaced_text.split("\n")
    for i, line in enumerate(lines):
        if i == 0:
            cell.paragraphs[0].add_run(line)
        else:
            cell.add_paragraph(line)

    return True


def group_placeholders_by_cell(placeholders):
    grouped = defaultdict(list)

    for ph in placeholders:
        if ph["scope"] == "table":
            loc = ph["location"]
            key = (loc["table"], loc["row"], loc["cell"])
            grouped[key].append(ph)

    return grouped

def replace_all_in_cell(cell, ph_list, data, trace_id):
    replaced_any = False

    for p in cell.paragraphs:
        runs = p.runs
        if not runs:
            continue

        full_text = "".join(r.text for r in runs)
        original = full_text

        for ph in ph_list:
            value = str(data.get(ph["key"], ""))
            if ph["syntax"] in full_text:
                full_text = full_text.replace(ph["syntax"], value)

        if full_text != original:
            # clear runs safely
            for r in runs:
                r.text = ""

            runs[0].text = full_text
            replaced_any = True

            print(
                f"[{trace_id}] CELL REPLACED (GROUPED): "
                f"{len(ph_list)} placeholders"
            )

    return replaced_any


def replace_in_cell_runs(cell, target, value, trace_id):
    for p in cell.paragraphs:
        runs = p.runs
        full = "".join(r.text for r in runs)

        if target not in full:
            continue

        new_full = full.replace(target, value)

        # clear runs safely
        for r in runs:
            r.text = ""

        runs[0].text = new_full

        print(f"[{trace_id}] CELL REPLACED (RUN SAFE): {target} → {value}")
        return True

    return False

@app.route("/api/documents/generate", methods=["POST"])
def generate_document():
    start = time.time()
    trace_id = uuid.uuid4().hex[:8]

    payload = request.json
    template_id = payload["template_id"]
    data = payload["data"]

    template = template_json
    doc = Document("templates/Hos_AFS.docx")

    replaced = 0

    # 1. Group table placeholders by (table, row, col)
    table_groups = group_placeholders_by_cell(template["placeholders"])

    # 2. Replace BODY placeholders
    for ph in template["placeholders"]:
        if ph["scope"] != "body":
            continue

        try:
            para = get_target_paragraph(doc, ph, trace_id)
            value = str(data.get(ph["key"], ""))
            if replace_in_paragraph(para, ph["syntax"], value, trace_id):
                replaced += 1
        except Exception:
            continue

    # 3. Replace TABLE placeholders
    for (t, r, c), ph_list in table_groups.items():
        cell = doc.tables[t].rows[r].cells[c]
        if replace_all_in_cell(cell, ph_list, data, trace_id):
            replaced += len(ph_list)

    output_path = f"outputs/{template_id}_{int(time.time())}.docx"
    doc.save(output_path)

    print(f"[{trace_id}] DONE: {replaced}/{len(template['placeholders'])}")

    return jsonify({
        "status": "success",
        "replaced": replaced,
        "total": len(template["placeholders"]),
        "output": output_path,
        "generation_time_ms": int((time.time() - start) * 1000),
        "trace_id": trace_id
    })

# test

template_json = {        "placeholders": [
            {
                "id": "ph_011a96d8",
                "key": "Godrej_Green_Estate_AFS1st_English",
                "location": {
                    "paragraph_index": 43
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "$<$Godrej_Green_Estate_AFS1st_English$>$",
                "type": "clause_block"
            },
            {
                "id": "ph_260bbede",
                "key": "Godrej_Green_Estate_AFS1st_English",
                "location": {
                    "paragraph_index": 43
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<$Godrej_Green_Estate_AFS1st_English$>",
                "type": "single_clause_block"
            },
            {
                "id": "ph_c808262c",
                "key": "Godrej_Green_Estate_AFS_2nd_English",
                "location": {
                    "paragraph_index": 44
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "$<$Godrej_Green_Estate_AFS_2nd_English$>$",
                "type": "clause_block"
            },
            {
                "id": "ph_72435576",
                "key": "Godrej_Green_Estate_AFS_2nd_English",
                "location": {
                    "paragraph_index": 44
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<$Godrej_Green_Estate_AFS_2nd_English$>",
                "type": "single_clause_block"
            },
            {
                "id": "ph_d9e1a5e1",
                "key": "Godrej_Green_Estate_AFS_3rd_English",
                "location": {
                    "paragraph_index": 45
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "$<$Godrej_Green_Estate_AFS_3rd_English$>$",
                "type": "clause_block"
            },
            {
                "id": "ph_fdab724b",
                "key": "Godrej_Green_Estate_AFS_3rd_English",
                "location": {
                    "paragraph_index": 45
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<$Godrej_Green_Estate_AFS_3rd_English$>",
                "type": "single_clause_block"
            },
            {
                "id": "ph_493f6fd0",
                "key": "Godrej_Green_Estate_AFS_4th_English",
                "location": {
                    "paragraph_index": 46
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "$<$Godrej_Green_Estate_AFS_4th_English$>$",
                "type": "clause_block"
            },
            {
                "id": "ph_993942d1",
                "key": "Godrej_Green_Estate_AFS_4th_English",
                "location": {
                    "paragraph_index": 46
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<$Godrej_Green_Estate_AFS_4th_English$>",
                "type": "single_clause_block"
            },
            {
                "id": "ph_a41e5eb6",
                "key": "propstrength__application_booking__c.propstrength__tower__c.rera_registration_number__c",
                "location": {
                    "paragraph_index": 95
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__tower__c.rera_registration_number__c>>",
                "type": "angle"
            },
            {
                "id": "ph_a089037e",
                "key": "propstrength__application_booking__c.propstrength__property__c.propstrength__house_unit_no__c",
                "location": {
                    "paragraph_index": 109
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.propstrength__house_unit_no__c>>",
                "type": "angle"
            },
            {
                "id": "ph_2ca32b09",
                "key": "propstrength__application_booking__c.propstrength__property__c.floor_name__c",
                "location": {
                    "paragraph_index": 109
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.floor_name__c>>",
                "type": "angle"
            },
            {
                "id": "ph_5482033c",
                "key": "propstrength__application_booking__c.propstrength__property__c.tower_name__c",
                "location": {
                    "paragraph_index": 109
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.tower_name__c>>",
                "type": "angle"
            },
            {
                "id": "ph_c53b7d15",
                "key": "propstrength__application_booking__c.propstrength__cp_number_of_parking_purchased__c",
                "location": {
                    "paragraph_index": 109
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__cp_number_of_parking_purchased__c>>",
                "type": "angle"
            },
            {
                "id": "ph_831d8631",
                "key": "propstrength__application_booking__c.propstrength__property__c.open_balc_sq_mt__c",
                "location": {
                    "paragraph_index": 111
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.open_balc_sq_mt__c>>",
                "type": "angle"
            },
            {
                "id": "ph_1325d012",
                "key": "propstrength__application_booking__c.propstrength__property__c.open_balc_sq_ft__c",
                "location": {
                    "paragraph_index": 111
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.open_balc_sq_ft__c>>",
                "type": "angle"
            },
            {
                "id": "ph_d09ad143",
                "key": "propstrength__application_booking__c.propstrength__property__c.appurtenant_area_sq_mt__c",
                "location": {
                    "paragraph_index": 111
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.appurtenant_area_sq_mt__c>>",
                "type": "angle"
            },
            {
                "id": "ph_9ab0e187",
                "key": "propstrength__application_booking__c.propstrength__property__c.appurtenant_area_sq_ft__c",
                "location": {
                    "paragraph_index": 111
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.appurtenant_area_sq_ft__c>>",
                "type": "angle"
            },
            {
                "id": "ph_cf83c5bf",
                "key": "propstrength__application_booking__c.total_payment_realized__c",
                "location": {
                    "paragraph_index": 117
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<$propstrength__application_booking__c.total_payment_realized__c$>",
                "type": "single_clause_block"
            },
            {
                "id": "ph_0f9ff051",
                "key": "propstrength__application_booking__c.total_payment_realized__c",
                "location": {
                    "paragraph_index": 117
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.total_payment_realized__c>>",
                "type": "angle"
            },
            {
                "id": "ph_8b616482",
                "key": "propstrength__application_booking__c.propstrength__property__c.open_balc_sq_mt__c",
                "location": {
                    "paragraph_index": 138
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.open_balc_sq_mt__c>>",
                "type": "angle"
            },
            {
                "id": "ph_b2b5306c",
                "key": "propstrength__application_booking__c.propstrength__property__c.open_balc_sq_ft__c",
                "location": {
                    "paragraph_index": 138
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.open_balc_sq_ft__c>>",
                "type": "angle"
            },
            {
                "id": "ph_2b31f74f",
                "key": "propstrength__application_booking__c.propstrength__property__c.appurtenant_area_sq_mt__c",
                "location": {
                    "paragraph_index": 138
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.appurtenant_area_sq_mt__c>>",
                "type": "angle"
            },
            {
                "id": "ph_0dcb311c",
                "key": "propstrength__application_booking__c.propstrength__property__c.appurtenant_area_sq_ft__c",
                "location": {
                    "paragraph_index": 138
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.appurtenant_area_sq_ft__c>>",
                "type": "angle"
            },
            {
                "id": "ph_6aa91d4c",
                "key": "propstrength__application_booking__c.propstrength__property__c.fsi_area_sq_mt__c",
                "location": {
                    "paragraph_index": 138
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.fsi_area_sq_mt__c>>",
                "type": "angle"
            },
            {
                "id": "ph_9645d023",
                "key": "propstrength__application_booking__c.propstrength__property__c.fsi_area_sq_ft__c",
                "location": {
                    "paragraph_index": 138
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.fsi_area_sq_ft__c>>",
                "type": "angle"
            },
            {
                "id": "ph_c82003c1",
                "key": "propstrength__application_booking__c.propstrength__revised_agreement_amount__c",
                "location": {
                    "paragraph_index": 140
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<$propstrength__application_booking__c.propstrength__revised_agreement_amount__c$>",
                "type": "single_clause_block"
            },
            {
                "id": "ph_c063c10b",
                "key": "propstrength__application_booking__c.propstrength__revised_agreement_amount__c",
                "location": {
                    "paragraph_index": 140
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__revised_agreement_amount__c>>",
                "type": "angle"
            },
            {
                "id": "ph_0df9839c",
                "key": "propstrength__application_booking__c.total_payment_realized__c",
                "location": {
                    "paragraph_index": 144
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<$propstrength__application_booking__c.total_payment_realized__c$>",
                "type": "single_clause_block"
            },
            {
                "id": "ph_7e24bd08",
                "key": "propstrength__application_booking__c.total_payment_realized__c",
                "location": {
                    "paragraph_index": 144
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.total_payment_realized__c>>",
                "type": "angle"
            },
            {
                "id": "ph_f8d55458",
                "key": "propstrength__application_booking__c.propstrength__property__c.propstrength__unit_type__c",
                "location": {
                    "paragraph_index": 862
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.propstrength__unit_type__c>>",
                "type": "angle"
            },
            {
                "id": "ph_bc8ad16e",
                "key": "propstrength__application_booking__c.propstrength__property__c.propstrength__house_unit_no__c",
                "location": {
                    "paragraph_index": 862
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.propstrength__house_unit_no__c>>",
                "type": "angle"
            },
            {
                "id": "ph_167ce1ac",
                "key": "propstrength__application_booking__c.propstrength__property__c.floor_name__c",
                "location": {
                    "paragraph_index": 862
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.floor_name__c>>",
                "type": "angle"
            },
            {
                "id": "ph_cba512e2",
                "key": "propstrength__application_booking__c.propstrength__property__c.tower_name__c",
                "location": {
                    "paragraph_index": 862
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.tower_name__c>>",
                "type": "angle"
            },
            {
                "id": "ph_28213462",
                "key": "propstrength__application_booking__c.propstrength__property__c.open_balc_sq_mt__c",
                "location": {
                    "paragraph_index": 862
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.open_balc_sq_mt__c>>",
                "type": "angle"
            },
            {
                "id": "ph_d7d99a77",
                "key": "propstrength__application_booking__c.propstrength__property__c.appurtenant_area_sq_mt__c",
                "location": {
                    "paragraph_index": 862
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.appurtenant_area_sq_mt__c>>",
                "type": "angle"
            },
            {
                "id": "ph_956c245f",
                "key": "propstrength__application_booking__c.propstrength__property__c.build_common_area__c",
                "location": {
                    "paragraph_index": 862
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.build_common_area__c>>",
                "type": "angle"
            },
            {
                "id": "ph_99ecff29",
                "key": "propstrength__application_booking__c.propstrength__cp_number_of_parking_purchased__c",
                "location": {
                    "paragraph_index": 862
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__cp_number_of_parking_purchased__c>>",
                "type": "angle"
            },
            {
                "id": "ph_d9d93233",
                "key": "propstrength__property__c.bizimagepath",
                "location": {
                    "paragraph_index": 915
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "{{propstrength__property__c.bizimagepath}}",
                "type": "text"
            },
            {
                "id": "ph_edfc5930",
                "key": "propstrength__application_booking__c.propstrength__revised_agreement_amount__c",
                "location": {
                    "paragraph_index": 946
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<$propstrength__application_booking__c.propstrength__revised_agreement_amount__c$>",
                "type": "single_clause_block"
            },
            {
                "id": "ph_996aafe1",
                "key": "propstrength__application_booking__c.propstrength__revised_agreement_amount__c",
                "location": {
                    "paragraph_index": 946
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "body",
                "syntax": "<<propstrength__application_booking__c.propstrength__revised_agreement_amount__c>>",
                "type": "angle"
            },
            {
                "id": "ph_e3477ba7",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.salutation",
                "location": {
                    "cell": 0,
                    "row": 1,
                    "table": 3
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.salutation>>",
                "type": "angle"
            },
            {
                "id": "ph_a08a3db3",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.name",
                "location": {
                    "cell": 0,
                    "row": 1,
                    "table": 3
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.name>>",
                "type": "angle"
            },
            {
                "id": "ph_68af3665",
                "key": "propstrength__application_booking__c.propstrength__total_basic_sales_price__c",
                "location": {
                    "cell": 2,
                    "row": 1,
                    "table": 4
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__total_basic_sales_price__c>>",
                "type": "angle"
            },
            {
                "id": "ph_0c41834d",
                "key": "propstrength__application_booking__c.propstrength__other_charges_opted__c.INFRACHARGES",
                "location": {
                    "cell": 2,
                    "row": 2,
                    "table": 4
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__other_charges_opted__c.INFRACHARGES>>",
                "type": "angle"
            },
            {
                "id": "ph_6ddf6eb3",
                "key": "propstrength__application_booking__c.propstrength__other_charges_opted__c.AdvancemaintenanceCharges",
                "location": {
                    "cell": 2,
                    "row": 1,
                    "table": 5
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__other_charges_opted__c.AdvancemaintenanceCharges>>",
                "type": "angle"
            },
            {
                "id": "ph_f51ded89",
                "key": "propstrength__application_booking__c.propstrength__other_charges_opted__c.SinkingFundDeposit",
                "location": {
                    "cell": 2,
                    "row": 2,
                    "table": 5
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__other_charges_opted__c.SinkingFundDeposit>>",
                "type": "angle"
            },
            {
                "id": "ph_60f8888b",
                "key": "ParkRP2",
                "location": {
                    "cell": 2,
                    "row": 3,
                    "table": 5
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<ParkRP2>>",
                "type": "angle"
            },
            {
                "id": "ph_a5e3da7d",
                "key": "srno",
                "location": {
                    "cell": 0,
                    "row": 1,
                    "table": 6
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<srno>>",
                "type": "angle"
            },
            {
                "id": "ph_f480cae1",
                "key": "msname",
                "location": {
                    "cell": 1,
                    "row": 1,
                    "table": 6
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<msname>>",
                "type": "angle"
            },
            {
                "id": "ph_e5bbf56b",
                "key": "amtpertage",
                "location": {
                    "cell": 2,
                    "row": 1,
                    "table": 6
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<amtpertage>>",
                "type": "angle"
            },
            {
                "id": "ph_16becf3a",
                "key": "ttlamtdueCOP",
                "location": {
                    "cell": 3,
                    "row": 1,
                    "table": 6
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<ttlamtdueCOP>>",
                "type": "angle"
            },
            {
                "id": "ph_197ea058",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.salutation",
                "location": {
                    "cell": 0,
                    "row": 1,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.salutation>>",
                "type": "angle"
            },
            {
                "id": "ph_a79a4706",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.name",
                "location": {
                    "cell": 0,
                    "row": 1,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.name>>",
                "type": "angle"
            },
            {
                "id": "ph_96dd6b3a",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.salutation",
                "location": {
                    "cell": 0,
                    "row": 1,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.salutation>>",
                "type": "angle"
            },
            {
                "id": "ph_e3a9445a",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.name",
                "location": {
                    "cell": 0,
                    "row": 1,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.name>>",
                "type": "angle"
            },
            {
                "id": "ph_007668ac",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.salutation",
                "location": {
                    "cell": 0,
                    "row": 1,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.salutation>>",
                "type": "angle"
            },
            {
                "id": "ph_15be056f",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.name",
                "location": {
                    "cell": 0,
                    "row": 1,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.name>>",
                "type": "angle"
            },
            {
                "id": "ph_7df9d11e",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.salutation",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.salutation>>",
                "type": "angle"
            },
            {
                "id": "ph_70d5b6fa",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.name",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.name>>",
                "type": "angle"
            },
            {
                "id": "ph_31245560",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.propstrength__income_tax_permanent_account_no__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.propstrength__income_tax_permanent_account_no__c>>",
                "type": "angle"
            },
            {
                "id": "ph_01ef37be",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.relationship__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.relationship__c>>",
                "type": "angle"
            },
            {
                "id": "ph_1834b2d2",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.relationship_name__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.relationship_name__c>>",
                "type": "angle"
            },
            {
                "id": "ph_71c963b4",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.age__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.age__c>>",
                "type": "angle"
            },
            {
                "id": "ph_1b5421fb",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.residential_street1__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.residential_street1__c>>",
                "type": "angle"
            },
            {
                "id": "ph_2e3859f9",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.residential_street2__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.residential_street2__c>>",
                "type": "angle"
            },
            {
                "id": "ph_e880aeda",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.residential_street3__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.residential_street3__c>>",
                "type": "angle"
            },
            {
                "id": "ph_48090c66",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.residential_city__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.residential_city__c>>",
                "type": "angle"
            },
            {
                "id": "ph_9bcec5e9",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.residential_state__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.residential_state__c>>",
                "type": "angle"
            },
            {
                "id": "ph_ea54ed67",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.residential_post_code__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.residential_post_code__c>>",
                "type": "angle"
            },
            {
                "id": "ph_2786b63c",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.residential_country__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.residential_country__c>>",
                "type": "angle"
            },
            {
                "id": "ph_ce51d1ca",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.salutation",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.salutation>>",
                "type": "angle"
            },
            {
                "id": "ph_73a8ca1c",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.name",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.name>>",
                "type": "angle"
            },
            {
                "id": "ph_7174a7ae",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.propstrength__income_tax_permanent_account_no__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.propstrength__income_tax_permanent_account_no__c>>",
                "type": "angle"
            },
            {
                "id": "ph_6d8e6cef",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.relationship__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.relationship__c>>",
                "type": "angle"
            },
            {
                "id": "ph_954e62f7",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.relationship_name__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.relationship_name__c>>",
                "type": "angle"
            },
            {
                "id": "ph_9f5dcbb5",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.age__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.age__c>>",
                "type": "angle"
            },
            {
                "id": "ph_5a9e0f27",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.residential_street1__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.residential_street1__c>>",
                "type": "angle"
            },
            {
                "id": "ph_c19e288c",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.residential_street2__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.residential_street2__c>>",
                "type": "angle"
            },
            {
                "id": "ph_4cc559dd",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.residential_street3__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.residential_street3__c>>",
                "type": "angle"
            },
            {
                "id": "ph_523d6d0d",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.residential_city__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.residential_city__c>>",
                "type": "angle"
            },
            {
                "id": "ph_08c7eaa0",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.residential_state__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.residential_state__c>>",
                "type": "angle"
            },
            {
                "id": "ph_f52b68d4",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.residential_post_code__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.residential_post_code__c>>",
                "type": "angle"
            },
            {
                "id": "ph_baf7eeb9",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.residential_country__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.residential_country__c>>",
                "type": "angle"
            },
            {
                "id": "ph_cdf59a81",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.salutation",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.salutation>>",
                "type": "angle"
            },
            {
                "id": "ph_13f263ac",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.name",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.name>>",
                "type": "angle"
            },
            {
                "id": "ph_19c28ee1",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.propstrength__income_tax_permanent_account_no__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.propstrength__income_tax_permanent_account_no__c>>",
                "type": "angle"
            },
            {
                "id": "ph_e5493228",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.relationship__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.relationship__c>>",
                "type": "angle"
            },
            {
                "id": "ph_d5ea0651",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.relationship_name__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.relationship_name__c>>",
                "type": "angle"
            },
            {
                "id": "ph_b14d0afb",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.age__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.age__c>>",
                "type": "angle"
            },
            {
                "id": "ph_5fb8a6dc",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.residential_street1__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.residential_street1__c>>",
                "type": "angle"
            },
            {
                "id": "ph_6aa931f4",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.residential_street2__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.residential_street2__c>>",
                "type": "angle"
            },
            {
                "id": "ph_e32231a8",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.residential_street3__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.residential_street3__c>>",
                "type": "angle"
            },
            {
                "id": "ph_0ffa0bef",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.residential_city__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.residential_city__c>>",
                "type": "angle"
            },
            {
                "id": "ph_d3983cc3",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.residential_state__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.residential_state__c>>",
                "type": "angle"
            },
            {
                "id": "ph_501266d4",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.residential_post_code__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.residential_post_code__c>>",
                "type": "angle"
            },
            {
                "id": "ph_9e00211e",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.residential_country__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.residential_country__c>>",
                "type": "angle"
            },
            {
                "id": "ph_9bf271c5",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.salutation",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.salutation>>",
                "type": "angle"
            },
            {
                "id": "ph_2f43c3c3",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.name",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.name>>",
                "type": "angle"
            },
            {
                "id": "ph_24ea3b8e",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.propstrength__income_tax_permanent_account_no__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.propstrength__income_tax_permanent_account_no__c>>",
                "type": "angle"
            },
            {
                "id": "ph_f94a761c",
                "key": "propstrength__application_booking__c.propstrength__property__c.propstrength__unit_type__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.propstrength__unit_type__c>>",
                "type": "angle"
            },
            {
                "id": "ph_a10cbc76",
                "key": "propstrength__application_booking__c.propstrength__property__c.propstrength__house_unit_no__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.propstrength__house_unit_no__c>>",
                "type": "angle"
            },
            {
                "id": "ph_0cbdc7c2",
                "key": "propstrength__application_booking__c.propstrength__property__c.floor_name__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.floor_name__c>>",
                "type": "angle"
            },
            {
                "id": "ph_9124742a",
                "key": "propstrength__application_booking__c.propstrength__property__c.tower_name__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.tower_name__c>>",
                "type": "angle"
            },
            {
                "id": "ph_db76b80d",
                "key": "propstrength__application_booking__c.propstrength__property__c.open_balc_sq_mt__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.open_balc_sq_mt__c>>",
                "type": "angle"
            },
            {
                "id": "ph_113647ea",
                "key": "propstrength__application_booking__c.propstrength__property__c.appurtenant_area_sq_mt__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.appurtenant_area_sq_mt__c>>",
                "type": "angle"
            },
            {
                "id": "ph_55fed5ab",
                "key": "propstrength__application_booking__c.propstrength__property__c.build_common_area__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__property__c.build_common_area__c>>",
                "type": "angle"
            },
            {
                "id": "ph_bf26a677",
                "key": "propstrength__application_booking__c.propstrength__cp_number_of_parking_purchased__c",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__cp_number_of_parking_purchased__c>>",
                "type": "angle"
            },
            {
                "id": "ph_7935bfdf",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.salutation",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.salutation>>",
                "type": "angle"
            },
            {
                "id": "ph_136d4da0",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.name",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.1stapplicant.name>>",
                "type": "angle"
            },
            {
                "id": "ph_21aa3971",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.salutation",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.salutation>>",
                "type": "angle"
            },
            {
                "id": "ph_092dac8d",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.name",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.2ndapplicant.name>>",
                "type": "angle"
            },
            {
                "id": "ph_e455d224",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.salutation",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.salutation>>",
                "type": "angle"
            },
            {
                "id": "ph_92687bdf",
                "key": "propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.name",
                "location": {
                    "cell": 0,
                    "row": 2,
                    "table": 10
                },
                "rules": {
                    "allow_multiline": True,
                    "preserve_style": True
                },
                "scope": "table",
                "syntax": "<<propstrength__application_booking__c.propstrength__applicant_detail__c.3rdapplicant.name>>",
                "type": "angle"
            }
        ]}


if __name__ == "__main__":
    app.run(debug=True)