

def group_placeholders_by_cell(placeholders):
    grouped = defaultdict(list)

    for ph in placeholders:
        if ph["scope"] == "table":
            loc = ph["location"]
            key = (loc["table"], loc["row"], loc["cell"])
            grouped[key].append(ph)

    return grouped


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








