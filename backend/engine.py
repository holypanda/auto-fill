"""Core engine for filling Word/PDF document templates."""

import json
import os
import re
import zipfile
import copy
from io import BytesIO
from pathlib import Path
from lxml import etree
from pypdf import PdfReader, PdfWriter
from pypdf.generic import BooleanObject, NameObject

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": WORD_NS}


def _is_xml_part(name: str) -> bool:
    """Check if a ZIP entry is an XML part that may contain text."""
    return name.endswith(".xml") and not name.startswith("docProps/") and name != "[Content_Types].xml"


def _replace_in_paragraph(p_elem, data: dict) -> bool:
    """Replace placeholders in a single w:p element.

    Strategy: build a char-position map to trace each character back to its
    source run. Only merge/modify runs that participate in a placeholder span.
    Runs with no placeholder involvement are left completely untouched.
    """
    runs = p_elem.findall(".//w:r", NSMAP)
    if not runs:
        return False

    # Build list of (run_element, w:t_element, text) preserving order
    run_info = []  # [(run_elem, t_elem, original_text), ...]
    for r in runs:
        t_list = r.findall("w:t", NSMAP)
        if t_list:
            # A run may have multiple w:t (rare) — treat as one
            combined = "".join((t.text or "") for t in t_list)
            run_info.append((r, t_list, combined))

    if not run_info:
        return False

    full_text = "".join(info[2] for info in run_info)

    # Check if any placeholder exists
    if not re.search(r"\{[^{}]+\}", full_text):
        return False

    # Do replacements on the full text
    new_text = full_text
    for key, value in data.items():
        if key in new_text:
            new_text = new_text.replace(key, str(value))

    if new_text == full_text:
        return False

    # Build a char→run_index map for the ORIGINAL text
    char_to_run = []
    for idx, (_, _, text) in enumerate(run_info):
        char_to_run.extend([idx] * len(text))

    # Find which runs are touched by placeholders (in original text)
    touched_runs = set()
    for m in re.finditer(r"\{[^{}]+\}", full_text):
        for pos in range(m.start(), m.end()):
            touched_runs.add(char_to_run[pos])

    # Group consecutive touched runs into merge spans
    # For each span, we need to figure out what the replacement text is.
    # Approach: split the original text into segments (touched-group vs untouched-run)
    # and match against the new text.

    # Build segments: list of (type, run_indices, original_text)
    segments = []
    i = 0
    while i < len(run_info):
        if i in touched_runs:
            # Start a group of consecutive touched runs
            group_indices = []
            while i < len(run_info) and i in touched_runs:
                group_indices.append(i)
                i += 1
            group_text = "".join(run_info[idx][2] for idx in group_indices)
            segments.append(("touched", group_indices, group_text))
        else:
            segments.append(("untouched", [i], run_info[i][2]))
            i += 1

    # Now reconstruct the new text for each segment.
    # Untouched segments keep their original text. Touched segments get
    # their portion of the replacement. We find touched portions by
    # replacing the untouched anchors to locate boundaries.
    # Simpler approach: replay replacements on each touched group independently.
    for seg in segments:
        if seg[0] == "untouched":
            continue
        group_indices = seg[1]
        group_orig = seg[2]
        group_new = group_orig
        for key, value in data.items():
            if key in group_new:
                group_new = group_new.replace(key, str(value))

        # Write the replaced text into the first run of the group, clear the rest
        first_idx = group_indices[0]
        first_run, first_t_list, _ = run_info[first_idx]
        first_t_list[0].text = group_new
        first_t_list[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        # Clear remaining w:t in first run
        for t in first_t_list[1:]:
            t.text = ""

        # Remove subsequent runs in the group from the paragraph
        for idx in group_indices[1:]:
            rm_run = run_info[idx][0]
            parent = rm_run.getparent()
            if parent is not None:
                parent.remove(rm_run)

    return True


def _process_xml(xml_bytes: bytes, data: dict) -> bytes | None:
    """Process an XML part, replacing placeholders. Returns modified bytes or None if unchanged."""
    try:
        tree = etree.fromstring(xml_bytes)
    except etree.XMLSyntaxError:
        return None

    changed = False
    for p in tree.iter(f"{{{WORD_NS}}}p"):
        if _replace_in_paragraph(p, data):
            changed = True

    if not changed:
        return None

    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def fill_template(template_path: str, data: dict, output_path: str) -> dict:
    """Fill a single template (.docx or .pdf) with data. Returns dict with path and warnings."""
    if template_path.lower().endswith(".pdf"):
        return fill_pdf_template(template_path, data, output_path)

    warnings = []
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    with zipfile.ZipFile(template_path, "r") as zin:
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                raw = zin.read(item.filename)
                if _is_xml_part(item.filename):
                    result = _process_xml(raw, data)
                    if result is not None:
                        zout.writestr(item, result)
                        continue
                zout.writestr(item, raw)

    # Check for remaining placeholders in output
    remaining = scan_placeholders(output_path)
    if remaining:
        warnings = [f"Unfilled placeholder: {p}" for p in remaining]

    return {"path": output_path, "warnings": warnings}


def fill_all_templates(template_dir: str, data: dict, output_dir: str, client_name: str = "", template_names: list[str] | None = None) -> list[dict]:
    """Fill all .docx templates in a directory. Returns list of result dicts."""
    os.makedirs(output_dir, exist_ok=True)
    results = []

    for fname in sorted(os.listdir(template_dir)):
        ext_lower = fname.lower()
        if not (ext_lower.endswith(".docx") or ext_lower.endswith(".pdf")):
            continue
        if template_names is not None and fname not in template_names:
            continue
        template_path = os.path.join(template_dir, fname)
        base, ext = os.path.splitext(fname)
        suffix = f"_filled_{client_name}" if client_name else "_filled"
        output_name = f"{base}{suffix}{ext}"
        output_path = os.path.join(output_dir, output_name)

        result = fill_template(template_path, data, output_path)
        result["template"] = fname
        results.append(result)

    return results


def scan_placeholders(template_path: str) -> list[str]:
    """Scan a .docx or .pdf file for placeholders. Returns sorted unique list."""
    if template_path.lower().endswith(".pdf"):
        return scan_pdf_placeholders(template_path)

    placeholders = set()

    with zipfile.ZipFile(template_path, "r") as z:
        for name in z.namelist():
            if not _is_xml_part(name):
                continue
            try:
                tree = etree.fromstring(z.read(name))
            except etree.XMLSyntaxError:
                continue

            for p in tree.iter(f"{{{WORD_NS}}}p"):
                runs = p.findall(".//w:r", NSMAP)
                t_elements = []
                for r in runs:
                    for t in r.findall("w:t", NSMAP):
                        t_elements.append(t)
                if not t_elements:
                    continue
                full_text = "".join((t.text or "") for t in t_elements)
                for m in re.finditer(r"\{[^{}]+\}", full_text):
                    placeholders.add(m.group())

    return sorted(placeholders)


# ─── PDF support ─────────────────────────────────────────────────────────────

def _load_pdf_field_map(template_path: str) -> dict:
    """Load the JSON field map for a PDF template.

    Looks for a .json file with the same base name next to the PDF.
    e.g. fw8bene.pdf → fw8bene.json
    The JSON maps our {placeholder} keys to PDF form field names.
    """
    map_path = Path(template_path).with_suffix(".json")
    if not map_path.exists():
        return {}
    with open(map_path, "r", encoding="utf-8") as f:
        return json.load(f)


def scan_pdf_placeholders(template_path: str) -> list[str]:
    """Return the placeholder keys that this PDF template accepts.

    Reads from the companion .json field map.
    """
    field_map = _load_pdf_field_map(template_path)
    return sorted(field_map.keys())


def _set_need_appearances(writer: PdfWriter):
    """Tell PDF readers to regenerate field appearances from values (required for XFA forms)."""
    if "/AcroForm" in writer._root_object:
        writer._root_object["/AcroForm"][NameObject("/NeedAppearances")] = BooleanObject(True)


def fill_pdf_template(template_path: str, data: dict, output_path: str) -> dict:
    """Fill a PDF form template using a field map. Returns dict with path and warnings."""
    warnings = []
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    field_map = _load_pdf_field_map(template_path)
    if not field_map:
        warnings.append("No field map (.json) found for this PDF — nothing filled")
        # Just copy the file
        import shutil
        shutil.copy2(template_path, output_path)
        return {"path": output_path, "warnings": warnings}

    reader = PdfReader(template_path)
    writer = PdfWriter()
    writer.append(reader)

    # Build PDF field name → value mapping
    pdf_values = {}
    mapped_keys = set()
    for placeholder_key, pdf_field in field_map.items():
        if placeholder_key in data and data[placeholder_key]:
            pdf_values[pdf_field] = data[placeholder_key]
            mapped_keys.add(placeholder_key)

    # Fill each page
    for page in writer.pages:
        writer.update_page_form_field_values(page, pdf_values)

    # Force PDF readers to regenerate field appearances from values
    _set_need_appearances(writer)

    with open(output_path, "wb") as f:
        writer.write(f)

    # Warn about unmapped placeholders
    for key in field_map:
        if key not in data:
            warnings.append(f"Missing data for: {key}")
        elif not data[key]:
            warnings.append(f"Empty value for: {key}")

    return {"path": output_path, "warnings": warnings}
