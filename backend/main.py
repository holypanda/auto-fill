"""FastAPI backend for auto-fill document system."""

import os
import io
import json
import zipfile
import shutil
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, StreamingResponse, HTMLResponse
from pydantic import BaseModel

from backend.engine import fill_template, fill_all_templates, scan_placeholders

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
FRONTEND_DIR = BASE_DIR / "frontend"

app = FastAPI(title="Auto-Fill Document System")


@app.get("/", response_class=HTMLResponse)
async def index():
    html_path = FRONTEND_DIR / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


SUPPORTED_EXT = {".docx", ".pdf"}


def _is_template(f: Path) -> bool:
    return f.suffix.lower() in SUPPORTED_EXT


@app.get("/api/templates")
async def list_templates():
    """List all templates and their placeholders."""
    results = []
    for f in sorted(TEMPLATE_DIR.iterdir()):
        if _is_template(f):
            placeholders = scan_placeholders(str(f))
            results.append({"name": f.name, "placeholders": placeholders, "type": f.suffix.lower()})
    return results


@app.get("/api/fields")
async def get_fields():
    """Return all unique placeholder fields across all templates."""
    all_fields = set()
    for f in TEMPLATE_DIR.iterdir():
        if _is_template(f):
            all_fields.update(scan_placeholders(str(f)))
    return sorted(all_fields)


class FillRequest(BaseModel):
    data: dict[str, str]
    client_name: str = ""
    template_names: list[str] = []


@app.post("/api/fill")
async def fill_all(req: FillRequest):
    """Fill all templates and return as ZIP."""
    # Clean output dir
    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    OUTPUT_DIR.mkdir(exist_ok=True)

    client_name = req.client_name or req.data.get("{Entity Name}", "client")
    selected = req.template_names if req.template_names else None
    results = fill_all_templates(str(TEMPLATE_DIR), req.data, str(OUTPUT_DIR), client_name, template_names=selected)

    # Create ZIP in memory
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in results:
            path = r["path"]
            zf.write(path, os.path.basename(path))
    buf.seek(0)

    # Collect all warnings
    all_warnings = []
    for r in results:
        for w in r.get("warnings", []):
            all_warnings.append(f"{r['template']}: {w}")

    safe_name = client_name.replace(" ", "_").replace("/", "_")
    return StreamingResponse(
        buf,
        media_type="application/zip",
        headers={
            "Content-Disposition": f'attachment; filename="{safe_name}_documents.zip"',
            "X-Warnings": json.dumps(all_warnings),
        },
    )


class FillSingleRequest(BaseModel):
    data: dict[str, str]
    template_name: str
    client_name: str = ""


@app.post("/api/fill-single")
async def fill_single(req: FillSingleRequest):
    """Fill a single template and return the docx."""
    template_path = TEMPLATE_DIR / req.template_name
    if not template_path.exists():
        raise HTTPException(404, f"Template not found: {req.template_name}")

    OUTPUT_DIR.mkdir(exist_ok=True)
    client_name = req.client_name or req.data.get("{Entity Name}", "client")
    base, ext = os.path.splitext(req.template_name)
    safe_name = client_name.replace(" ", "_").replace("/", "_")
    output_name = f"{base}_filled_{safe_name}{ext}"
    output_path = OUTPUT_DIR / output_name

    result = fill_template(str(template_path), req.data, str(output_path))

    if ext.lower() == ".pdf":
        media = "application/pdf"
    else:
        media = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    return FileResponse(str(output_path), media_type=media, filename=output_name)


@app.get("/api/download-template/{template_name}")
async def download_template(template_name: str):
    """Download the original template file."""
    template_path = TEMPLATE_DIR / template_name
    if not template_path.exists() or not _is_template(template_path):
        raise HTTPException(404, "Template not found")
    return FileResponse(str(template_path), filename=template_name)


@app.delete("/api/templates/{template_name}")
async def delete_template(template_name: str):
    """Delete a template and its associated metadata files."""
    template_path = TEMPLATE_DIR / template_name
    if not template_path.exists() or not _is_template(template_path):
        raise HTTPException(404, "Template not found")

    template_path.unlink()

    # Clean up associated JSON files for PDF templates
    if template_path.suffix.lower() == ".pdf":
        stem = Path(template_name).stem
        for suffix_file in [
            TEMPLATE_DIR / f"{stem}.json",        # field mapping
            TEMPLATE_DIR / f"{stem}_desc.json",    # descriptions
        ]:
            if suffix_file.exists():
                suffix_file.unlink()

    return {"status": "ok", "deleted": template_name}


@app.post("/api/upload-template")
async def upload_template(file: UploadFile = File(...)):
    """Upload a new template file."""
    ext = Path(file.filename).suffix.lower()
    if ext not in SUPPORTED_EXT:
        raise HTTPException(400, "Only .docx and .pdf files are allowed")

    dest = TEMPLATE_DIR / file.filename
    content = await file.read()
    dest.write_bytes(content)

    placeholders = scan_placeholders(str(dest))
    return {"name": file.filename, "placeholders": placeholders}


def _load_pdf_descriptions(template_name: str) -> dict:
    """Load human-readable descriptions for PDF fields from a *_desc.json file."""
    desc_path = (TEMPLATE_DIR / template_name).with_suffix("")
    # e.g. fw8bene.pdf → fw8bene_desc.json
    desc_file = TEMPLATE_DIR / (Path(template_name).stem + "_desc.json")
    if desc_file.exists():
        with open(desc_file, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


@app.get("/api/pdf-fields/{template_name}")
async def get_pdf_form_fields(template_name: str):
    """List all form fields in a PDF template (for building field mappings)."""
    template_path = TEMPLATE_DIR / template_name
    if not template_path.exists() or template_path.suffix.lower() != ".pdf":
        raise HTTPException(404, "PDF template not found")

    from pypdf import PdfReader
    reader = PdfReader(str(template_path))
    fields = reader.get_fields() or {}
    descriptions = _load_pdf_descriptions(template_name)

    result = []
    for name, field in sorted(fields.items()):
        ft = field.get("/FT", "")
        val = field.get("/V", "")
        if ft == "/Tx":
            result.append({
                "name": name,
                "type": "text",
                "value": str(val),
                "description": descriptions.get(name, ""),
            })
        elif ft == "/Btn":
            result.append({
                "name": name,
                "type": "checkbox",
                "value": str(val),
                "description": descriptions.get(name, ""),
            })
    return result


@app.get("/api/pdf-preview-labeled/{template_name}")
async def pdf_preview_labeled(template_name: str):
    """Return labeled page images as JSON array of base64 PNGs, with annotations drawn on each field."""
    import fitz
    import base64

    template_path = TEMPLATE_DIR / template_name
    if not template_path.exists() or template_path.suffix.lower() != ".pdf":
        raise HTTPException(404, "PDF template not found")

    descriptions = _load_pdf_descriptions(template_name)
    doc = fitz.open(str(template_path))

    pages = []
    for page in doc:
        # Draw red rectangles and labels on text fields
        for widget in page.widgets():
            if widget.field_type != fitz.PDF_WIDGET_TYPE_TEXT:
                continue
            name = widget.field_name
            short = name.split(".")[-1].replace("[0]", "")
            desc = descriptions.get(name, "")
            label_text = short + (f" - {desc}" if desc else "")

            rect = widget.rect
            # Red border around field
            page.draw_rect(rect, color=(1, 0, 0), width=1.2)
            # Red label above the field
            fontsize = 6
            label_rect = fitz.Rect(rect.x0, rect.y0 - 9, rect.x0 + len(label_text) * 3.5 + 6, rect.y0 - 0.5)
            page.draw_rect(label_rect, color=(1, 0, 0), fill=(1, 0, 0))
            page.insert_text(
                fitz.Point(rect.x0 + 2, rect.y0 - 2),
                label_text,
                fontsize=fontsize,
                color=(1, 1, 1),
                fontname="helv",
            )

        pix = page.get_pixmap(dpi=150)
        img_bytes = pix.tobytes("png")
        pages.append(base64.b64encode(img_bytes).decode())

    doc.close()
    return {"pages": pages, "count": len(pages)}


class PdfFieldMapRequest(BaseModel):
    template_name: str
    field_map: dict[str, str]  # {placeholder_key: pdf_field_name}


@app.post("/api/pdf-field-map")
async def save_pdf_field_map(req: PdfFieldMapRequest):
    """Save/update the field mapping for a PDF template."""
    template_path = TEMPLATE_DIR / req.template_name
    if not template_path.exists():
        raise HTTPException(404, "Template not found")

    map_path = template_path.with_suffix(".json")
    with open(map_path, "w", encoding="utf-8") as f:
        json.dump(req.field_map, f, indent=2, ensure_ascii=False)

    return {"status": "ok", "fields_mapped": len(req.field_map)}


@app.get("/api/pdf-field-map/{template_name}")
async def get_pdf_field_map(template_name: str):
    """Get the current field mapping for a PDF template."""
    map_path = (TEMPLATE_DIR / template_name).with_suffix(".json")
    if not map_path.exists():
        return {}
    with open(map_path, "r", encoding="utf-8") as f:
        return json.load(f)
