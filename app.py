import streamlit as st
import io, os, zipfile, difflib
import xml.etree.ElementTree as ET
from datetime import datetime
from pypdf import PdfReader, PdfWriter
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation

st.set_page_config(page_title="DataFix - File Repair Center", page_icon="🛠️", layout="wide")

# ---------- File identification and inspection ----------

def identify_file(name, data):
    ext = os.path.splitext(name)[1].lower()
    if ext == ".pdf":
        return "PDF" if b"%PDF" in data[:1024] else "PDF (signature warning)"
    required = {
        ".docx": ("Word DOCX", "word/document.xml"),
        ".xlsx": ("Excel XLSX", "xl/workbook.xml"),
        ".pptx": ("PowerPoint PPTX", "ppt/presentation.xml"),
    }
    if ext in required:
        label, part = required[ext]
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                return label if part in z.namelist() else f"{label} (required component missing)"
        except Exception:
            return "Damaged Office file"
    return "Unsupported file format"


def inspect_zip(data, main_part):
    findings, errors = [], []
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            names = z.namelist()
            bad = z.testzip()
            if bad:
                errors.append(f"ZIP integrity check failed: {bad}")
            else:
                findings.append("ZIP container integrity check passed.")
            if main_part not in names:
                errors.append(f"Required component is missing: {main_part}")
            xml_count = 0
            for name in names:
                if name.lower().endswith(".xml"):
                    xml_count += 1
                    try:
                        ET.fromstring(z.read(name))
                    except Exception as exc:
                        errors.append(f"Invalid XML component {name}: {exc}")
            if not any("Invalid XML" in e for e in errors):
                findings.append(f"All {xml_count} XML component(s) parsed successfully.")
    except Exception as exc:
        errors.append(f"Could not read Office ZIP structure: {exc}")
    return findings, errors


# ---------- Extract readable content for before/after comparison ----------

def extract_pdf(data):
    reader = PdfReader(io.BytesIO(data), strict=False)
    if reader.is_encrypted:
        raise ValueError("Password-protected PDFs are not supported.")
    pages, lines, warnings = [], [], []
    for i, page in enumerate(reader.pages, 1):
        try:
            text = page.extract_text() or ""
        except Exception as exc:
            text = ""
            warnings.append(f"Page {i}: text extraction failed: {exc}")
        pages.append(text)
        lines.append(f"PAGE {i}\n{text}")
    metadata = {str(k): str(v) for k, v in (reader.metadata or {}).items() if v is not None}
    return {
        "lines": lines, "metadata": metadata, "warnings": warnings,
        "summary": {
            "Pages": len(reader.pages),
            "Metadata fields": len(metadata),
            "Pages with extractable text": sum(bool(x.strip()) for x in pages),
        },
    }


def extract_docx(data):
    doc = Document(io.BytesIO(data))
    lines = [f"Paragraph {i}: {p.text}" for i, p in enumerate(doc.paragraphs, 1)]
    for ti, table in enumerate(doc.tables, 1):
        lines.append(f"TABLE {ti}")
        for ri, row in enumerate(table.rows, 1):
            lines.append(f"Row {ri}: " + " | ".join(c.text for c in row.cells))
    return {"lines": lines, "summary": {
        "Paragraphs": len(doc.paragraphs), "Tables": len(doc.tables),
        "Sections": len(doc.sections),
        "Non-empty paragraphs": sum(bool(p.text.strip()) for p in doc.paragraphs),
    }, "warnings": []}


def extract_xlsx(data):
    wb = load_workbook(io.BytesIO(data), data_only=False, keep_links=True)
    lines, sheets, cells, formulas = [], {}, 0, 0
    for ws in wb.worksheets:
        sheets[ws.title] = {"Rows": ws.max_row, "Columns": ws.max_column}
        lines.append(f"WORKSHEET: {ws.title}")
        for row in ws.iter_rows():
            values = []
            for cell in row:
                if cell.value is not None:
                    cells += 1
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formulas += 1
                    values.append(f"{cell.coordinate}={cell.value}")
            if values:
                lines.append(" | ".join(values))
    summary = {"Worksheets": len(wb.sheetnames), "Sheet details": sheets,
               "Non-empty cells": cells, "Formula cells": formulas}
    wb.close()
    return {"lines": lines, "summary": summary, "warnings": []}


def extract_pptx(data):
    prs = Presentation(io.BytesIO(data))
    lines, text_shapes = [], 0
    for si, slide in enumerate(prs.slides, 1):
        lines.append(f"SLIDE {si}")
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False) and shape.text.strip():
                text_shapes += 1
                lines.append(shape.text)
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    lines.append(" | ".join(cell.text for cell in row.cells))
    return {"lines": lines, "summary": {
        "Total slides": len(prs.slides), "Text-containing shapes": text_shapes,
    }, "warnings": []}


def extract_content(kind, data):
    if kind.startswith("PDF"):
        return extract_pdf(data)
    if kind == "Word DOCX":
        return extract_docx(data)
    if kind == "Excel XLSX":
        return extract_xlsx(data)
    if kind == "PowerPoint PPTX":
        return extract_pptx(data)
    raise ValueError("Unsupported or damaged file format.")


def inspect_file(kind, data):
    findings, errors, warnings = [], [], []
    if kind.startswith("PDF"):
        if b"%PDF" not in data[:1024]:
            warnings.append("PDF signature was not found near the start of the file.")
    elif kind == "Word DOCX":
        f, e = inspect_zip(data, "word/document.xml")
        findings += f
        errors += e
    elif kind == "Excel XLSX":
        f, e = inspect_zip(data, "xl/workbook.xml")
        findings += f
        errors += e
    elif kind == "PowerPoint PPTX":
        f, e = inspect_zip(data, "ppt/presentation.xml")
        findings += f
        errors += e
    else:
        errors.append("Unsupported or unrecognized file format.")

    try:
        details = extract_content(kind, data)
        findings.append("The document parser could read the file content.")
        findings.append(f"Content summary: {details['summary']}")
        warnings += details.get("warnings", [])
        summary = details["summary"]
        if kind.startswith("PDF") and summary["Pages"] and not summary["Pages with extractable text"]:
            warnings.append("No extractable text found. This may be a scanned/image-only PDF.")
        if kind == "Word DOCX" and not summary["Non-empty paragraphs"] and not summary["Tables"]:
            warnings.append("No paragraph text or tables were found.")
        if kind == "Excel XLSX" and not summary["Non-empty cells"]:
            warnings.append("No non-empty cells were found.")
        if kind == "PowerPoint PPTX" and not summary["Total slides"]:
            warnings.append("The presentation contains no slides.")
    except Exception as exc:
        errors.append(f"Content parsing failed: {exc}")

    return {"findings": findings, "errors": errors, "warnings": warnings}


# ---------- Re-save supported files ----------

def repair_file(kind, data):
    if kind.startswith("PDF"):
        reader = PdfReader(io.BytesIO(data), strict=False)
        if reader.is_encrypted:
            raise ValueError("Password-protected PDFs are not supported.")
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        if reader.metadata:
            metadata = {str(k): str(v) for k, v in reader.metadata.items() if v is not None}
            if metadata:
                writer.add_metadata(metadata)
        out = io.BytesIO()
        writer.write(out)
        return out.getvalue(), "application/pdf", ".pdf"

    if kind == "Word DOCX":
        doc = Document(io.BytesIO(data))
        out = io.BytesIO()
        doc.save(out)
        return out.getvalue(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx"

    if kind == "Excel XLSX":
        wb = load_workbook(io.BytesIO(data), data_only=False, keep_links=True)
        out = io.BytesIO()
        wb.save(out)
        wb.close()
        return out.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"

    if kind == "PowerPoint PPTX":
        prs = Presentation(io.BytesIO(data))
        out = io.BytesIO()
        prs.save(out)
        return out.getvalue(), "application/vnd.openxmlformats-officedocument.presentationml.presentation", ".pptx"

    raise ValueError("Unsupported file format.")


def compare_files(kind, original, repaired):
    before = extract_content(kind, original)
    after = extract_content(kind, repaired)
    diff = list(difflib.unified_diff(
        before["lines"], after["lines"],
        fromfile="Original", tofile="Repaired", lineterm=""
    ))
    added = sum(1 for x in diff if x.startswith("+") and not x.startswith("+++"))
    removed = sum(1 for x in diff if x.startswith("-") and not x.startswith("---"))
    return before, after, diff, added, removed


# ---------- Website UI ----------

st.title("🛠️ DataFix")
st.subheader("Universal File Identification & Repair Center")
st.write(
    "Inspect files before repair, identify readable structural problems, "
    "compare original and repaired content, and download the repaired file."
)
st.divider()
st.info("Supported formats: PDF, Word DOCX, Excel XLSX and PowerPoint PPTX.")

st.header("📁 Upload Your File")
uploaded = st.file_uploader(
    "Choose a PDF, Word, Excel or PowerPoint file",
    type=["pdf", "docx", "xlsx", "pptx"],
    help="Upload one file at a time."
)

if uploaded:
    original = uploaded.getvalue()
    filename = uploaded.name
    kind = identify_file(filename, original)

    st.divider()
    st.header("📋 File Information")
    c1, c2, c3 = st.columns(3)
    c1.metric("File name", filename)
    c2.metric("File format", kind)
    c3.metric("File size", f"{len(original) / 1024:.2f} KB")

    st.divider()
    st.header("🔍 1. Pre-repair Error Inspection")
    st.write("Inspection date:", datetime.now().strftime("%d-%m-%Y %H:%M:%S"))
    inspection = inspect_file(kind, original)

    if inspection["errors"]:
        st.error("Potential structural or reading problems were detected.")
    else:
        st.success("The file passed the available structure and readability checks.")

    if inspection["findings"]:
        st.markdown("**Inspection findings**")
        for item in inspection["findings"]:
            st.write("•", item)
    if inspection["warnings"]:
        st.markdown("**Warnings**")
        for item in inspection["warnings"]:
            st.warning(item)
    if inspection["errors"]:
        st.markdown("**Errors**")
        for item in inspection["errors"]:
            st.error(item)

    st.caption(
        "These checks detect issues visible to the supported parsers. "
        "They cannot identify every possible content, formatting, or application-specific defect."
    )

    st.divider()
    st.header("🛠️ 2. Repair File")
    if st.button("🛠️ Repair File", type="primary", use_container_width=True):
        try:
            with st.spinner("Rebuilding and verifying the document..."):
                repaired, mime, ext = repair_file(kind, original)
                repaired_name = os.path.splitext(filename)[0] + "_repaired" + ext

                # Re-open the generated file to verify that it remains readable.
                output_check = inspect_file(kind, repaired)
                if output_check["errors"]:
                    raise ValueError("Output verification failed: " + "; ".join(output_check["errors"]))

                before, after, diff, added, removed = compare_files(kind, original, repaired)
                st.session_state["datafix_result"] = {
                    "filename": filename, "repaired_name": repaired_name,
                    "repaired": repaired, "mime": mime, "before": before,
                    "after": after, "diff": diff, "added": added,
                    "removed": removed, "output_check": output_check,
                }
        except Exception as exc:
            st.session_state.pop("datafix_result", None)
            st.error(f"Repair was unsuccessful: {exc}")
            st.warning("The original upload has not been changed.")

    result = st.session_state.get("datafix_result")
    if result and result["filename"] == filename:
        st.success("Repaired output generated and passed the available readability checks.")

        st.header("📊 3. Repair Report")
        a, b, c = st.columns(3)
        a.metric("Original size", f"{len(original) / 1024:.2f} KB")
        b.metric("Repaired size", f"{len(result['repaired']) / 1024:.2f} KB")
        c.metric("Detected content differences", result["added"] + result["removed"])
        st.write("**Original file:**", filename)
        st.write("**Repaired file:**", result["repaired_name"])
        st.write("**Original format preserved:**", kind)
        st.write("**Output verification:** Passed")

        st.divider()
        st.header("🔄 4. Original vs Repaired Comparison")
        left, right = st.columns(2)
        with left:
            st.subheader("Original file")
            st.json(result["before"]["summary"])
        with right:
            st.subheader("Repaired file")
            st.json(result["after"]["summary"])

        x, y = st.columns(2)
        x.metric("Content lines added", result["added"])
        y.metric("Content lines removed/changed", result["removed"])

        if result["diff"]:
            st.markdown("**Detected text / cell / content differences**")
            st.code("\n".join(result["diff"][:120]), language="diff")
            if len(result["diff"]) > 120:
                st.info("Only the first 120 difference lines are displayed.")
        else:
            st.success(
                "No extracted text, paragraph, slide-text, or cell-value differences were found. "
                "Internal formatting or binary changes may still exist."
            )

        if kind.startswith("PDF"):
            with st.expander("Original PDF metadata"):
                st.json(result["before"].get("metadata", {}))
            with st.expander("Repaired PDF metadata"):
                st.json(result["after"].get("metadata", {}))

        st.divider()
        st.header("⬇️ 5. Download")
        st.download_button(
            "⬇️ Download Repaired File",
            data=result["repaired"],
            file_name=result["repaired_name"],
            mime=result["mime"],
            use_container_width=True,
        )
        st.caption(
            "DataFix re-saves files using supported Python libraries. This does not guarantee "
            "that all corruption has been recovered. Review important files and compare them "
            "with the original before relying on the repaired copy."
        )

st.divider()
st.caption("🛠️ DataFix | Universal File Identification & Repair Center")
st.caption("PDF • Word • Excel • PowerPoint | Original uploads are preserved.")
