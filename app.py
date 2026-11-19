import io
import os
import re
import zipfile
from datetime import datetime
from collections import Counter

import streamlit as st
from pypdf import PdfReader
from docx import Document
from docx.shared import Pt, RGBColor
from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Pt as PPTPt
from pptx.dml.color import RGBColor as PPTRGBColor
from pptx.dml.color import MSO_COLOR_TYPE
import fitz


# ================= PAGE CONFIG =================

st.set_page_config(
    page_title="DataFix - File Repair Center",
    page_icon="🛠️",
    layout="wide"
)

st.title("🛠️ DataFix")
st.subheader("Universal File Identification, Repair & Cleaning Center")

st.write(
    "Identify errors, detect duplicate data, correct spacing, "
    "standardize fonts and colours, repair files and compare results."
)

st.info(
    "Supported formats: PDF, Word DOCX, Excel XLSX and PowerPoint PPTX."
)


# ================= GENERAL HELPERS =================

def normalize_text(value, fix_spacing=True):
    if value is None:
        return value

    text = str(value).replace("\u00a0", " ")

    if fix_spacing:
        # Join obvious letter-spaced words.
        text = re.sub(
            r"(?<![\w])(?:[A-Za-z]\s+){2,}[A-Za-z](?![\w])",
            lambda m: re.sub(r"\s+", "", m.group(0)),
            text
        )

        text = re.sub(r"[ \t]{2,}", " ", text)
        text = re.sub(r"[ \t]+([,.;:!?])", r"\1", text)
        text = re.sub(r"([,.;:!?])(?=\S)", r"\1 ", text)
        text = text.strip()

    return text


def rgb_from_hex(color):
    value = color.lstrip("#")

    if len(value) != 6:
        raise ValueError("Enter a valid six-digit colour.")

    return tuple(
        int(value[i:i + 2], 16)
        for i in (0, 2, 4)
    )


def add_finding(findings, category, location, detail):
    findings.append({
        "Category": category,
        "Location": location,
        "Finding": detail
    })


def identify_file(name, data):

    ext = os.path.splitext(name)[1].lower()

    if ext == ".pdf":
        if b"%PDF" in data[:1024]:
            return "PDF"
        return "PDF (signature warning)"

    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:

            if z.testzip():
                return "Damaged Office file"

            names = set(z.namelist())

            if ext == ".docx" and "word/document.xml" in names:
                return "Word DOCX"

            if ext == ".xlsx" and "xl/workbook.xml" in names:
                return "Excel XLSX"

            if ext == ".pptx" and "ppt/presentation.xml" in names:
                return "PowerPoint PPTX"

    except Exception:
        if ext in [".docx", ".xlsx", ".pptx"]:
            return "Damaged Office file"

    return "Unsupported file format"


# ================= WORD FUNCTIONS =================

def get_word_paragraphs(doc):

    for p in doc.paragraphs:
        yield p

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def set_word_text(paragraph, text):

    if paragraph.runs:
        paragraph.runs[0].text = text

        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text)


def inspect_word(data):

    doc = Document(io.BytesIO(data))
    paragraphs = list(get_word_paragraphs(doc))

    texts = [
        p.text.strip()
        for p in paragraphs
        if p.text.strip()
    ]

    counts = Counter(texts)

    duplicates = sum(
        count - 1
        for count in counts.values()
        if count > 1
    )

    findings = []

    for text, count in counts.items():
        if count > 1:
            add_finding(
                findings,
                "Duplicate paragraph",
                "Word document",
                f"Repeated {count} times: {text[:120]}"
            )

    for index, p in enumerate(paragraphs, 1):

        if re.search(r"[ \t]{2,}|[ \t]+[,.!?;:]", p.text):
            add_finding(
                findings,
                "Spacing",
                f"Paragraph {index}",
                "Extra spaces detected."
            )

    sizes = Counter(
        round(r.font.size.pt, 1)
        for p in paragraphs
        for r in p.runs
        if r.font.size
    )

    colours = Counter(
        str(r.font.color.rgb)
        for p in paragraphs
        for r in p.runs
        if r.font.color and r.font.color.rgb
    )

    if len(sizes) > 1:
        add_finding(
            findings,
            "Font size inconsistency",
            "Word document",
            str(dict(sizes))
        )

    if len(colours) > 1:
        add_finding(
            findings,
            "Font colour inconsistency",
            "Word document",
            str(dict(colours))
        )

    return {
        "summary": {
            "Paragraphs": len(paragraphs),
            "Duplicate paragraphs": duplicates,
            "Tables": len(doc.tables)
        },
        "findings": findings
    }


def clean_word(data, options):

    doc = Document(io.BytesIO(data))
    changes = Counter()
    seen = set()

    paragraphs = list(get_word_paragraphs(doc))

    if options["remove_duplicates"]:

        for p in paragraphs:

            key = normalize_text(p.text).casefold()

            if key and key in seen:

                try:
                    p._element.getparent().remove(p._element)
                    changes["Duplicate paragraphs removed"] += 1
                except Exception:
                    changes["Duplicate paragraphs not removed"] += 1

            elif key:
                seen.add(key)

    for p in get_word_paragraphs(doc):

        if options["fix_spacing"]:

            old = p.text
            new = normalize_text(old)

            if new != old:
                set_word_text(p, new)
                changes["Paragraphs corrected"] += 1

        for run in p.runs:

            if options["font_size"]:
                run.font.size = Pt(options["font_size_value"])
                changes["Font sizes corrected"] += 1

            if options["font_color"]:
                run.font.color.rgb = RGBColor(
                    *rgb_from_hex(options["font_color_value"])
                )
                changes["Font colours corrected"] += 1

    output = io.BytesIO()
    doc.save(output)

    return output.getvalue(), dict(changes)


# ================= EXCEL FUNCTIONS =================

def inspect_excel(data):

    wb = load_workbook(
        io.BytesIO(data),
        data_only=False
    )

    findings = []

    summary = {
        "Worksheets": len(wb.sheetnames),
        "Duplicate rows": 0,
        "Cells with extra spaces": 0
    }

    for ws in wb.worksheets:

        seen = {}
        sizes = Counter()
        colours = Counter()

        for row in range(1, ws.max_row + 1):

            values = tuple(
                ws.cell(row, col).value
                for col in range(1, ws.max_column + 1)
            )

            if not any(v is not None and v != "" for v in values):
                continue

            key = tuple(
                (type(v).__name__, str(v))
                for v in values
            )

            if key in seen:

                summary["Duplicate rows"] += 1

                add_finding(
                    findings,
                    "Duplicate row",
                    f"{ws.title}, row {row}",
                    f"Matches row {seen[key]}"
                )

            else:
                seen[key] = row

            for col in range(1, ws.max_column + 1):

                cell = ws.cell(row, col)

                if isinstance(cell.value, str):

                    if normalize_text(cell.value) != cell.value:

                        summary["Cells with extra spaces"] += 1

                        add_finding(
                            findings,
                            "Spacing",
                            f"{ws.title}!{cell.coordinate}",
                            "Extra spaces detected."
                        )

                if cell.value is not None:

                    if cell.font.sz:
                        sizes[round(cell.font.sz, 1)] += 1

                    if (
                        cell.font.color
                        and cell.font.color.type == "rgb"
                        and cell.font.color.rgb
                    ):
                        colours[str(cell.font.color.rgb)] += 1

        if len(sizes) > 1:
            add_finding(
                findings,
                "Font size inconsistency",
                ws.title,
                str(dict(sizes))
            )

        if len(colours) > 1:
            add_finding(
                findings,
                "Font colour inconsistency",
                ws.title,
                str(dict(colours))
            )

    wb.close()

    return {
        "summary": summary,
        "findings": findings
    }


def clean_excel(data, options):

    wb = load_workbook(
        io.BytesIO(data),
        data_only=False,
        keep_links=True
    )

    changes = Counter()

    for ws in wb.worksheets:

        if options["remove_duplicates"]:

            seen = set()
            duplicates = []

            for row in range(1, ws.max_row + 1):

                values = tuple(
                    ws.cell(row, col).value
                    for col in range(1, ws.max_column + 1)
                )

                if not any(v is not None and v != "" for v in values):
                    continue

                key = tuple(
                    (type(v).__name__, str(v))
                    for v in values
                )

                if key in seen:
                    duplicates.append(row)
                else:
                    seen.add(key)

            for row in reversed(duplicates):
                ws.delete_rows(row, 1)
                changes["Duplicate rows removed"] += 1

        for row in ws.iter_rows():

            for cell in row:

                if options["fix_spacing"] and isinstance(cell.value, str):

                    new = normalize_text(cell.value)

                    if new != cell.value:
                        cell.value = new
                        changes["Cells corrected"] += 1

                if options["font_size"] or options["font_color"]:

                    font = cell.font.copy()

                    if options["font_size"]:
                        font.sz = options["font_size_value"]

                    if options["font_color"]:
                        font.color = options["font_color_value"].replace(
                            "#", ""
                        ).upper()

                    cell.font = font

    output = io.BytesIO()
    wb.save(output)
    wb.close()

    return output.getvalue(), dict(changes)


# ================= POWERPOINT FUNCTIONS =================

def get_ppt_paragraphs(prs):

    for slide_no, slide in enumerate(prs.slides, 1):

        for shape_no, shape in enumerate(slide.shapes, 1):

            if getattr(shape, "has_text_frame", False):

                for p in shape.text_frame.paragraphs:
                    yield slide_no, shape_no, p

            if getattr(shape, "has_table", False):

                for row in shape.table.rows:

                    for cell in row.cells:

                        for p in cell.text_frame.paragraphs:
                            yield slide_no, shape_no, p


def set_ppt_text(paragraph, text):

    if paragraph.runs:

        paragraph.runs[0].text = text

        for run in paragraph.runs[1:]:
            run.text = ""

    else:
        paragraph.add_run().text = text


def inspect_powerpoint(data):

    prs = Presentation(io.BytesIO(data))
    paragraphs = list(get_ppt_paragraphs(prs))

    texts = [
        p.text.strip()
        for _, _, p in paragraphs
        if p.text.strip()
    ]

    counts = Counter(texts)

    duplicates = sum(
        n - 1
        for n in counts.values()
        if n > 1
    )

    findings = []

    for text, count in counts.items():

        if count > 1:
            add_finding(
                findings,
                "Duplicate text",
                "Presentation",
                f"Repeated {count} times: {text[:120]}"
            )

    sizes = Counter(
        round(run.font.size.pt, 1)
        for _, _, p in paragraphs
        for run in p.runs
        if run.font.size
    )

    colours = Counter(
        str(run.font.color.rgb)
        for _, _, p in paragraphs
        for run in p.runs
        if run.font.color.type == MSO_COLOR_TYPE.RGB
        and run.font.color.rgb is not None
    )

    for slide_no, shape_no, p in paragraphs:

        if re.search(r"[ \t]{2,}|[ \t]+[,.!?;:]", p.text):

            add_finding(
                findings,
                "Spacing",
                f"Slide {slide_no}, shape {shape_no}",
                "Extra spaces detected."
            )

    if len(sizes) > 1:
        add_finding(
            findings,
            "Font size inconsistency",
            "Presentation",
            str(dict(sizes))
        )

    if len(colours) > 1:
        add_finding(
            findings,
            "Font colour inconsistency",
            "Presentation",
            str(dict(colours))
        )

    return {
        "summary": {
            "Slides": len(prs.slides),
            "Duplicate paragraphs": duplicates
        },
        "findings": findings
    }


def clean_powerpoint(data, options):

    prs = Presentation(io.BytesIO(data))
    changes = Counter()
    seen = set()

    for slide_no, shape_no, p in list(get_ppt_paragraphs(prs)):

        key = normalize_text(p.text).casefold()

        if options["remove_duplicates"] and key:

            if key in seen:

                try:
                    p._p.getparent().remove(p._p)
                    changes["Duplicate paragraphs removed"] += 1
                    continue
                except Exception:
                    changes["Duplicate paragraphs not removed"] += 1

            else:
                seen.add(key)

        if options["fix_spacing"]:

            old = p.text
            new = normalize_text(old)

            if new != old:
                set_ppt_text(p, new)
                changes["Spacing corrected"] += 1

        for run in p.runs:

            if options["font_size"]:

                run.font.size = PPTPt(
                    options["font_size_value"]
                )

                changes["Font sizes corrected"] += 1

            if options["font_color"]:

                run.font.color.rgb = PPTRGBColor(
                    *rgb_from_hex(options["font_color_value"])
                )

                changes["Font colours corrected"] += 1

    output = io.BytesIO()
    prs.save(output)

    return output.getvalue(), dict(changes)


# ================= PDF FUNCTIONS =================

def get_pdf_lines(data):

    doc = fitz.open(
        stream=data,
        filetype="pdf"
    )

    lines = []

    for page_no, page in enumerate(doc, 1):

        page_dict = page.get_text("dict")

        for block in page_dict.get("blocks", []):

            for line in block.get("lines", []):

                spans = line.get("spans", [])

                text = "".join(
                    span.get("text", "")
                    for span in spans
                )

                rects = [
                    fitz.Rect(span["bbox"])
                    for span in spans
                    if span.get("text", "").strip()
                ]

                if text.strip() and rects:

                    bbox = rects[0]

                    for rect in rects[1:]:
                        bbox |= rect

                    lines.append({
                        "page": page_no - 1,
                        "bbox": bbox,
                        "text": text,
                        "size": spans[0].get("size", 10),
                        "color": spans[0].get("color", 0)
                    })

    doc.close()

    return lines


def inspect_pdf(data):

    reader = PdfReader(io.BytesIO(data), strict=False)
    lines = get_pdf_lines(data)

    findings = []

    seen = set()

    for line in lines:

        text = normalize_text(line["text"])
        key = text.casefold()

        if key and key in seen:

            add_finding(
                findings,
                "Duplicate text line",
                f"Page {line['page'] + 1}",
                text[:120]
            )

        elif key:
            seen.add(key)

        if text != line["text"]:

            add_finding(
                findings,
                "Spacing",
                f"Page {line['page'] + 1}",
                "Extra spaces or letter spacing detected."
            )

    sizes = Counter(
        round(float(line["size"]), 1)
        for line in lines
    )

    colours = Counter(
        str(line["color"])
        for line in lines
    )

    if len(sizes) > 1:
        add_finding(
            findings,
            "Font size inconsistency",
            "PDF",
            str(dict(sizes))
        )

    if len(colours) > 1:
        add_finding(
            findings,
            "Font colour inconsistency",
            "PDF",
            str(dict(colours))
        )

    return {
        "summary": {
            "Pages": len(reader.pages),
            "Extracted text lines": len(lines),
            "Metadata fields": len(reader.metadata or {})
        },
        "findings": findings
    }


def clean_pdf(data, options):

    doc = fitz.open(
        stream=data,
        filetype="pdf"
    )

    lines = get_pdf_lines(data)
    changes = Counter()
    seen = set()
    replacements = []

    for item in lines:

        original = item["text"]

        cleaned = normalize_text(
            original,
            options["fix_spacing"]
        )

        key = cleaned.casefold()

        duplicate = (
            options["remove_duplicates"]
            and key
            and key in seen
        )

        if key and not duplicate:
            seen.add(key)

        if duplicate:

            changes["Duplicate text lines removed"] += 1
            replacements.append((item, None))

        elif (
            cleaned != original
            or options["font_size"]
            or options["font_color"]
        ):

            if cleaned != original:
                changes["Spacing corrected"] += 1

            if options["font_size"]:
                changes["Font size applied"] += 1

            if options["font_color"]:
                changes["Font colour applied"] += 1

            replacements.append((item, cleaned))

    per_page = {}

    for item, replacement in replacements:

        page = doc[item["page"]]

        page.add_redact_annot(
            item["bbox"],
            fill=(1, 1, 1)
        )

        per_page.setdefault(
            item["page"], []
        ).append((item, replacement))

    for page_no in per_page:
        doc[page_no].apply_redactions(
            images=0,
            graphics=0
        )

    target_color = (
        rgb_from_hex(options["font_color_value"])
        if options["font_color"]
        else None
    )

    for page_no, entries in per_page.items():

        page = doc[page_no]

        for item, replacement in entries:

            if replacement is None:
                continue

            size = (
                options["font_size_value"]
                if options["font_size"]
                else max(5, float(item["size"]))
            )

            color = target_color if target_color else (
                (
                    (item["color"] >> 16 & 255) / 255,
                    (item["color"] >> 8 & 255) / 255,
                    (item["color"] & 255) / 255
                )
            )

            rect = fitz.Rect(item["bbox"])

            rect.x0 += 0.3
            rect.y0 += 0.2
            rect.x1 -= 0.3
            rect.y1 -= 0.2

            inserted = False

            for attempt in range(4):

                result = page.insert_textbox(
                    rect,
                    replacement,
                    fontname="helv",
                    fontsize=max(4, size - attempt * 0.5),
                    color=color,
                    overlay=True
                )

                if result >= 0:
                    inserted = True
                    break

            if not inserted:
                changes["Lines needing manual review"] += 1

    output = doc.tobytes(
        garbage=4,
        deflate=True
    )

    doc.close()

    return output, dict(changes)


# ================= PROCESS ROUTER =================

def inspect_file(file_type, data):

    if file_type.startswith("PDF"):
        return inspect_pdf(data)

    if file_type == "Word DOCX":
        return inspect_word(data)

    if file_type == "Excel XLSX":
        return inspect_excel(data)

    if file_type == "PowerPoint PPTX":
        return inspect_powerpoint(data)

    raise ValueError("Unsupported file format.")


def clean_file(file_type, data, options):

    if file_type.startswith("PDF"):
        return clean_pdf(data, options)

    if file_type == "Word DOCX":
        return clean_word(data, options)

    if file_type == "Excel XLSX":
        return clean_excel(data, options)

    if file_type == "PowerPoint PPTX":
        return clean_powerpoint(data, options)

    raise ValueError("Unsupported file format.")


def verify_file(file_type, data):

    if file_type.startswith("PDF"):

        reader = PdfReader(io.BytesIO(data), strict=False)

        return {
            "Pages": len(reader.pages)
        }

    if file_type == "Word DOCX":

        doc = Document(io.BytesIO(data))

        return {
            "Paragraphs": len(list(get_word_paragraphs(doc)))
        }

    if file_type == "Excel XLSX":

        wb = load_workbook(
            io.BytesIO(data),
            read_only=True
        )

        result = {
            "Worksheets": len(wb.sheetnames)
        }

        wb.close()

        return result

    if file_type == "PowerPoint PPTX":

        prs = Presentation(io.BytesIO(data))

        return {
            "Slides": len(prs.slides)
        }

    raise ValueError("Verification failed.")


# ================= MAIN WEBSITE =================

st.header("📁 Upload Your File")

uploaded = st.file_uploader(
    "Choose a PDF, Word, Excel or PowerPoint file",
    type=["pdf", "docx", "xlsx", "pptx"]
)

if uploaded:

    original = uploaded.getvalue()
    filename = uploaded.name

    file_type = identify_file(
        filename,
        original
    )

    st.divider()

    st.header("📋 File Information")

    col1, col2, col3 = st.columns(3)

    col1.metric("File Name", filename)

    col2.metric("File Format", file_type)

    col3.metric(
        "File Size",
        f"{len(original) / 1024:.2f} KB"
    )

    if file_type in [
        "Unsupported file format",
        "Damaged Office file"
    ]:

        st.error("This file is unsupported or damaged.")

    else:

        try:

            st.header("🔍 1. Pre-cleaning Inspection")

            inspection = inspect_file(
                file_type,
                original
            )

            st.caption(
                "Inspection date: "
                + datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            )

            st.json(inspection["summary"])

            findings = inspection["findings"]

            if findings:

                st.warning(
                    f"{len(findings)} possible issue(s) detected."
                )

                st.dataframe(
                    findings,
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.success(
                    "No duplicate text or common spacing issues "
                    "were detected by the available checks."
                )

            st.divider()

            st.header("🧹 2. Choose Cleaning Actions")

            st.write(
                "Review the findings and select the corrections "
                "you want DataFix to perform."
            )

            remove_duplicates = st.checkbox(
                "Detect and remove duplicate data",
                value=True
            )

            fix_spacing = st.checkbox(
                "Correct extra spaces and letter spacing",
                value=True
            )

            font_size = st.checkbox(
                "Standardize font size",
                value=False
            )

            font_size_value = st.number_input(
                "Target font size (points)",
                min_value=5.0,
                max_value=48.0,
                value=11.0,
                step=1.0,
                disabled=not font_size
            )

            font_color = st.checkbox(
                "Standardize font colour",
                value=False
            )

            font_color_value = st.color_picker(
                "Choose font colour",
                value="#000000",
                disabled=not font_color
            )

            st.warning(
                "PDF text correction can change page layout and uses "
                "a standard font when reinserting text. Scanned PDFs "
                "require OCR. Complex formatting, graphics, embedded "
                "fonts and annotations may not be preserved. "
                "Always inspect the downloaded copy."
            )

            options = {
                "remove_duplicates": remove_duplicates,
                "fix_spacing": fix_spacing,
                "font_size": font_size,
                "font_size_value": float(font_size_value),
                "font_color": font_color,
                "font_color_value": font_color_value
            }

            st.divider()

            if st.button(
                "🛠️ Apply Selected Cleaning",
                type="primary",
                use_container_width=True
            ):

                if not any([
                    remove_duplicates,
                    fix_spacing,
                    font_size,
                    font_color
                ]):

                    st.error(
                        "Select at least one cleaning action."
                    )

                else:

                    with st.spinner(
                        "Cleaning and verifying your file..."
                    ):

                        cleaned, changes = clean_file(
                            file_type,
                            original,
                            options
                        )

                        verification = verify_file(
                            file_type,
                            cleaned
                        )

                        after = inspect_file(
                            file_type,
                            cleaned
                        )

                    ext = os.path.splitext(filename)[1].lower()

                    cleaned_name = (
                        os.path.splitext(filename)[0]
                        + "_datafix_cleaned"
                        + ext
                    )

                    mime_types = {
                        ".pdf": "application/pdf",
                        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    }

                    st.success(
                        "Cleaning completed and output passed readability verification."
                    )

                    st.header("📊 3. Cleaning Report")

                    a, b, c = st.columns(3)

                    a.metric(
                        "Original Size",
                        f"{len(original) / 1024:.2f} KB"
                    )

                    b.metric(
                        "Cleaned Size",
                        f"{len(cleaned) / 1024:.2f} KB"
                    )

                    c.metric(
                        "Changes Recorded",
                        sum(changes.values())
                    )

                    st.write("**Original file:**", filename)

                    st.write("**Cleaned file:**", cleaned_name)

                    st.write("**Format preserved:**", file_type)

                    st.write("**Output verification:** Passed")

                    st.subheader("Changes performed")

                    if changes:
                        st.json(changes)
                    else:
                        st.info(
                            "No selected corrections were required "
                            "or applied."
                        )

                    st.subheader("Output verification")

                    st.json(verification)

                    st.divider()

                    st.header("🔄 4. Before and After Comparison")

                    left, right = st.columns(2)

                    with left:

                        st.subheader("Original File")

                        st.json(inspection["summary"])

                    with right:

                        st.subheader("Cleaned File")

                        st.json(after["summary"])

                    st.subheader("Remaining findings")

                    if after["findings"]:

                        st.dataframe(
                            after["findings"],
                            use_container_width=True,
                            hide_index=True
                        )

                    else:

                        st.success(
                            "No findings from the supported checks remain."
                        )

                    st.divider()

                    st.header("⬇️ 5. Download Cleaned File")

                    st.download_button(
                        label="Download Cleaned File",
                        data=cleaned,
                        file_name=cleaned_name,
                        mime=mime_types.get(
                            ext,
                            "application/octet-stream"
                        ),
                        use_container_width=True
                    )

                    st.caption(
                        "The original uploaded file has not been changed. "
                        "Keep it as a backup and compare the cleaned copy "
                        "before relying on it."
                    )

        except Exception as e:

            st.error(
                f"Processing failed: {str(e)}"
            )

            st.warning(
                "The original file remains unchanged. "
                "The file may contain unsupported or severely "
                "damaged components."
            )


# ================= FOOTER =================

st.divider()

st.caption(
    "🛠️ DataFix | Universal File Repair & Cleaning Center"
)

st.caption(
    "PDF • Word • Excel • PowerPoint | "
    "Original files are preserved."
)
