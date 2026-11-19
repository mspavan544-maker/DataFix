import streamlit as st
import pandas as pd
import io
import os
import zipfile
from datetime import datetime

from pypdf import PdfReader, PdfWriter
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation


# ---------------- PAGE CONFIG ----------------

st.set_page_config(
    page_title="DataFix - File Repair Center",
    page_icon="🛠️",
    layout="wide"
)


# ---------------- DESIGN ----------------

st.title("🛠️ DataFix")
st.subheader("Universal File Identification & Repair Center")

st.write(
    "Identify, inspect, detect errors and repair your files. "
    "Download repaired files in their original format."
)

st.divider()

st.info(
    "Supported formats: PDF, Word DOCX, Excel XLSX and PowerPoint PPTX."
)


# ---------------- FILE IDENTIFICATION ----------------

def identify_file(file_name, file_bytes):

    extension = os.path.splitext(file_name)[1].lower()

    if extension == ".pdf":
        if b"%PDF" in file_bytes[:1024]:
            return "PDF"
        return "PDF (signature warning)"

    if extension in [".docx", ".xlsx", ".pptx"]:

        if not zipfile.is_zipfile(io.BytesIO(file_bytes)):
            return "Damaged Office file"

        try:
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
                names = z.namelist()

                if "word/document.xml" in names:
                    return "Word DOCX"

                if "xl/workbook.xml" in names:
                    return "Excel XLSX"

                if "ppt/presentation.xml" in names:
                    return "PowerPoint PPTX"

        except Exception:
            return "Damaged Office file"

    return "Unsupported file format"


# ---------------- ZIP VALIDATION ----------------

def check_zip(file_bytes):

    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            bad_file = z.testzip()

            if bad_file:
                return False, f"Damaged internal component: {bad_file}"

        return True, "Office file structure is readable."

    except Exception as e:
        return False, str(e)


# ---------------- PDF REPAIR ----------------

def repair_pdf(file_bytes):

    reader = PdfReader(
        io.BytesIO(file_bytes),
        strict=False
    )

    if reader.is_encrypted:
        raise ValueError(
            "Password-protected PDFs are not supported."
        )

    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    if reader.metadata:
        metadata = {}

        for key, value in reader.metadata.items():
            if value is not None:
                metadata[str(key)] = str(value)

        if metadata:
            writer.add_metadata(metadata)

    output = io.BytesIO()
    writer.write(output)

    return output.getvalue(), {
        "Pages": len(reader.pages),
        "Metadata fields": len(reader.metadata or {})
    }


# ---------------- WORD REPAIR ----------------

def repair_word(file_bytes):

    doc = Document(io.BytesIO(file_bytes))

    paragraph_count = len(doc.paragraphs)
    table_count = len(doc.tables)

    output = io.BytesIO()
    doc.save(output)

    return output.getvalue(), {
        "Paragraphs": paragraph_count,
        "Tables": table_count
    }


# ---------------- EXCEL REPAIR ----------------

def repair_excel(file_bytes):

    workbook = load_workbook(
        io.BytesIO(file_bytes),
        data_only=False,
        keep_links=True
    )

    sheet_details = {}

    for sheet in workbook.worksheets:

        sheet_details[sheet.title] = {
            "Rows": sheet.max_row,
            "Columns": sheet.max_column
        }

    output = io.BytesIO()

    workbook.save(output)

    return output.getvalue(), {
        "Worksheets": len(workbook.sheetnames),
        "Sheet details": sheet_details
    }


# ---------------- POWERPOINT REPAIR ----------------

def repair_powerpoint(file_bytes):

    presentation = Presentation(
        io.BytesIO(file_bytes)
    )

    slide_count = len(presentation.slides)

    output = io.BytesIO()
    presentation.save(output)

    return output.getvalue(), {
        "Total slides": slide_count
    }


# ---------------- FILE UPLOAD ----------------

st.header("📁 Upload Your File")

uploaded_file = st.file_uploader(
    "Choose a PDF, Word, Excel or PowerPoint file",
    type=["pdf", "docx", "xlsx", "pptx"],
    help="Upload one file at a time."
)


if uploaded_file is not None:

    original_bytes = uploaded_file.getvalue()
    file_name = uploaded_file.name

    file_type = identify_file(file_name, original_bytes)

    st.divider()

    st.header("📋 File Information")

    col1, col2, col3 = st.columns(3)

    col1.metric("File name", file_name)
    col2.metric("File format", file_type)
    col3.metric(
        "File size",
        f"{len(original_bytes) / 1024:.2f} KB"
    )

    st.divider()

    st.header("🔍 File Inspection Report")

    st.write("**Inspection date:**", datetime.now().strftime(
        "%d-%m-%Y %H:%M:%S"
    ))

    if file_type == "Damaged Office file":
        st.error(
            "The Office file is not a readable ZIP document. "
            "Automatic repair may not be possible."
        )

    elif file_type == "Unsupported file format":
        st.error("This file format is not supported.")

    else:

        if file_type in ["Word DOCX", "Excel XLSX", "PowerPoint PPTX"]:

            zip_ok, zip_message = check_zip(original_bytes)

            if not zip_ok:
                st.error(zip_message)

            else:
                st.success(zip_message)

        if file_type == "PDF (signature warning)":
            st.warning(
                "The PDF extension was detected, but its signature "
                "could not be verified."
            )

        st.info(
            "The repair engine will attempt to read and rebuild "
            "the supported document structure."
        )

        if st.button(
            "🛠️ Repair File",
            type="primary",
            use_container_width=True
        ):

            try:

                with st.spinner(
                    "Inspecting and attempting to repair your file..."
                ):

                    if file_type in ["PDF", "PDF (signature warning)"]:

                        repaired_bytes, report = repair_pdf(
                            original_bytes
                        )

                        repaired_name = (
                            os.path.splitext(file_name)[0]
                            + "_repaired.pdf"
                        )

                        mime = "application/pdf"

                    elif file_type == "Word DOCX":

                        repaired_bytes, report = repair_word(
                            original_bytes
                        )

                        repaired_name = (
                            os.path.splitext(file_name)[0]
                            + "_repaired.docx"
                        )

                        mime = (
                            "application/vnd.openxmlformats-officedocument."
                            "wordprocessingml.document"
                        )

                    elif file_type == "Excel XLSX":

                        repaired_bytes, report = repair_excel(
                            original_bytes
                        )

                        repaired_name = (
                            os.path.splitext(file_name)[0]
                            + "_repaired.xlsx"
                        )

                        mime = (
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        )

                    elif file_type == "PowerPoint PPTX":

                        repaired_bytes, report = repair_powerpoint(
                            original_bytes
                        )

                        repaired_name = (
                            os.path.splitext(file_name)[0]
                            + "_repaired.pptx"
                        )

                        mime = (
                            "application/vnd.openxmlformats-officedocument."
                            "presentationml.presentation"
                        )

                    else:
                        raise ValueError("Unsupported file type.")

                    # Verify the generated output can be reopened

                    if file_type in ["PDF", "PDF (signature warning)"]:
                        PdfReader(io.BytesIO(repaired_bytes))

                    elif file_type == "Word DOCX":
                        Document(io.BytesIO(repaired_bytes))

                    elif file_type == "Excel XLSX":
                        load_workbook(
                            io.BytesIO(repaired_bytes),
                            read_only=True
                        ).close()

                    elif file_type == "PowerPoint PPTX":
                        Presentation(io.BytesIO(repaired_bytes))

                    st.success(
                        "File successfully processed and output verified!"
                    )

                    st.header("📊 Repair Report")

                    st.json(report)

                    st.write(
                        "**Original file:**", file_name
                    )

                    st.write(
                        "**Repaired file:**", repaired_name
                    )

                    st.write(
                        "**Original format preserved:**",
                        file_type
                    )

                    st.write(
                        "**Output size:**",
                        f"{len(repaired_bytes) / 1024:.2f} KB"
                    )

                    st.download_button(
                        label="⬇️ Download Repaired File",
                        data=repaired_bytes,
                        file_name=repaired_name,
                        mime=mime,
                        use_container_width=True
                    )

                    st.caption(
                        "Your original uploaded file has not been changed. "
                        "Compare the repaired copy with your original "
                        "before relying on it."
                    )

            except Exception as e:

                st.error(
                    f"Repair was unsuccessful: {str(e)}"
                )

                st.warning(
                    "The file may be severely damaged, encrypted, "
                    "or contain unsupported features. "
                    "Your original file remains unchanged."
                )


# ---------------- FOOTER ----------------

st.divider()

st.caption(
    "🛠️ DataFix | Universal File Identification & Repair Center"
)

st.caption(
    "PDF • Word • Excel • PowerPoint | "
    "Original uploads are preserved."
)
