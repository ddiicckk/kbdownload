import os
import re
import uuid
import pandas as pd
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ==========================
# Configuration
# ==========================
EXCEL_PATH = "kb_knowledge.xlsx"     # <-- set to your Excel file path
SHEET_NAME = 0                       # index or sheet name
COL_ID = "Number"                    # column that contains the file name
COL_HTML = "Article body"            # column that contains the HTML
OUTPUT_DIR = "exported_docs"         # output folder for .docx files

# ==========================
# Helpers
# ==========================
def sanitize_filename(name: str) -> str:
    """
    Make a safe filename for most OSes.
    """
    name = str(name).strip()
    # Replace forbidden characters on Windows and macOS
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    # Collapse whitespace
    name = re.sub(r"\s+", " ", name)
    # Truncate to a reasonable length
    return name[:200] if len(name) > 200 else name

def add_html_as_altchunk(doc: Document, html_text: str) -> None:
    """
    Embed HTML in a .docx so Word renders it when opened using the 'altChunk' technique.
    This creates an HTML part in the package and inserts an <w:altChunk> reference.

    Notes:
    - Word renders the HTML on first open (you might see a brief 'Converting file...' message).
    - The HTML becomes part of the document content once saved from Word.
    """
    # Create a unique partname for the HTML
    partname = f"/word/htmlDoc_{uuid.uuid4().hex}.html"
    # Add the HTML as a new part in the package
    # (this returns a relationship id we will reference in the altChunk element)
    html_part = doc.part.package.part_factory(
        partname, "text/html", html_text.encode("utf-8")
    )
    rel = doc.part.relate_to(html_part.partname, RT.A_F_CHUNK)

    # Create the altChunk element and attach the relationship id
    alt_chunk = OxmlElement("w:altChunk")
    alt_chunk.set(qn("r:id"), rel.rId)

    # Insert it into the document body (at the end)
    doc._body._element.append(alt_chunk)

# ==========================
# Main
# ==========================
def main():
    # Ensure output directory exists
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Read Excel using openpyxl engine
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, engine="openpyxl")

    # Validate necessary columns
    missing = [c for c in (COL_ID, COL_HTML) if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required column(s) in Excel: {missing}")

    # Iterate rows and export
    exported = 0
    for idx, row in df.iterrows():
        file_id = row.get(COL_ID)
        html = row.get(COL_HTML)

        # Skip rows without file name or content
        if pd.isna(file_id) or pd.isna(html) or str(html).strip() == "":
            continue

        safe_name = sanitize_filename(file_id)
        out_path = os.path.join(OUTPUT_DIR, f"{safe_name}.docx")

        # Create a new Word document and inject the HTML
        doc = Document()
        # Optional: title or header (comment out if not needed)
        # doc.add_heading(str(file_id), level=1)

        add_html_as_altchunk(doc, str(html))
        doc.save(out_path)
        exported += 1

    print(f"Export completed. Files created: {exported} in '{OUTPUT_DIR}'.")

if __name__ == "__main__":
    main()
