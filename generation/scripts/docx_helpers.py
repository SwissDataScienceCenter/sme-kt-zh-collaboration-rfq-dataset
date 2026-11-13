"""Shared helpers for rendering Fronalp documents (Swiss German)."""
import shutil, subprocess, tempfile
from pathlib import Path
import yaml
from docx import Document
from docx.shared import Pt, Cm, RGBColor

ROOT = Path(__file__).resolve().parents[2]
ORG = yaml.safe_load(open(ROOT / "generation/world/organisation.yaml", encoding="utf-8"))
C = ORG["company"]
RED = RGBColor(0x8C, 0x1D, 0x18)

def chf(x, dec=2):
    """Swiss number format with apostrophe: 12'345.60"""
    s = f"{float(x):,.{dec}f}"
    return s.replace(",", "'")

def new_doc(title_line=None, header=True):
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = "Arial"; st.font.size = Pt(10)
    for sec in doc.sections:
        sec.left_margin = sec.right_margin = Cm(2.2); sec.top_margin = Cm(2.0)
    if header:
        h = doc.sections[0].header.paragraphs[0]
        h.text = f'{C["name"]}  |  {C["street"]}, {C["zip"]} {C["city"]}  |  {C["phone_main"]}  |  {C["web"]}'
        h.runs[0].font.size = Pt(7.5)
    if title_line:
        p = doc.add_paragraph(); r = p.add_run("FRONALP"); r.bold = True; r.font.size = Pt(18); r.font.color.rgb = RED
        r = p.add_run("  Metall & Gebäudehülle"); r.font.size = Pt(10)
    return doc

def table(doc, header, rows, widths=None, style="Table Grid", font=9):
    t = doc.add_table(rows=1, cols=len(header)); t.style = style
    for c, h in zip(t.rows[0].cells, header):
        c.text = ""; r = c.paragraphs[0].add_run(str(h)); r.bold = True; r.font.size = Pt(font)
    for row in rows:
        cells = t.add_row().cells
        for c, v in zip(cells, row):
            c.text = ""; r = c.paragraphs[0].add_run("" if v is None else str(v)); r.font.size = Pt(font)
    if widths:
        t.autofit = False
        for col, w in zip(t.columns, widths):
            col.width = Cm(w)
        for row in t.rows:
            for c, w in zip(row.cells, widths):
                c.width = Cm(w)
    return t

def to_pdf(docx_path):
    """Convert with LibreOffice in a scratch dir outside the dataset, then copy the PDF next to the source."""
    docx_path = Path(docx_path)
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / docx_path.name
        shutil.copy(docx_path, src)
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, str(src)],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        shutil.copy(Path(tmp) / (docx_path.stem + ".pdf"), docx_path.with_suffix(".pdf"))
