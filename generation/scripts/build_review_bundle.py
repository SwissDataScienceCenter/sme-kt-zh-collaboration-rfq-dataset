"""Build a text and image review bundle of one Case for QA agents (DL-161).

Usage: python3 generation/scripts/build_review_bundle.py C1
Writes _review/<id>/ (git ignored): request/, operations/, offer/, company/, past/ as .txt and .png.
The offer/ folder and the estimator costing sheet are the answer; derivability reviewers must not read them.
"""
import sys, email, subprocess, io, re, csv, shutil
from email import policy
from pathlib import Path
from docx import Document
import openpyxl, yaml
ROOT = Path(__file__).resolve().parents[2]; CID = sys.argv[1]
C = ROOT / "cases" / CID; B = ROOT / "_review" / CID
CASE = yaml.safe_load(open(C / "case.yaml", encoding="utf-8"))
for sub in ("request", "operations", "offer", "company", "past"):
    (B / sub).mkdir(parents=True, exist_ok=True)
    for f in (B / sub).iterdir(): f.write_bytes(b"") if False else None
def docx_text(x):
    d = Document(io.BytesIO(x) if isinstance(x, bytes) else x)
    return "\n".join([p.text for p in d.paragraphs] + [" | ".join(c.text for c in r.cells) for t in d.tables for r in t.rows])
def pdf_text(p): return subprocess.run(["pdftotext", "-layout", str(p), "-"], capture_output=True, text=True).stdout
def pdf_png(p, out, r=60): subprocess.run(["pdftoppm", "-r", str(r), "-png", str(p), str(out)], check=True)
def xlsx_text(p):
    wb = openpyxl.load_workbook(p); out = []
    for ws in wb.worksheets:
        out.append(f"## sheet {ws.title}")
        for row in ws.iter_rows():
            vals = [str(c.value) if c.value is not None else "" for c in row]
            cm = [f"[comment {c.coordinate}: {c.comment.text}]" for c in row if c.comment]
            out.append(" ; ".join(vals) + (" " + " ".join(cm) if cm else ""))
    return "\n".join(out)
for sub in ("request", "operations", "offer"):
    for f in sorted((C / sub).iterdir()):
        stem = f.stem
        if f.suffix == ".eml":
            m = email.message_from_bytes(f.read_bytes(), policy=policy.default)
            txt = "\n".join(f"{h}: {m[h]}" for h in ("From", "To", "Cc", "Date", "Subject") if m[h]) + "\n\n" + m.get_body(("plain",)).get_content()
            for part in m.iter_attachments():
                fn = part.get_filename(); data = part.get_content(); txt += f"\n[ATTACHMENT: {fn}]"
                if fn.endswith(".pdf") and sub != "offer":
                    tmp = B / sub / f"{stem}__{fn}"; tmp.write_bytes(data); pdf_png(tmp, B / sub / f"{stem}__{Path(fn).stem}", 75); tmp.unlink(missing_ok=True) if False else None
                elif fn.endswith(".docx"):
                    (B / sub / f"{stem}__{Path(fn).stem}.txt").write_text(docx_text(data))
                elif fn.endswith((".jpg", ".png")):
                    (B / sub / f"{stem}__{fn}").write_bytes(data)
                elif fn.endswith(".json"):
                    (B / sub / f"{stem}__{Path(fn).stem}.txt").write_text(data.decode("utf-8") if isinstance(data, bytes) else str(data))
            (B / sub / f"{stem}.txt").write_text(txt)
        elif f.suffix == ".docx": (B / sub / f"{stem}.txt").write_text(docx_text(str(f)))
        elif f.suffix == ".xlsx": (B / sub / f"{stem}.txt").write_text(xlsx_text(f))
        elif f.suffix == ".pdf": (B / sub / f"{stem}.txt").write_text(pdf_text(f)); pdf_png(f, B / sub / stem, 60)
CW = ROOT / "company-world"
(B / "company/Kalkulationsrichtlinie.txt").write_text(docx_text(str(CW / "operations/Kalkulationsrichtlinie_v3.2_2024-03.docx")))
(B / "company/Faustregeln_BUl.txt").write_text(xlsx_text(CW / "operations/Faustregeln_BUl.xlsx"))
(B / "company/Betriebshandbuch.txt").write_text(docx_text(str(CW / "operations/Betriebshandbuch_Auftragsabwicklung_v2.1_2023-05.docx")))
(B / "company/Regieansaetze_2026.txt").write_text(xlsx_text(CW / "operations/Regieansaetze_2018-2026.xlsx").split("## sheet 2025")[0])
(B / "company/Artikelstamm.csv").write_text(open(next((CW / "systems").glob("MetaBau_Export_Artikelstamm_*.csv")), encoding="cp1252").read())
for pl in sorted((CW / "operations").glob("Preisliste_Systemprodukte_*.pdf")):
    (B / "company" / f"{pl.stem}.txt").write_text(pdf_text(pl))
(B / "company/Personal.txt").write_text(xlsx_text(next((CW / "organisation").glob("MetaBau_Export_Personal_*.xlsx"))))
ids = set(re.findall(r"\b(?:KC-\d{5}|D-\d{5}|L-20\d\d-\d{3}|OF-\d\d-\d{4}|P-\d\d-\d{3})\b", (C / "case.yaml").read_text()))
ids |= {"Talkessel", "Reding"}
for pp in CASE.get("past_projects", []):                 # customers of past projects as well
    ids |= set(re.findall(r"\bD-\d{5}\b", (ROOT / "generation/past_projects" / f'{pp["nr"]}.yaml').read_text(encoding="utf-8")))
out = []
for f in sorted((CW / "systems").glob("*.csv")):
    enc = "cp1252" if f.name.startswith("MetaBau") else "utf-8-sig"
    lines = open(f, encoding=enc).read().splitlines()
    hits = [l for l in lines[1:] if any(i in l for i in ids)]
    if hits: out.append(f"## {f.name}\n{lines[0]}\n" + "\n".join(hits))
(B / "company/CRM_ERP_relevant_rows.txt").write_text("\n\n".join(out))
(B / "company/Subunternehmer_und_Partner.txt").write_text(xlsx_text(CW / "operations/Subunternehmer_und_Partner.xlsx"))
(B / "company/AGB.txt").write_text(docx_text(str(CW / "operations/AGB_Fronalp_2022-08.docx")))
for f in sorted((CW / "operations/templates").glob("*.docx")): (B / "company" / f"{f.stem}.txt").write_text(docx_text(str(f)))
if CASE.get("business_line") == "advisory":       # estimates come from past advisory work (ADR 0007)
    adv = []
    for f in sorted((CW / "systems").glob("MetaBau_Export_*.csv")):
        lines = open(f, encoding="cp1252").read().splitlines()
        hits = [l for l in lines[1:] if ";BERA;" in l]
        if f.name.startswith("MetaBau_Export_Stunden"):
            pnr = {l.split(";")[0] for l in open(next((CW / "systems").glob("MetaBau_Export_Projekte_*.csv")), encoding="cp1252").read().splitlines() if ";BERA;" in l}
            hits = [l for l in lines[1:] if l.split(";")[0] in pnr]
        if hits: adv.append(f"## {f.name} (Beratung)\n{lines[0]}\n" + "\n".join(hits))
    (B / "company/ERP_Beratung_alle.txt").write_text("\n\n".join(adv))
    (B / "company/Regieansaetze_alle_Jahre.txt").write_text(xlsx_text(CW / "operations/Regieansaetze_2018-2026.xlsx"))
for pp in CASE.get("past_projects", []):
    pdir = next((CW / "past-projects").glob(pp["nr"] + "*"))
    for f in sorted(pdir.rglob("*")):
        if f.suffix == ".pdf": (B / "past" / f"{f.stem}.txt").write_text(pdf_text(f)); pdf_png(f, B / "past" / f.stem, 50)
        elif f.suffix == ".xlsx": (B / "past" / f"{f.stem}.txt").write_text(xlsx_text(f))
for f in B.rglob("*.pdf"): f.rename(f.with_suffix(".pdf.bin")) if False else None
print(sum(1 for f in B.rglob("*") if f.is_file() and f.suffix in (".txt", ".png", ".jpg", ".csv")), "files in", B)
