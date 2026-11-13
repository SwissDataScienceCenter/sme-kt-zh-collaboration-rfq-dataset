"""Render the Norm register (ADR 0004) and, if law PDFs are present, verbatim law excerpts.

company-world/regulations/Normenverzeichnis_Fronalp.xlsx
company-world/regulations/Gesetzesauszuege_*.pdf   (only when sources are in generation/sources/law/)
"""
import sys, yaml
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
sys.path.insert(0, str(Path(__file__).parent))
from docx_helpers import ROOT

R = yaml.safe_load(open(ROOT / "generation/world/norm_register.yaml", encoding="utf-8"))
OUT = ROOT / "company-world/regulations"; OUT.mkdir(parents=True, exist_ok=True)

wb = Workbook(); ws = wb.active; ws.title = "Normen"
ws.append(["Nr.", "Titel", "Kurzbeschrieb", "Relevant für", "Ausgabe"])
for n in R["norms"]:
    ws.append([n["nr"], n["title"], n["de"], ", ".join(n["lines"]), "aktuelle Ausgabe beim Herausgeber prüfen"])
ws2 = wb.create_sheet("Gesetze")
ws2.append(["SR / Kanton", "Titel", "Kurzbeschrieb", "Auszug im Ordner"])
for l in R["laws"]:
    ws2.append([l["nr"], l["title"], l["de"], l.get("out", "nein")])
for s, widths in [(ws, [20, 46, 90, 30, 34]), (ws2, [18, 52, 70, 16])]:
    for c in s[1]: c.font = Font(bold=True)
    for col, w in zip("ABCDE", widths): s.column_dimensions[col].width = w
    for row in s.iter_rows(min_row=2):
        for c in row: c.alignment = Alignment(wrap_text=True, vertical="top")
ws.cell(ws.max_row + 2, 1, "Gepflegt von N. Suter. Normtexte sind lizenziert und liegen nur in der Normensammlung im Büro (Papier).")
wb.save(OUT / "Normenverzeichnis_Fronalp.xlsx")

# verbatim excerpts: whole pages copied from the official PDFs, nothing retyped (DL-130)
import subprocess
SRC = ROOT / "generation/sources/law"
for l in R["laws"]:
    if not l.get("excerpt"):
        continue
    src = SRC / l["source"]
    if not src.exists():
        print("missing source:", l["source"]); continue
    a, b = l["pages"]
    subprocess.run(["qpdf", str(src), "--pages", ".", f"{a}-{b}", "--", str(OUT / l["out"])], check=True)
    print("excerpt:", l["out"], f"(pages {a} to {b} of {l['source']})")
print("Normenverzeichnis written")
