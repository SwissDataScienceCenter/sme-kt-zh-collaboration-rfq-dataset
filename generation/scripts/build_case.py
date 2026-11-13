"""Build one Case from cases/<id>/case.yaml and generation/cases/<id>/content.yaml (ADR 0006).

Steps: price the Offer (pricing.py) and write 'expected' into case.yaml, render past projects (ADR 0005),
render request, operations and offer documents, register CRM/ERP rows in systems_additions.yaml, validate.
Usage: python3 generation/scripts/build_case.py C1
"""
import datetime as dt, io, math, re, subprocess, sys, tempfile
from decimal import Decimal
from pathlib import Path
import yaml
sys.path.insert(0, str(Path(__file__).parent))
from pricing import World, compute_offer, D, round_to
from docx_helpers import ROOT, ORG, C, new_doc, table, to_pdf, chf
import renderers as R
import fee_offer as FO
import lump_offer as LO
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.comments import Comment

W = World(); OPS = W.ops
CID = sys.argv[1]
CASE_DIR = ROOT / "cases" / CID
CASE = yaml.safe_load(open(CASE_DIR / "case.yaml", encoding="utf-8"))
CONTENT = yaml.safe_load(open(ROOT / f"generation/cases/{CID}/content.yaml", encoding="utf-8"))
CUST = {r["id"]: r for r in __import__("csv").DictReader(open(ROOT / "generation/world/customers.csv", encoding="utf-8"))}
for sub in ("request", "operations", "offer"):
    (CASE_DIR / sub).mkdir(parents=True, exist_ok=True)

MANUAL_TEXT = {
    "H1": "Die in dieser Offerte aufgeführten Leistungen und Preise beruhen auf Annahmen, da keine Besichtigung vor Ort "
          "stattgefunden hat. Die Kalkulation wurde auf Basis von Erfahrungswerten und dem vorhandenen Plan erstellt. "
          "Zeigen sich bei der Ausführung abweichende Verhältnisse, behalten wir uns entsprechende Anpassungen der "
          "Leistungen und Kosten vor."}

def person(key):
    """'RBe' style initials or a key from content.people -> 'Name <mail>'"""
    if key in R.STAFF:
        s = R.STAFF[key]; return f'{s["first"]} {s["last"]} <{R.staff_mail(key)}>'
    return CONTENT["people"].get(key, key)

def priced_sections(sections, date, world=W):
    flat, out = [], []
    for sec in sections:
        rows = []
        for ln in sec["lines"]:
            it = world.items.get(ln.get("code", ""), {})
            row = dict(ln)
            if isinstance(row.get("unit_price"), str):    # price list rule, e.g. "rf(0.18 * price('97.200.01'))" (S2)
                y_ = int(str(date)[:4])
                env_ = {"price": lambda c, y_=y_: float(world.price(c, y_)), "rf": lambda x: math.floor(x + 0.5 + 1e-9), "F": PRICE_F[0], "ceil": math.ceil}
                row["unit_price"] = eval(row["unit_price"], {"__builtins__": {}}, env_)
            row.setdefault("text", it.get("text", ""))
            row.setdefault("unit", it.get("unit", ""))
            rows.append(row); flat.append(row)
        out.append({"title": sec["title"], "lines": rows})
    res = compute_offer(flat, date, discount=CASE_OFFER_DISCOUNT[0], skonto=CASE_OFFER_DISCOUNT[1], world=world)
    # put computed prices back into section rows (same order)
    it = iter(res["lines"])
    for sec in out:
        sec["lines"] = [next(it) for _ in sec["lines"]]
    return res, out

CASE_OFFER_DISCOUNT = [0, False]
PRICE_F = [None]                                     # facts available to price list rule expressions

def _facts_early(f):
    """Derived facts for costing before the offer is priced (same rules as derive_facts below)."""
    F = type("NS", (dict,), {"__getattr__": dict.get})(f)
    for k, expr in (CASE.get("derived") or {}).items():
        F[k] = eval(expr, {"__builtins__": {}, "ceil": math.ceil, "floor": math.floor, "round": round, "sqrt": math.sqrt, "max": max, "min": min}, {"F": F})
    return F

# ---------------------------------------------------------------------------------------------- 0 lump sum costing (DL-189)
off = CASE["offer"]; date = str(CASE["dates"]["offer"])
LUMP = None
PRICE_F[0] = _facts_early(CASE["facts"])
if off.get("format") == "word_lump":
    _F = _facts_early(CASE["facts"]); _year = int(date[:4])
    LUMP = {k: LO.cost(v, _F, W, _year) for k, v in CASE["costing"].items()}
    main, alts, opt = off["main_variant"], off.get("alternatives", []), off.get("option")
    lines = [{"text": CASE["costing"][main]["title"], "qty": 1, "unit": "Pau", "unit_price": LUMP[main]["lump"], "src": f"costing {main}"}]
    if opt: lines.append({"text": CASE["costing"][opt]["title"], "qty": 1, "unit": "Pau", "unit_price": LUMP[opt]["lump"], "optional": True, "src": f"costing {opt}"})
    off["sections"] = [{"title": "Pauschalpreis", "lines": lines}]

# ---------------------------------------------------------------------------------------------- 1 price the Offer
CASE_OFFER_DISCOUNT[:] = [off.get("discount", 0), off.get("skonto", False)]
result, sections = priced_sections(off["sections"], date)
expected = {"gross": float(result["gross"]), "discount": float(result["discount"]), "skonto": float(result["skonto"]),
            "net": float(result["net"]), "vat_rate": float(result["vat_rate"]), "vat": float(result["vat"]),
            "total": float(result["total"]), "optional_sum": float(result["optional_sum"]),
            "needs_owner_approval": bool(result["needs_owner_approval"]),
            "lines": [{"pos": f"{si}.{li:02d}", "code": l.get("code", ""), "qty": float(l["qty"]), "unit": l["unit"],
                       "unit_price": float(l["unit_price"]), "total": float(l["total"]), "optional": bool(l.get("optional", False))}
                      for si, s in enumerate(sections, 1) for li, l in enumerate(s["lines"], 1)]}
txt = (CASE_DIR / "case.yaml").read_text(encoding="utf-8")
block = "expected:                      # filled by build_case.py via pricing.py, never by hand\n" + \
        "\n".join("  " + l for l in yaml.safe_dump(expected, allow_unicode=True, sort_keys=False, width=200).splitlines()) + "\n"
txt = re.sub(r"^expected:.*?(?=^\S)", block + "\n", txt, flags=re.S | re.M)
(CASE_DIR / "case.yaml").write_text(txt, encoding="utf-8")

# ---------------------------------------------------------------------------------------------- 2 offer PDF
payment_text = OPS["commercial_rules"]["payment_terms"].get(off["payment_terms"], off["payment_terms"])
off_r = dict(off, valid_until=CASE["dates"]["valid_until"], payment_text=payment_text,
             hint_text=MANUAL_TEXT.get(off.get("hint", ""), ""), client_provisions=CASE["client_provisions"],
             reservations=CASE["reservations"], assumptions=CASE.get("assumptions", []), description=off.get("description", []))
offer_pdf = CASE_DIR / "offer" / off["file_name"]
FEE_PRINTED = None
def fee_rates(year):
    return {k: W.rate(c, year) for k, c in (("B1", "advisory_owner"), ("B2", "advisory_pm"), ("B3", "advisory_technician"), ("MA", "fitter_a"))}
if off.get("format") == "word_fee_offer":         # Advisory, Word template (DL-184)
    fo = dict(off_r, texts=CONTENT["fee_offer"], exclusions=CASE.get("exclusions", []), payment_terms=off["payment_terms"])
    FEE_PRINTED = FO.render(offer_pdf.with_suffix(".docx"), fo, result, sections, date, fee_rates(int(date[:4])))
elif off.get("format") == "word_lump":            # Metal construction lump sum, Word template 2016 (DL-189)
    vrows = []
    for k in [off["main_variant"]] + off.get("alternatives", []):
        r_ = compute_offer([{"text": "x", "qty": 1, "unit": "Pau", "unit_price": LUMP[k]["lump"]}], date, world=W)
        vrows.append((k, CASE["costing"][k]["title"], r_["net"], r_["vat"], r_["total"], CONTENT["lump_offer"]["variant_scope"][k]))
        expected.setdefault("alternatives", {})[k] = {"lump": float(LUMP[k]["lump"]), "vat": float(r_["vat"]), "total": float(r_["total"]),
                                                      "cost": float(LUMP[k]["cost"]), "risk_profit": float(LUMP[k]["risk_profit"])}
    optk = off.get("option")
    lo = dict(off, vat_rate=result["vat_rate"], validity_text=f'3 Monate, bis {R.de_date(CASE["dates"]["valid_until"])}' + (f'; Preise gültig für Ausführung bis {off["price_hold"]}' if off.get("price_hold") else ''),
              payment_text=payment_text, client_provisions=CASE["client_provisions"], exclusions=CASE.get("exclusions", []),
              assumptions=CASE.get("assumptions", []), reservations=CASE.get("reservations", []))
    FEE_PRINTED = LO.render(offer_pdf.with_suffix(".docx"), lo, vrows, (CASE["costing"][optk]["title"], LUMP[optk]["lump"]) if optk else None,
                            date, CONTENT["lump_offer"])
    if CONTENT.get("costing_sheet"):
        LO.sheet(CASE_DIR / CONTENT["costing_sheet"]["file"], {k: (CASE["costing"][k]["title"], LUMP[k]) for k in LUMP},
                 f'Kalkulation {off["number"]}', CONTENT["costing_sheet"]["by"], _F)
    txt = (CASE_DIR / "case.yaml").read_text(encoding="utf-8")          # store variant results in expected
    block = "expected:                      # filled by build_case.py via pricing.py, never by hand\n" + \
            "\n".join("  " + l for l in yaml.safe_dump(expected, allow_unicode=True, sort_keys=False, width=200).splitlines()) + "\n"
    txt = re.sub(r"^expected:.*?(?=^\S)", block + "\n", txt, flags=re.S | re.M)
    (CASE_DIR / "case.yaml").write_text(txt, encoding="utf-8")
else:
    R.metabau_offer(offer_pdf, off_r, result, sections, date, layout="2024")

# ---------------------------------------------------------------------------------------------- 2a cost estimate sent earlier (A3, Richtofferte as input document)
CE_NET = None
if CONTENT.get("cost_estimate"):
    ce = CONTENT["cost_estimate"]; ce_date = str(ce["date"])
    CASE_OFFER_DISCOUNT[:] = [ce.get("discount", 0), False]
    ce_res, ce_secs = priced_sections(ce["sections"], ce_date)
    CASE_OFFER_DISCOUNT[:] = [off.get("discount", 0), off.get("skonto", False)]
    ce_o = dict(ce, kind="Richtofferte", texts=ce["texts"])
    ce_printed = FO.render(CASE_DIR / ce["file"], ce_o, ce_res, ce_secs, ce_date, fee_rates(int(ce_date[:4])))
    _ce_txt = subprocess.run(["pdftotext", "-layout", str((CASE_DIR / ce["file"]).with_suffix(".pdf")), "-"], capture_output=True, text=True).stdout
    CE_MISSING = [a for a in ce_printed if a not in _ce_txt]
    CE_NET = float(ce_res["net"]); expected["cost_estimate"] = {"number": ce["number"], "gross": float(ce_res["gross"]), "net": CE_NET, "total": float(ce_res["total"])}
    print(f'cost estimate {ce["number"]}: gross {chf(ce_res["gross"])}, net {chf(ce_res["net"])}, total {chf(ce_res["total"])}')
    txt = (CASE_DIR / "case.yaml").read_text(encoding="utf-8")
    block = "expected:                      # filled by build_case.py via pricing.py, never by hand\n" + \
            "\n".join("  " + l for l in yaml.safe_dump(expected, allow_unicode=True, sort_keys=False, width=200).splitlines()) + "\n"
    (CASE_DIR / "case.yaml").write_text(re.sub(r"^expected:.*?(?=^\S)", block + "\n", txt, flags=re.S | re.M), encoding="utf-8")

# ---------------------------------------------------------------------------------------------- 3 plan scan (attachment)
class _NS(dict):
    __getattr__ = dict.get

def derive_facts(f):
    F = _NS(f)
    if "terrace_main" in f:                                  # L-shaped terrace (C1)
        (a, b), (d, c) = (f["terrace_main"]["width"], f["terrace_main"]["depth"]), (f["terrace_wing"]["width"], f["terrace_wing"]["depth"])
        F.area_exact = round(a * b + d * c, 2); F.area = F.area_exact
        F.edge_open = round(a + d + b + c, 2)                # south edge + west edge + east edge
        F.edge_facade = round(a + (c - b) + d, 2)            # facade behind main, step, facade behind wing
    for k, expr in (CASE.get("derived") or {}).items():      # generic derived values (DL-171)
        F[k] = eval(expr, {"__builtins__": {}, "ceil": math.ceil, "floor": math.floor, "round": round, "sqrt": math.sqrt, "max": max, "min": min}, {"F": F})
    return F
F = derive_facts(CASE["facts"])

def check_plan(pl, F):
    """Geometric and topological consistency of the floor plan (DL-157). Returns summary or raises."""
    errs = []; rooms = pl["rooms"]
    a, b = F["terrace_main"]["width"], F["terrace_main"]["depth"]; d, c = F["terrace_wing"]["width"], F["terrace_wing"]["depth"]
    top = b + F["attic_depth_behind_main"]
    def ov(r1, r2):
        return max(0, min(r1[2], r2[2]) - max(r1[0], r2[0])) * max(0, min(r1[3], r2[3]) - max(r1[1], r2[1]))
    names = list(rooms)
    for i, n1 in enumerate(names):
        for n2 in names[i + 1:]:
            if ov(rooms[n1], rooms[n2]) > 1e-6: errs.append(f"rooms overlap: {n1} / {n2}")
    total = sum((r[2] - r[0]) * (r[3] - r[1]) for r in rooms.values())
    footprint = a * (top - b) + d * (top - c)
    if abs(total - footprint) > 0.01: errs.append(f"rooms cover {total:.2f} m2, footprint is {footprint:.2f} m2")
    def on_edge(r, axis, at, span):
        lo, hi = (r[1], r[3]) if axis == "x" else (r[0], r[2])
        on = (abs(r[0] - at) < 1e-6 or abs(r[2] - at) < 1e-6) if axis == "x" else (abs(r[1] - at) < 1e-6 or abs(r[3] - at) < 1e-6)
        return on and lo - 1e-6 <= span[0] < span[1] <= hi + 1e-6
    graph = {n: set() for n in names}
    for dr in pl["doors"]:
        r1, r2 = dr["rooms"]
        if not (on_edge(rooms[r1], dr["axis"], dr["at"], dr["span"]) and on_edge(rooms[r2], dr["axis"], dr["at"], dr["span"])):
            errs.append(f"door not on a wall shared by {r1} and {r2}: {dr}")
        graph[r1].add(r2); graph[r2].add(r1)
    seen, todo = set(), [pl["entrance"]]
    while todo:
        n = todo.pop(); seen.add(n); todo += [m for m in graph[n] if m not in seen]
    for n in names:
        if n not in seen: errs.append(f"room not reachable: {n}")
        if not graph[n] and n not in pl.get("common", []): errs.append(f"room without door: {n}")
    for t in pl["terrace_doors"]:
        r = rooms[t["room"]]; edge_y = b if r[2] <= a + 1e-6 else c
        if not (t["axis"] == "y" and abs(t["at"] - edge_y) < 1e-6 and on_edge(r, "y", t["at"], t["span"])):
            errs.append(f"terrace door not on terrace edge: {t}")
    outline_x = {0.0, a + d}; outline_y = {top, b, c}
    for w in pl["windows"]:
        ok = on_edge(rooms[w["room"]], w["axis"], w["at"], w["span"]) and (w["at"] in (outline_x if w["axis"] == "x" else outline_y))
        if not ok: errs.append(f"window not on exterior wall: {w}")
    if errs: raise SystemExit("PLAN ERRORS:\n  " + "\n  ".join(errs))
    return _NS(terrace_doors=len(pl["terrace_doors"]), drains=len(pl["drains"]), rooms=len(rooms))

def plan_pdf():
    """Attic floor plan 1:100 drawn from the layout data in content.yaml, then scanned (DL-154, DL-157)."""
    import matplotlib; matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import Polygon, Rectangle, Arc
    pl = CONTENT["plan"]; rooms = pl["rooms"]
    a, b = F["terrace_main"]["width"], F["terrace_main"]["depth"]; d, c = F["terrace_wing"]["width"], F["terrace_wing"]["depth"]
    top = b + F["attic_depth_behind_main"]; W_EXT, W_INT = 0.30, 0.12
    fig, ax = plt.subplots(figsize=(16.54, 11.69)); ax.set_aspect("equal"); ax.axis("off")
    def wall(x0, y0, x1, y1, t):
        if abs(y0 - y1) < 1e-9: ax.add_patch(Rectangle((min(x0, x1) - t / 2, y0 - t / 2), abs(x1 - x0) + t, t, fc="black", lw=0))
        else:                   ax.add_patch(Rectangle((x0 - t / 2, min(y0, y1) - t / 2), t, abs(y1 - y0) + t, fc="black", lw=0))
    def gap(axis, at, span, t):
        if axis == "y": ax.add_patch(Rectangle((span[0], at - t / 2 - 0.03), span[1] - span[0], t + 0.06, fc="white", lw=0))
        else:           ax.add_patch(Rectangle((at - t / 2 - 0.03, span[0]), t + 0.06, span[1] - span[0], fc="white", lw=0))
    def glazing(axis, at, span, t):
        for k in (-t / 2, 0, t / 2):
            if axis == "y": ax.plot(span, [at + k] * 2, color="black", lw=0.6)
            else:           ax.plot([at + k] * 2, span, color="black", lw=0.6)
    # terrace with parapet
    ax.add_patch(Polygon([(0, 0), (a + d, 0), (a + d, c), (a, c), (a, b), (0, b)], closed=True, fc="#f2f2f2", ec="black", lw=0.8, hatch=".."))
    for seg in ((0, 0, a + d, 0), (0, 0, 0, b), (a + d, 0, a + d, c)): wall(*seg, 0.18)
    # walls: every room edge interior, building outline and staircase exterior
    for n, (x0, y0, x1, y1) in rooms.items():
        t = W_EXT if n in pl.get("common", []) else W_INT
        for seg in ((x0, y0, x1, y0), (x1, y0, x1, y1), (x0, y1, x1, y1), (x0, y0, x0, y1)): wall(*seg, t)
    for seg in ((0, b, a, b), (a, b, a, c), (a, c, a + d, c), (a + d, c, a + d, top), (0, top, a + d, top), (0, b, 0, top)): wall(*seg, W_EXT)
    for dr in pl["doors"]:
        r1 = rooms[dr["rooms"][0]]; t = W_EXT if any(r in pl.get("common", []) for r in dr["rooms"]) else W_INT
        gap(dr["axis"], dr["at"], dr["span"], t); w = dr["span"][1] - dr["span"][0]
        if dr["axis"] == "y":
            into_up = r1[1] >= dr["at"] - 1e-6            # door swings into the first room
            sgn = 1 if into_up else -1
            ax.add_patch(Arc((dr["span"][0], dr["at"]), 2 * w, 2 * w, theta1=0 if sgn > 0 else 270, theta2=90 if sgn > 0 else 360, lw=0.6))
            ax.plot([dr["span"][0]] * 2, [dr["at"], dr["at"] + sgn * w], color="black", lw=0.9)
        else:
            into_right = r1[0] >= dr["at"] - 1e-6
            sgn = 1 if into_right else -1
            ax.add_patch(Arc((dr["at"], dr["span"][0]), 2 * w, 2 * w, theta1=0 if sgn > 0 else 90, theta2=90 if sgn > 0 else 180, lw=0.6))
            ax.plot([dr["at"], dr["at"] + sgn * w], [dr["span"][0]] * 2, color="black", lw=0.9)
    for wd in pl["terrace_doors"] + pl["windows"]:
        gap(wd["axis"], wd["at"], wd["span"], W_EXT); glazing(wd["axis"], wd["at"], wd["span"], W_EXT)
    for wd in pl["terrace_doors"]:              # window doors marked FT with threshold line
        mid = (wd["span"][0] + wd["span"][1]) / 2
        ax.plot(wd["span"], [wd["at"] - W_EXT / 2 - 0.08] * 2, color="black", lw=1.4)
        ax.text(mid, wd["at"] - 0.55, "FT", ha="center", fontsize=6.5)
    # staircase and lift
    tx0, ty0, tx1, ty1 = rooms["Treppenhaus"]
    for k in range(9): ax.plot([tx0 + 0.2, tx0 + 1.6], [ty0 + 0.2 + k * 0.24] * 2, color="black", lw=0.5)
    ax.annotate("", (tx0 + 0.9, ty0 + 2.1), (tx0 + 0.9, ty0 + 0.25), arrowprops=dict(arrowstyle="->", lw=0.6))
    lx, ly = tx1 - 1.5, ty0 + 0.3
    ax.add_patch(Rectangle((lx, ly), 1.2, 1.4, fill=False, lw=0.7)); ax.plot([lx, lx + 1.2], [ly, ly + 1.4], lw=0.4, color="black"); ax.plot([lx, lx + 1.2], [ly + 1.4, ly], lw=0.4, color="black")
    ax.text(lx + 0.6, ly + 0.62, "Lift", ha="center", fontsize=6.5, bbox=dict(fc="white", ec="none", pad=0.5))
    # labels with areas
    for n, (x0, y0, x1, y1) in rooms.items():
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        if n == "Treppenhaus":
            ax.text(lx + 0.6, ty1 - 0.3, n, ha="center", fontsize=6.5); continue
        def inner(lo, hi, axis):     # net dimension: half wall thickness off each side (exterior 0.30, interior 0.12)
            def t(v):
                ext = (v in (0.0, a + d)) if axis == "x" else (v in (b, c, top))
                return (W_EXT if ext else W_INT) / 2
            return (hi - lo) - t(lo) - t(hi)
        net = inner(x0, x1, "x") * inner(y0, y1, "y")
        ax.text(cx, cy + 0.15, n.split(" ")[0] if n.startswith("Zimmer") else n, ha="center", fontsize=9)
        ax.text(cx, cy - 0.35, f"{net:.1f} m2".replace(".", ","), ha="center", fontsize=7.5)
    ax.text(a / 2, b / 2 + 0.2, pl["labels"]["main"], ha="center", fontsize=10, bbox=dict(fc="white", ec="none", pad=1))
    ax.text(a + d / 2, c / 2, pl["labels"]["wing"], ha="center", fontsize=9.5, bbox=dict(fc="white", ec="none", pad=1))
    ax.text(a / 2, 0.55, "Plattenbelag in Splitt", ha="center", fontsize=7, style="italic")
    for (x, y) in pl["drains"]:
        ax.plot(x, y, marker="o", ms=8, mfc="white", mec="black"); ax.text(x + (0.3 if x < a else -0.75), y + 0.2, "DA", fontsize=7)
    for xs_, ys_ in (([0.25, a + d - 0.25], [0.25, 0.25]), ([0.25, 0.25], [0.25, b - 0.2]), ([a + d - 0.25, a + d - 0.25], [0.25, c - 0.2])):
        ax.plot(xs_, ys_, ls=(0, (4, 2)), lw=0.9, color="black")
    ax.text(a + d + 0.35, 0.3, "Geländer h = 1.00", fontsize=7, rotation=90)
    def dim(x0, y0, x1, y1, label, rot=0, off=0.0):
        ax.annotate("", (x0, y0), (x1, y1), arrowprops=dict(arrowstyle="|-|", lw=0.6, mutation_scale=3))
        ax.text((x0 + x1) / 2 + off, (y0 + y1) / 2 + (0.25 if rot == 0 else 0), label, ha="center", va="bottom" if rot == 0 else "center", fontsize=8, rotation=rot)
    dim(0, -1.0, a, -1.0, f"{a:.2f}"); dim(a, -1.0, a + d, -1.0, f"{d:.2f}"); dim(0, -2.0, a + d, -2.0, f"{a + d:.2f}")
    dim(-1.0, 0, -1.0, b, f"{b:.2f}", rot=90, off=-0.4); dim(-1.0, b, -1.0, top, f"{top - b:.2f}", rot=90, off=-0.4)
    dim(a + d + 1.2, 0, a + d + 1.2, c, f"{c:.2f}", rot=90, off=0.4)
    xs = sorted({v for r in rooms.values() if abs(r[3] - top) < 1e-6 for v in (r[0], r[2])})
    for x0, x1 in zip(xs, xs[1:]): dim(x0, top + 0.8, x1, top + 0.8, f"{x1 - x0:.2f}")
    dim(0, top + 1.8, a + d, top + 1.8, f"{a + d:.2f}")
    ax.plot([3.0, 3.0], [-0.4, top + 0.4], ls=(0, (6, 3, 1, 3)), lw=0.7, color="black"); ax.text(2.75, -0.75, "A", fontsize=9, weight="bold"); ax.text(2.75, top + 2.4, "A", fontsize=9, weight="bold")
    ax.annotate("N", (a + d + 2.6, top - 0.5), (a + d + 2.6, top - 2.0), ha="center", arrowprops=dict(arrowstyle="-|>", lw=1), fontsize=11)
    for k in range(5): ax.add_patch(Rectangle((k, -3.4), 1, 0.25, fc="black" if k % 2 == 0 else "white", ec="black", lw=0.5))
    ax.text(0, -3.9, "0", fontsize=7); ax.text(5, -3.9, "5 m", fontsize=7)
    x0 = a + d - 7.8; ax.add_patch(Rectangle((x0, -7.4), 7.8, 4.6, fill=False, lw=1))
    for i, t in enumerate(pl["title"]): ax.text(x0 + 0.2, -3.3 - i * 0.55, t, fontsize=9.5 if i == 0 else 8, weight="bold" if i == 0 else "normal")
    ax.text(0, top + 3.0, "ATTIKAGESCHOSS", fontsize=14, weight="bold")
    ax.set_xlim(-2.5, a + d + 3.5); ax.set_ylim(-8, top + 4)
    with tempfile.TemporaryDirectory() as tmp:
        raw = Path(tmp) / "plan.pdf"; fig.savefig(raw); plt.close(fig)
        out = Path(tmp) / "scan.pdf"; R.scanify(raw, out, seed=1994, dpi=110)
        return out.read_bytes()
def roof_plan_pdf():
    """Roof plan (Dachaufsicht) of an old building, drawn from facts and content.plan, then scanned yellowed (A1)."""
    import matplotlib; matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import Rectangle, Circle
    pl = CONTENT["plan"]; w, dpt, gapx = F.roof_width, F.roof_depth, pl["gap"]
    fig, ax = plt.subplots(figsize=(16.54, 11.69)); ax.set_aspect("equal"); ax.axis("off")
    hand = dict(fontfamily="DejaVu Sans Mono")
    def dim(x0, y0, x1, y1, label, rot=0, off=0.0):
        ax.annotate("", (x0, y0), (x1, y1), arrowprops=dict(arrowstyle="|-|", lw=0.7, mutation_scale=4))
        ax.text((x0 + x1) / 2 + off, (y0 + y1) / 2 + (0.35 if rot == 0 else 0), label, ha="center", va="bottom" if rot == 0 else "center",
                fontsize=10, rotation=rot, **hand)
    for i, house in enumerate(pl["houses"]):
        x0 = i * (w + gapx)
        ax.add_patch(Rectangle((x0, 0), w, dpt, fill=False, lw=2.2))
        ax.add_patch(Rectangle((x0 + 0.3, 0.3), w - 0.6, dpt - 0.6, fill=False, lw=0.6, ls=(0, (5, 3))))
        ax.text(x0 + w / 2, dpt / 2 + 1.2, house["name"], ha="center", fontsize=15, weight="bold", **hand)
        ax.text(x0 + w / 2, dpt / 2 - 0.2, "Flachdach Kies", ha="center", fontsize=10, **hand)
        for (dx, dy) in house["drains"]:
            ax.add_patch(Circle((x0 + dx, dy), 0.35, fill=False, lw=1)); ax.text(x0 + dx + 0.5, dy - 0.15, "DE", fontsize=8, **hand)
        hx, hy = house["hatch"]; ax.add_patch(Rectangle((x0 + hx, hy), 1.0, 1.0, fill=False, lw=1.2, hatch="//"))
        ax.text(x0 + hx + 1.2, hy + 0.3, "Dachluke", fontsize=8, **hand)
        for (vx, vy, t) in house.get("vents", []):
            ax.add_patch(Rectangle((x0 + vx, vy), 0.6, 0.6, fc="black")); ax.text(x0 + vx + 0.8, vy + 0.1, t, fontsize=8, **hand)
        dim(x0, -1.4, x0 + w, -1.4, f"{w:.2f}"); dim(x0 - 1.4, 0, x0 - 1.4, dpt, f"{dpt:.2f}", rot=90, off=-0.9)
    dim(w, -1.4, w + gapx, -1.4, f"{gapx:.2f}")
    ax.annotate("N", (2 * w + gapx + 2.5, dpt - 1), (2 * w + gapx + 2.5, dpt - 4), ha="center", arrowprops=dict(arrowstyle="-|>", lw=1.2), fontsize=13)
    tx = 2 * w + gapx - 12.0; ax.add_patch(Rectangle((tx, -9.6), 12.0, 6.6, fill=False, lw=1.2))
    for k, t in enumerate(pl["title"]):
        ax.text(tx + 0.4, -3.9 - k * 0.8, t, fontsize=12 if k == 0 else 9.5, weight="bold" if k == 0 else "normal", **hand)
    ax.text(0, dpt + 2.0, pl["heading"], fontsize=16, weight="bold", **hand)
    ax.set_xlim(-3, 2 * w + gapx + 4); ax.set_ylim(-10.5, dpt + 4)
    with tempfile.TemporaryDirectory() as tmp:
        raw = Path(tmp) / "plan.pdf"; fig.savefig(raw); plt.close(fig)
        out = Path(tmp) / "scan.pdf"; R.scanify(raw, out, seed=1974, dpi=100)
        return out.read_bytes()

def siteplan_pdf():
    """Situationsplan as printed from a map portal: parcels, buildings, planned object, altitude spot (S2)."""
    import matplotlib; matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import Polygon, Rectangle
    sp = CONTENT["siteplan"]
    fig, ax = plt.subplots(figsize=(8.27, 11.69)); ax.set_aspect("equal"); ax.axis("off")
    for poly in sp.get("neighbours", []):
        ax.add_patch(Polygon(poly["xy"], closed=True, fill=False, lw=0.6, ec="#777777"))
        ax.text(*poly["label_xy"], poly["label"], fontsize=7, color="#777777", ha="center")
    ax.add_patch(Polygon(sp["parcel"]["xy"], closed=True, fc="#fff6d8", ec="#c03030", lw=1.6))
    ax.text(*sp["parcel"]["label_xy"], sp["parcel"]["label"], fontsize=9, color="#c03030", ha="center", weight="bold")
    for (x, y, w, h, t) in sp["buildings"]:
        ax.add_patch(Rectangle((x, y), w, h, fc="#d9d9d9", ec="black", lw=0.9)); ax.text(x + w / 2, y + h / 2, t, ha="center", va="center", fontsize=8)
    for (x, y, w, h, t) in sp.get("planned", []):
        ax.add_patch(Rectangle((x, y), w, h, fill=False, ec="#1f4e9c", lw=1.4, ls=(0, (4, 2))))
        ax.text(x + w / 2, y + h / 2, t, ha="center", va="center", fontsize=8, color="#1f4e9c")
    xs, ys = zip(*sp["road"]["xy"]); ax.plot(xs, ys, color="#999999", lw=9, solid_capstyle="butt", zorder=0)
    ax.text(*sp["road"]["label_xy"], sp["road"]["label"], fontsize=8, rotation=sp["road"].get("rot", 0), ha="center")
    x, y, t = sp["spot"]; ax.plot(x, y, marker="+", color="black", ms=8); ax.text(x + 0.6, y + 0.4, t, fontsize=7.5)
    for (x, y, t, fs) in sp.get("texts", []): ax.text(x, y, t, fontsize=fs)
    x0, y0 = sp["view"][0], sp["view"][1]
    for k in range(4): ax.add_patch(Rectangle((x0 + 2 + k * 2.5, y0 + 2), 2.5, 0.5, fc="black" if k % 2 == 0 else "white", ec="black", lw=0.5))
    ax.text(x0 + 2, y0 + 3, "0", fontsize=6.5); ax.text(x0 + 12, y0 + 3, "10 m", fontsize=6.5)
    ax.annotate("N", (sp["view"][2] - 4, sp["view"][3] - 3), (sp["view"][2] - 4, sp["view"][3] - 8), ha="center",
                arrowprops=dict(arrowstyle="-|>", lw=1), fontsize=10)
    ax.set_xlim(sp["view"][0], sp["view"][2]); ax.set_ylim(sp["view"][1] - 18, sp["view"][3] + 12)
    for k, t in enumerate(sp["header"]):
        ax.text(sp["view"][0], sp["view"][3] + 9 - k * 2.4, t, fontsize=12 if k == 0 else 8, weight="bold" if k == 0 else "normal")
    for k, t in enumerate(sp["footer"]):
        ax.text(sp["view"][0], sp["view"][1] - 6 - k * 2.2, t, fontsize=7)
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "sp.pdf"; fig.savefig(out); plt.close(fig); return out.read_bytes()
SITEPLAN = siteplan_pdf() if CONTENT.get("siteplan") else None

HAS_PLAN = "rooms" in (CONTENT.get("plan") or {})
PLAN_INFO = check_plan(CONTENT["plan"], F) if HAS_PLAN else _NS()
PLAN = plan_pdf() if HAS_PLAN else roof_plan_pdf() if "houses" in (CONTENT.get("plan") or {}) else None

# ---------------------------------------------------------------------------------------------- 3a sketch photo and web form (S1)
def sketch_jpg():
    """Hand sketch of the site on squared paper, photographed with a phone (DL-172)."""
    import matplotlib; matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
    from PIL import Image, ImageFilter, ImageEnhance
    import numpy as np
    sk = CONTENT["sketch"]
    hand = "/usr/share/fonts/opentype/urw-base35/Z003-MediumItalic.otf"
    fp = font_manager.FontProperties(fname=hand) if Path(hand).exists() else None
    T = dict(fontproperties=fp, color="#1d2a6b")
    with plt.xkcd(scale=1.2, length=80, randomness=3):
        fig, ax = plt.subplots(figsize=(8.27, 5.8)); ax.set_xlim(0, 30); ax.set_ylim(0, 20); ax.axis("off")
        for g in np.arange(0, 30.1, 0.75): ax.axvline(g, color="#9fb8d6", lw=0.35, zorder=0).set_sketch_params(None)   # printed grid stays straight
        for g in np.arange(0, 20.1, 0.75): ax.axhline(g, color="#9fb8d6", lw=0.35, zorder=0).set_sketch_params(None)
        for (x, y, w, h, t) in sk["boxes"]:
            ax.add_patch(plt.Rectangle((x, y), w, h, fill=False, lw=2, ec="#1d2a6b"))
            ax.text(x + w / 2, y + h / 2, t, ha="center", va="center", fontsize=16, **T)
        for (x0, y0, x1, y1, t) in sk["lines"]:
            ax.plot([x0, x1], [y0, y1], lw=1.6, color="#1d2a6b")
            if t: ax.text((x0 + x1) / 2, (y0 + y1) / 2 + 0.5, t, ha="center", fontsize=14, **T)
        for (x, y, t, fs) in sk["texts"]:
            ax.text(x, y, t, fontsize=fs + 3, **T)
        for (x0, y0, x1, y1, t) in sk["arrows"]:
            ax.annotate(t, (x1, y1), (x0, y0), arrowprops=dict(arrowstyle="->", lw=1.3, color="#1d2a6b"), fontsize=13, **T)
        with tempfile.TemporaryDirectory() as tmp:
            png = Path(tmp) / "s.png"; fig.savefig(png, dpi=110, facecolor="#fbfaf4"); plt.close(fig)
            im = Image.open(png).convert("RGB")
            # paper texture: fine noise plus a light gradient from the lamp
            rng = np.random.default_rng(152)
            a = np.asarray(im).astype("float32")
            a += rng.normal(0, 5, a.shape[:2])[..., None]
            h, w = a.shape[:2]
            a *= (0.86 + 0.14 * np.linspace(1, 0, w))[None, :, None]
            im = Image.fromarray(a.clip(0, 255).astype("uint8"))
            # keystone: phone held slightly tilted over the table
            W, H = im.size; d = int(W * 0.04)
            canvas = Image.new("RGB", (W + 2 * d, H + 2 * d), (95, 85, 70)); canvas.paste(im, (d, d))
            src = [(d + d, d), (d, H + d - d // 3), (W + d, H + d), (W + d - d // 2, d + d // 3)]   # PIL order: UL, LL, LR, UR
            im = canvas.transform((W, H), Image.QUAD, data=[c for pt in src for c in pt], resample=Image.BICUBIC)
            im = im.rotate(-3, expand=True, fillcolor=(95, 85, 70), resample=Image.BICUBIC)
            im = ImageEnhance.Color(im).enhance(0.8); im = ImageEnhance.Brightness(im).enhance(0.95)
            im = im.filter(ImageFilter.GaussianBlur(0.6))
            buf = io.BytesIO(); im.save(buf, "JPEG", quality=78); return buf.getvalue()
SKETCH = sketch_jpg() if CONTENT.get("sketch") else None

# ---------------------------------------------------------------------------------------------- 3b meeting notes by the property manager (docx attachment)
def meeting_notes():
    mn = CONTENT["meeting_notes"]
    doc = new_doc(header=False)
    p = doc.add_paragraph(); r = p.add_run("Talkessel Immobilien-Treuhand AG"); r.bold = True
    doc.add_paragraph("Rebbergstrasse 25, 6430 Schwyz, Bewirtschaftung Stockwerkeigentum")
    doc.add_heading(mn["title"], level=1)
    table(doc, ["", ""], mn["meta"], widths=[4, 12])
    for title, items in mn["sections"]:
        doc.add_heading(title, level=2)
        for t in items: doc.add_paragraph(t, style="List Bullet")
    doc.add_heading("Pendenzen", level=2)
    table(doc, ["Was", "Wer", "Bis"], mn["todos"], widths=[9, 3.5, 3.5])
    doc.add_paragraph(mn["footer"]).runs[0].italic = True
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue(), " ".join(t for _, items in mn["sections"] for t in items)
NOTES, NOTES_TEXT = meeting_notes() if CONTENT.get("meeting_notes") else (None, "")

# ---------------------------------------------------------------------------------------------- 4 mails
ids = {}
for m in CONTENT["mails"]:
    quoted = []
    for q in m.get("quote", []):
        qm = next(x for x in CONTENT["mails"] if x["id"] == q)
        quoted.append({"frm": person(qm["from"]), "to": person(qm["to"]), "date": qm["date"], "subject": qm["subject"], "body": qm["body"]})
    for q in m.get("quote_inline", []):
        quoted.append({"frm": person(q["from"]), "to": person(q["to"]), "date": q["date"], "subject": q["subject"], "body": q["body"]})
    att = []
    if m.get("attach_plan"): att.append((CONTENT["plan"]["attachment_name"], PLAN, "application/pdf"))
    if m.get("attach_siteplan"): att.append((CONTENT["siteplan"]["attachment_name"], SITEPLAN, "application/pdf"))
    if m.get("attach_webform"):
        import json as _json
        att.append((CONTENT["webform"]["json_name"], _json.dumps(CONTENT["webform"]["data"], ensure_ascii=False, indent=2).encode("utf-8"), "application/json"))
    if m.get("attach_sketch"): att.append((CONTENT["sketch"]["file_name"], SKETCH, "image/jpeg"))
    if m.get("attach_notes"): att.append((CONTENT["meeting_notes"]["file_name"], NOTES, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
    if m.get("attach_offer"): att.append((offer_pdf.name, offer_pdf.read_bytes(), "application/pdf"))
    for a_ in m.get("attach_files", []):     # files rendered into the Case folder, attached as well
        fp = CASE_DIR / a_; att.append((fp.name, fp.read_bytes(), "application/pdf" if fp.suffix == ".pdf" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
    ids[m["id"]] = R.eml(CASE_DIR / m["file"], person(m["from"]), person(m["to"]), m["subject"], m["date"], m["body"],
                         cc=person(m["cc"]) if m.get("cc") else None, attachments=att, quoted=quoted, mailer=m.get("mailer"))

# ---------------------------------------------------------------------------------------------- 5 transcript and file note
if CONTENT.get("crm_print"):
    R.crm_print(CASE_DIR / CONTENT["crm_print"]["file"], CONTENT["crm_print"])
for fn in ([CONTENT["file_note"]] if CONTENT.get("file_note") else []) + CONTENT.get("file_notes", []):
    R.file_note(CASE_DIR / fn["file"], fn["fields"], fn["content"], fn["agreements"], fn["todos"])

# ---------------------------------------------------------------------------------------------- 6 costing sheet (estimator)
def costing():
    """Estimator sheet. Formulas are evaluated here as well and must reproduce the offered quantities (DL-162)."""
    a, b = F["terrace_main"]["width"], F["terrace_main"]["depth"]; d, c = F["terrace_wing"]["width"], F["terrace_wing"]["depth"]
    rows = [  # (label, formula or value, unit, codes that must equal the value)
        ("Hauptfläche", f"=ROUND({a}*{b},2)", "m2", []), ("Sitzplatz", f"=ROUND({d}*{c},2)", "m2", []),
        ("Fläche total", "=B2+B3", "m2", []), ("Fläche gerundet", "=ROUND(B4,0)", "m2", ["20.110.01", "50.100.04", "30.110.01"]),
        ("Abdichtung +10 % aufgerundet", "=ROUNDUP(B4*1.1,0)", "m2", ["30.120.01"]),
        ("Platten neu +10 % aufgerundet (Eventualpos.)", "=ROUNDUP(B5*1.1,0)", "m2", ["opt:50.100.01"]),
        (f"Dachrand (Süd {a + d:.2f} + West {b:.2f} + Ost {c:.2f})", F.edge_open, "m", []),
        (f"Anschluss Fassade ({a:.2f} + {c - b:.2f} + {d:.2f})", F.edge_facade, "m", []),
        ("Aufbordung und Demontage Spengler", "=ROUNDUP(B8+B9,0)", "m", ["30.130.01", "20.130.01"]),
        ("Dachrand neu +5 %", "=ROUNDUP(B8*1.05,0)", "m", ["60.100.01"]), ("Deckstreifen neu +5 %", "=ROUNDUP(B9*1.05,0)", "m", ["60.120.01"]),
        ("Seitenschutz, Geländer, Randabschluss", "=ROUNDUP(B8,0)", "m", ["10.160.01", "80.200.06", "50.140.01", "opt:80.200.07"]),
        ("Mulden 7 m3 (1 pro 60 m2)", "=ROUNDUP(B4/60,0)", "Stk", ["10.150.01"]),
        ("Ersatzplatten 10 %", "=ROUNDUP(B4*0.1,0)", "m2", ["50.150.01"]),
        ("Regie Reserve", "=12*125+12*115+800", "CHF", ["text:Reserve und Unvorhergesehenes"])]
    vals = {}
    def ev(expr):
        e = expr.lstrip("=")
        e = re.sub(r"B(\d+)", lambda m_: repr(vals[int(m_.group(1))]), e)
        e = re.sub(r"ROUNDUP\(", "_ru(", e); e = re.sub(r"ROUND\(", "_r(", e)
        return eval(e, {"_ru": lambda x, n: math.ceil(round(x, 9)), "_r": lambda x, n: round(x, n)})
    import math
    for i, (lab, f, u, codes) in enumerate(rows, 2):
        vals[i] = ev(f) if isinstance(f, str) else f
        for code in codes:
            if code.startswith("text:"):
                tgt = next(l["unit_price"] for s_ in sections for l in s_["lines"] if l["text"] == code[5:])
            else:
                opt = code.startswith("opt:"); cc = code[4:] if opt else code
                tgt = next(l["qty"] for s_ in sections for l in s_["lines"] if l.get("code") == cc and bool(l.get("optional")) == opt)
            if float(tgt) != float(vals[i]): raise SystemExit(f"costing sheet {lab}: {vals[i]} != offer {code} {tgt}")
    wb = Workbook(); ws = wb.active; ws.title = "Mengen"
    ws.append(["Was", "Menge", "EH"]); [setattr(x, "font", Font(bold=True)) for x in ws[1]]
    for lab, f, u, _ in rows: ws.append([lab, f, u])
    ws["A7"].comment = Comment(CONTENT["costing"]["notes"][1], "BUl")
    ws["A16"].comment = Comment(CONTENT["costing"]["notes"][2], "BUl")
    ws.column_dimensions["A"].width = 46
    k = wb.create_sheet("Kalkulation")
    k.append([f'Kalkulation {off["number"]}', "", "", "", "", "", f"Stand {R.de_date(date)} BUl/RBe"]); k["A1"].font = Font(bold=True)
    k.append(["Pos", "Art-Nr", "Text", "Menge", "EH", "Preis", "Betrag"]); [setattr(x, "font", Font(bold=True)) for x in k[2]]
    flat = [l for s_ in sections for l in s_["lines"]]
    for l, row in zip(expected["lines"], flat):
        k.append([l["pos"], l["code"], ("(EVENTUAL) " if l["optional"] else "") + row["text"], l["qty"], l["unit"], l["unit_price"], l["total"]])
    for lab, val in [("Brutto", expected["gross"]), ("Rabatt 2 %", -expected["discount"]), ("Skonto 2 %", -expected["skonto"]),
                     ("Netto", expected["net"]), ("MWST 8.1 %", expected["vat"]), ("Total", expected["total"])]:
        k.append(["", "", lab, "", "", "", val])
    for n in CONTENT["costing"]["notes"]:
        k.append(["", "", "Bem.: " + n])
    for col, wd in zip("ABCDEFG", [6, 11, 60, 8, 6, 10, 12]): k.column_dimensions[col].width = wd
    wb.save(CASE_DIR / CONTENT["costing"]["file"])
if CONTENT.get("costing"): costing()

# ---------------------------------------------------------------------------------------------- 7 past projects (frozen)
additions = {"customers": [], "offers": [], "projects": [], "leads": [], "hours": []}
def recipient_for(cid):
    c = CUST[cid]
    name = c["company"] or f'{c["salutation"]} {c["first"]} {c["last"]}'
    return [name, c["street"], f'{c["zip"]} {c["town"]}']

def build_past(pp_file):
    P = yaml.safe_load(open(pp_file, encoding="utf-8"))
    _r = recipient_for(P["customer"]); addr = f"{_r[1]}, {_r[2]}"
    P["object_lines"] = [x.replace("{addr}", addr) for x in P["object_lines"]]
    P["offer"]["heading"] = [x.replace("{addr}", addr) for x in P["offer"].get("heading", [])]
    pdir = ROOT / "company-world/past-projects" / P["folder"]
    for sub in ("03_Offerte", "05_Abrechnung"): (pdir / sub).mkdir(parents=True, exist_ok=True)
    o = P["offer"]; odate = str(o["date"]); year = int(odate[:4])
    CASE_OFFER_DISCOUNT[:] = [o.get("discount", 0), o.get("skonto", False)]
    PL_ = None
    if o["layout"] == "word_lump":                       # costing first, the lump sum is the only priced line
        PF = type("NS", (dict,), {"__getattr__": dict.get})(P["facts"])
        for k, expr in (P.get("derived") or {}).items():
            PF[k] = eval(expr, {"__builtins__": {}, "ceil": math.ceil, "floor": math.floor, "round": round, "sqrt": math.sqrt, "max": max, "min": min}, {"F": PF})
        PL_ = LO.cost(P["costing"], PF, W, year)
        o["sections"] = [{"title": "Pauschalpreis", "lines": [{"text": P["costing"]["title"], "qty": 1, "unit": "Pau", "unit_price": PL_["lump"]}]}]
    secs = o.get("sections") or [{"title": "Leistungen", "lines": o["lines"]}]
    res, psecs = priced_sections(secs, odate)
    rec = recipient_for(P["customer"])
    pm = R.STAFF[o["project_manager"]]
    if o["layout"] == "word_lump":
        r_ = res
        lo = dict(o, recipient=o.get("recipient") or rec, vat_rate=r_["vat_rate"], validity_text=o.get("validity_text", "3 Monate"),
                  payment_text=OPS["commercial_rules"]["payment_terms"][o["payment"]])
        LO.render(pdir / "03_Offerte" / f'{o["number"]}_Offerte.docx', lo,
                  [("A", P["costing"]["title"], r_["net"], r_["vat"], r_["total"], P["texts"]["variant_scope"])], None, odate, P["texts"])
        LO.sheet(pdir / "03_Offerte" / f'Kalkulation_{o["number"]}.xlsx', {"A": (P["costing"]["title"], PL_)}, f'Kalkulation {o["number"]}',
                 R.STAFF[o["project_manager"]].get("display_ini", o["project_manager"]) + " " + R.de_date(odate), PF)
    elif o["layout"] == "word_fee":
        vu = dt.date.fromisoformat(odate); vu = vu.replace(month=(vu.month + 2) % 12 + 1, year=vu.year + (vu.month + 3 > 12))
        c_ = CUST[P["customer"]]; rec_c = [rec[0], f'{c_["salutation"]} {c_["first"]} {c_["last"]}'] + rec[1:]
        fo = dict(o, recipient=o.get("recipient") or rec_c, valid_until=vu.isoformat(), texts=P["texts"])
        FO.render(pdir / "03_Offerte" / f'{o["number"]}_Honorarofferte.docx', fo, res, psecs, odate, fee_rates(year))
    elif o["layout"] == "2024":
        vu = dt.date.fromisoformat(odate); vu = vu.replace(month=(vu.month + 2) % 12 + 1, year=vu.year + (vu.month + 3 > 12))
        od = dict(number=o["number"], recipient=rec, object_lines=P["object_lines"], project_manager=o["project_manager"],
                  heading=o["heading"], client_provisions=o.get("client_provisions", []), description=o.get("description", []),
                  payment_text=OPS["commercial_rules"]["payment_terms"][o["payment"]], valid_until=vu.isoformat())
        R.metabau_offer(pdir / "03_Offerte" / f'{o["number"]}_Offerte.pdf', od, res, psecs, odate, layout="2024")
    elif o["layout"] == "2023":
        od = dict(number=o["number"], recipient=rec, object_lines=P["object_lines"],
                  project_manager=o["project_manager"], heading=o["heading"], client_provisions=o["client_provisions"],
                  payment_text=OPS["commercial_rules"]["payment_terms"][o["payment"]], valid_until=odate)
        R.metabau_offer(pdir / "03_Offerte" / f'{o["number"]}_Offerte.pdf', od, res, psecs, odate, layout="2023")
    else:   # 2019 Word letter, scanned
        doc = new_doc(title_line=True)
        for t in rec: doc.add_paragraph(t).paragraph_format.space_after = 0
        doc.add_paragraph(f"\nIbach, {R.de_date(odate, long=True)}")
        doc.add_heading(f'Offerte Nr. {o["number"]}', level=2)
        doc.add_paragraph("Objekt: " + ", ".join(P["object_lines"]))
        doc.add_paragraph(f'Sehr geehrter Herr {rec[0].split()[-1]}'); doc.add_paragraph(o["intro"])
        rows = [(i, l["text"], R.qty_fmt(l["qty"], 2), l["unit"], chf(l["unit_price"]), chf(l["total"])) for i, l in enumerate(res["lines"], 1)]
        table(doc, ["Pos.", "Beschreibung", "Menge", "EH", "Preis", "Betrag Fr."], rows, widths=[1, 8, 1.8, 1.2, 2, 2.4], font=8.5)
        table(doc, ["", "Fr."], [("Total exkl. MWST", chf(res["net"])), (f'MWST {res["vat_rate"]*100:.1f} %', chf(res["vat"])),
                                 ("Total inkl. MWST", chf(res["total"]))], widths=[10, 6])
        doc.add_paragraph("Gültigkeit: 30 Tage. Zahlung: 30 Tage netto. Garantie gemäss SIA-Norm 118.")
        doc.add_paragraph(f'Freundliche Grüsse\n{C["name"]}\n\n{pm["first"]} {pm["last"]}')
        with tempfile.TemporaryDirectory() as tmp:
            f = Path(tmp) / "o.docx"; doc.save(f); to_pdf(f)
            R.scanify(f.with_suffix(".pdf"), pdir / "03_Offerte" / f'Offerte_{o["number"]}_Scan.pdf', seed=year)
    # measured quantities and final invoice
    m = P["measured"]; mrows = []
    if m.get("lines"):                       # Advisory: invoice by actual hours and flat fees, prices of the invoice year
        iy = int(m.get("price_year") or str(m["invoice_date"])[:4])      # rates of the offer year if agreed (P-24-124)
        for ln in m["lines"]:
            if not ln.get("code"):                # pass through costs, e.g. lab invoice at net price
                mrows.append({"code": "", "text": ln["text"], "unit": ln["unit"], "qty": ln["qty"], "unit_price": D(str(PL_["lump"] if ln["unit_price"] == "LUMP" else ln["unit_price"]))}); continue
            it = W.items[ln["code"]]
            mrows.append({"code": ln["code"], "text": ln.get("text", it["text"]), "unit": it["unit"], "qty": ln["qty"], "unit_price": W.price(ln["code"], iy)})
        res_lines_for_ausmass = []
    for l in (res["lines"] if not m.get("lines") else []):
        q = m["qty"].get(l.get("code", ""), l["qty"])
        mrows.append({"code": l.get("code", ""), "text": l["text"], "unit": l["unit"], "qty": q, "unit_price": l["unit_price"]})
    for e in m.get("extra", []):
        up = e.get("unit_price") or W.rate(e["rate"], year)
        mrows.append({"text": e["text"], "unit": e["unit"], "qty": e["qty"], "unit_price": up})
    CASE_OFFER_DISCOUNT[:] = [o.get("discount", 0), o.get("skonto", False)]
    inv = compute_offer(mrows, str(m["invoice_date"]), discount=o.get("discount", 0), skonto=False, world=W)   # Skonto is a payment option, not deducted (DL-160)
    sk_note = ""
    if o.get("skonto"):
        sk = round_to(inv["gross"] * D("0.02"), "0.05"); sk_vat = round_to(sk * inv["vat_rate"], "0.05")
        sk_note = (f"Bei Zahlung innert 10 Tagen 2 % Skonto auf dem Bruttobetrag: CHF {chf(sk)} zuzüglich MWST CHF {chf(sk_vat)}, "
                   f"Zahlbetrag CHF {chf(inv['total'] - sk - sk_vat)}")
    wb = Workbook() if P.get("ausmass", True) else None
    if wb is None: pass
    else: ws = wb.active; ws.title = "Ausmass"
    if wb is not None:
        ws.append([f'Ausmass {P["nr"]}', "", "", "", f'Offerte {o["number"]}']); ws["A1"].font = Font(bold=True)
        ws.append(["Art-Nr", "Text", "EH", "Menge Offerte", "Menge Ausmass", "Differenz"])
        for l, mr in zip(res["lines"], mrows):
            ws.append([l.get("code", ""), l["text"], l["unit"], float(l["qty"]), float(mr["qty"]), float(mr["qty"]) - float(l["qty"])])
        for mr in mrows[len(res["lines"]):]:
            ws.append(["Regie", mr["text"], mr["unit"], 0, float(mr["qty"]), float(mr["qty"])])
        for col, wd in zip("ABCDEF", [11, 60, 6, 14, 14, 10]): ws.column_dimensions[col].width = wd
        wb.save(pdir / "05_Abrechnung" / f'Ausmass_{P["nr"]}.xlsx')
    invd = dict(basis=m.get("basis"), notes=m.get("notes", []), number=m["invoice_number"], project=P["nr"], recipient=rec, object=", ".join(P["object_lines"][:1] + P["object_lines"][2:]),
                offer_ref=f'{o["number"]} vom {R.de_date(odate)}', period=m["period"], skonto_note=sk_note, qr=year >= 2023)
    inv_pdf = pdir / "05_Abrechnung" / f'Schlussrechnung_{m["invoice_number"]}.pdf'
    if year >= 2021:
        R.metabau_invoice(inv_pdf, invd, inv, inv["lines"], str(m["invoice_date"]))
    else:   # KALKO era: Word letter, scanned
        doc = new_doc(title_line=True)
        for t in rec: doc.add_paragraph(t).paragraph_format.space_after = 0
        doc.add_paragraph(f'\nIbach, {R.de_date(m["invoice_date"], long=True)}')
        doc.add_heading(f'Schlussrechnung Nr. {m["invoice_number"]}', level=2)
        doc.add_paragraph(f'Objekt: {", ".join(P["object_lines"])}\nIhr Auftrag gemäss Offerte {o["number"]} vom {R.de_date(odate)}, Ausführung {m["period"]}')
        rows = [(i, l["text"], R.qty_fmt(l["qty"], 2), l["unit"], chf(l["unit_price"]), chf(l["total"])) for i, l in enumerate(inv["lines"], 1)]
        table(doc, ["Pos.", "Beschreibung (nach Ausmass)", "Menge", "EH", "Preis", "Betrag Fr."], rows, widths=[1, 8, 1.8, 1.2, 2, 2.4], font=8.5)
        trows = [("Total exkl. MWST", chf(inv["net"])), (f'MWST {inv["vat_rate"]*100:.1f} %', chf(inv["vat"]))]
        if inv["total"] != inv["net"] + inv["vat"]: trows.append(("Rundung", chf(inv["total"] - inv["net"] - inv["vat"])))
        trows.append(("Total inkl. MWST", chf(inv["total"])))
        table(doc, ["", "Fr."], trows, widths=[10, 6])
        doc.add_paragraph("Zahlbar innert 30 Tagen netto. Einzahlungsschein liegt bei.")
        doc.add_paragraph(f'Freundliche Grüsse\n{C["name"]}')
        with tempfile.TemporaryDirectory() as tmp:
            f = Path(tmp) / "r.docx"; doc.save(f); to_pdf(f)
            R.scanify(f.with_suffix(".pdf"), pdir / "05_Abrechnung" / f'Schlussrechnung_{m["invoice_number"]}_Scan.pdf', seed=year + 1)
    if P.get("summary"):
        doc = new_doc(title_line=True); doc.add_heading(f'Projektblatt {P["nr"]}', level=1)
        table(doc, ["", ""], [("Objekt", ", ".join(P["object_lines"])), ("Kunde", rec[0] + f' ({P["customer"]})'),
                              ("Offerte", f'{o["number"]} vom {R.de_date(odate)}, netto {"vor Skonto " if o.get("skonto") else ""}CHF {chf(res["gross"] - res["discount"])}'),
                              ("Schlussrechnung", f'{m["invoice_number"]} vom {R.de_date(m["invoice_date"])}, netto CHF {chf(inv["net"])}'),
                              ("Ausführung", m["period"]), ("Projektleitung", f'{pm["first"]} {pm["last"]}'),
                              ("Equipe", ", ".join(R.STAFF[i]["first"] + " " + R.STAFF[i]["last"] for i in P["summary"]["team"]))], widths=[4, 12])
        doc.add_heading("Bemerkungen", level=2)
        for n in P["summary"]["notes"]: doc.add_paragraph(n, style="List Bullet")
        f = pdir / f'Projektblatt_{P["nr"]}.docx'; doc.save(f); to_pdf(f)
    additions["offers"].append({"nr": o["number"], "date": odate, "cust": P["customer"], "obj": P.get("erp_object") or (P["object_lines"][2] + " " + rec[2].split(" ", 1)[1])[:40],
                                "line": P.get("line", "Projektgeschäft"), "type": o.get("type", "Einheitspreisofferte"), "net": float(res["gross"] - res["discount"]), "status": "A",
                                "pl": R.STAFF[o["project_manager"]].get("display_ini", o["project_manager"]), "proj": P["nr"], **({"adopt": P["adopt_base"]} if P.get("adopt_base") else {})})
    s0, e0 = [dt.datetime.strptime(x, "%d.%m.%Y").date().isoformat() for x in m["period"].split(" bis ")]
    additions["hours"] += [[P["nr"]] + list(h) for h in P.get("hours", [])]
    additions["projects"].append({"nr": P["nr"], "alt": P.get("alt", ""), "cust": P["customer"], "obj": additions["offers"][-1]["obj"],
                                  "town": rec[2].split(" ", 1)[1], "line": P.get("line", "Projektgeschäft"), "offer": o["number"], "offer_net": float(res["gross"] - res["discount"]),
                                  "final_net": float(inv["net"]), "start": s0, "end": e0, "pl": additions["offers"][-1]["pl"], "status": "abgeschlossen", **({"adopt": P["adopt_base"]} if P.get("adopt_base") else {})})
    return P["nr"], float(res["net"]), float(inv["net"])

past = [build_past(ROOT / f'generation/past_projects/{p["nr"]}.yaml') for p in CASE.get("past_projects", [])]

# ---------------------------------------------------------------------------------------------- 8 systems additions
add_yaml = yaml.safe_load(open(ROOT / f"generation/cases/{CID}/systems.yaml", encoding="utf-8"))
for k in ("customers", "leads"):
    additions[k] += add_yaml.get(k, [])
for o in add_yaml.get("offers", []):     # ERP offer amount: net after discount, before Skonto (DL-164)
    additions["offers"].append(dict(o, net=CE_NET if (CE_NET is not None and o["nr"] == CONTENT["cost_estimate"]["number"]) else round(expected["gross"] - expected["discount"], 2)))
sa_path = ROOT / "generation/world/systems_additions.yaml"
SA = yaml.safe_load(open(sa_path, encoding="utf-8")) if sa_path.exists() else {}
SA = SA or {}; SA[CID] = additions
sa_path.write_text("# Rows that Cases add to the CRM and ERP exports (DL-133). Written by build_case.py.\n" +
                   yaml.safe_dump(SA, allow_unicode=True, sort_keys=False, width=200), encoding="utf-8")
subprocess.run([sys.executable, str(ROOT / "generation/scripts/build_systems.py")], check=True)

# ---------------------------------------------------------------------------------------------- 9 validation (DL-150, DL-159)
import email as _email, csv as _csv
from email import policy as _policy
from email.utils import parsedate_to_datetime
def pdf_text(p):
    return subprocess.run(["pdftotext", "-layout", str(p), "-"], capture_output=True, text=True).stdout
def pdf_bytes_text(b):
    with tempfile.NamedTemporaryFile(suffix=".pdf") as t:
        t.write(b); t.flush(); return pdf_text(t.name)
def docx_bytes_text(b):
    from docx import Document as _Doc
    d = _Doc(io.BytesIO(b)); return "\n".join([p.text for p in d.paragraphs] + [c.text for t in d.tables for r in t.rows for c in r.cells])
errors = []
TEXTS = {}                                  # file -> full text incl. attachments
MAILS = []
for f in sorted(CASE_DIR.rglob("*")):
    if f.suffix == ".eml":
        m = _email.message_from_bytes(f.read_bytes(), policy=_policy.default)
        t = m.get_body(("plain",)).get_content()
        for part in m.iter_attachments():
            data = part.get_content(); fn = part.get_filename() or ""
            t += "\n" + (pdf_bytes_text(data) if fn.endswith(".pdf") else docx_bytes_text(data) if fn.endswith(".docx") else "")
        TEXTS[f] = t; MAILS.append((f, m))
    elif f.suffix == ".pdf": TEXTS[f] = pdf_text(f)
    elif f.suffix == ".docx": TEXTS[f] = docx_bytes_text(f.read_bytes())
    elif f.suffix in (".vtt", ".txt", ".json"): TEXTS[f] = f.read_text(encoding="utf-8")
ALL = "\n".join(TEXTS.values())
# lump sum costing sheet: Excel formulas recalculated by LibreOffice give the offered lump sums (DL-189)
if LUMP and CONTENT.get("costing_sheet"):
    import shutil as _sh, openpyxl as _ox
    with tempfile.TemporaryDirectory() as _t:
        _sh.copy(CASE_DIR / CONTENT["costing_sheet"]["file"], _t + "/k.xlsx")
        subprocess.run(["soffice", "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", _t + "/o", _t + "/k.xlsx"], capture_output=True)
        _wb = _ox.load_workbook(_t + "/o/k.xlsx", data_only=True)
        for _k in LUMP:
            _ws = _wb[f"Variante {_k}"]
            if float(_ws.cell(_ws.max_row, 6).value or 0) != float(LUMP[_k]["lump"]): errors.append(f"costing sheet {_k}: {_ws.cell(_ws.max_row, 6).value} != {LUMP[_k]['lump']}")
# offer arithmetic printed
ot = pdf_text(offer_pdf)
for key in ("gross", "net", "vat", "total"):
    if chf(expected[key]) not in ot: errors.append(f"offer PDF misses {key} {chf(expected[key])}")
if FEE_PRINTED is not None:
    for amt in FEE_PRINTED:
        if amt not in ot: errors.append(f"fee offer PDF misses amount {amt}")
else:
  for l in expected["lines"]:
    amt = chf(l["total"])
    if (f"({amt})" if l["optional"] else amt) not in ot: errors.append(f'offer PDF misses line {l["pos"]} amount {amt}')
if CONTENT.get("cost_estimate") and CE_MISSING: errors.append(f"cost estimate PDF misses {CE_MISSING}")
# flaw markers
for fl in CASE.get("flaws", []):
    _tpl = "\n".join(docx_bytes_text(f.read_bytes()) for f in (ROOT / "company-world/operations/templates").glob("*.docx"))
    if fl.get("marker") and fl["marker"] not in ALL + _tpl + open(ROOT / "company-world/systems/MetaBau_Export_Debitoren_2026-04-30.csv", encoding="cp1252").read():
        errors.append(f'flaw marker not found: {fl["id"]} "{fl["marker"]}"')
# language: no sharp s
for f, t in TEXTS.items():
    if "ß" in t: errors.append(f"ß in {f.name}")
# chronology and working days
offer_day = dt.date.fromisoformat(str(CASE["dates"]["offer"]))
staff_domains = "@fronalp-mgh.ch"
for f, m in MAILS:
    when = parsedate_to_datetime(m["Date"])
    if when.date() > offer_day: errors.append(f"mail after offer date: {f.name}")
    if staff_domains in m["From"] and "noreply@" not in m["From"] and (when.weekday() > 4 or not 6 <= when.hour <= 20):
        errors.append(f"staff mail outside working time: {f.name} {when}")
    if not f.name.startswith(when.date().isoformat()): errors.append(f"file name date differs from Date header: {f.name}")
    body = m.get_body(("plain",)).get_content()
    for qd in re.findall(r"Gesendet: (\d\d\.\d\d\.\d{4}) (\d\d:\d\d)", body):
        q = dt.datetime.strptime(" ".join(qd), "%d.%m.%Y %H:%M")
        if q >= when.replace(tzinfo=None): errors.append(f"quoted mail not earlier than {f.name}: {qd}")
vu = dt.date.fromisoformat(str(CASE["dates"]["valid_until"]))
if (vu.year - offer_day.year) * 12 + vu.month - offer_day.month != 3 or vu.day != offer_day.day: errors.append("validity is not 3 months")
# identifiers mentioned anywhere in the Case must exist
known = set()
for fn_ in ("customers.csv", "offers.csv", "projects.csv"):
    for r in _csv.DictReader(open(ROOT / "generation/world" / fn_, encoding="utf-8")):
        known |= {v for v in r.values() if v}
for c in add_yaml.get("customers", []): known |= {c.get("kc", ""), c.get("deb", "")}
for l in add_yaml.get("leads", []): known.add(l[0])
known |= {r[0] for r in _csv.reader(open(ROOT / "company-world/systems/KundenCockpit_Export_Kontakte_2026-04-30.csv", encoding="utf-8-sig"))}   # CRM contacts incl. duplicates
known.add(off["number"]); known |= {o["nr"] for o in add_yaml.get("offers", [])}
for pat in (r"\bL-20\d\d-\d{3}\b", r"\bKC-\d{5}\b", r"\bD-\d{5}\b", r"\bOF-\d\d-\d{4}\b", r"\bP-\d\d-\d{3}\b"):
    for ident in set(re.findall(pat, ALL)):
        if ident not in known: errors.append(f"unknown identifier in Case documents: {ident}")
# reserved numbers of all Cases sit at their date position in the base history (DL-186)
_SA = yaml.safe_load(open(ROOT / "generation/world/systems_additions.yaml", encoding="utf-8")) or {}
_offs = list(_csv.DictReader(open(ROOT / "generation/world/offers.csv", encoding="utf-8")))
_prjs = list(_csv.DictReader(open(ROOT / "generation/world/projects.csv", encoding="utf-8")))
for _cid, _add in _SA.items():
    for o in _add.get("offers", []):
        y = int(o["date"][:4])
        if y < 2021 or o["date"] > "2026-04-30": continue          # KALKO numbers; Case offers after the snapshot
        _same = [x for x in _offs if x["nr"].startswith(f"OF-{y % 100:02d}-") and x["nr"] != o["nr"]]
        lo_ = 1 + sum(1 for x in _same if x["date"] < o["date"]); hi_ = 1 + sum(1 for x in _same if x["date"] <= o["date"])   # same day: any order
        if not lo_ <= int(o["nr"][-4:]) <= hi_: errors.append(f"{_cid}: {o['nr']} ({o['date']}) not in date order, expected OF-{y % 100:02d}-{lo_:04d}")
    for pr in _add.get("projects", []):
        y = pr["nr"][2:4]
        _same = [x for x in _prjs if x["nr"].startswith(f"P-{y}-") and x["nr"] != pr["nr"]]
        lo_ = 1 + sum(1 for x in _same if x["start"] < pr["start"]); hi_ = 1 + sum(1 for x in _same if x["start"] <= pr["start"])
        if not lo_ <= int(pr["nr"][-3:]) <= hi_: errors.append(f"{_cid}: {pr['nr']} (start {pr['start']}) not in date order, expected P-{y}-{lo_:03d}")
# offers cited in the CRM print belong to the Case customer (DL-187)
if CONTENT.get("crm_print"):
    _deb = set(re.findall(r"D-\d{5}", str(CASE["customer"].get("ref", ""))))
    _byn = {o["nr"]: o for o in _offs}
    for act in CONTENT["crm_print"]["activities"]:
        for nr in re.findall(r"OF-\d\d-\d{4}", act[3]):
            if nr in _byn and _deb and _byn[nr]["cust"] not in _deb:
                errors.append(f"CRM print cites {nr} of {_byn[nr]['cust']}, not of the Case customer {sorted(_deb)}")
            if nr in _byn and _byn[nr]["obj"].split(" ")[0] not in act[3]:
                errors.append(f"CRM print text for {nr} does not match the ERP object '{_byn[nr]['obj']}'")
# every check and source
import math
QTY, OPT = {}, {}
for l in expected["lines"]:                  # summed per article (Advisory has several hour lines per category)
    if l["code"]:
        tgt = OPT if l["optional"] else QTY; tgt[l["code"]] = tgt.get(l["code"], 0) + l["qty"]
env = {"F": F, "PLAN": PLAN_INFO, "qty": lambda c: QTY.get(c, 0), "opt": lambda c: OPT.get(c), "ceil": math.ceil, "round": round}
for chk in CASE.get("checks", []):
    try:
        if not eval(chk, {"__builtins__": {}}, env): errors.append(f"check failed: {chk}")
    except Exception as e:
        errors.append(f"check error: {chk}: {e}")
for sec in off["sections"]:
    for ln in sec["lines"]:
        if not ln.get("src"): errors.append(f'line without source: {ln.get("code") or ln.get("text")}')
print(f"{CID} expected: gross {chf(expected['gross'])}, net {chf(expected['net'])}, VAT {chf(expected['vat'])}, total {chf(expected['total'])}, "
      f"owner approval {expected['needs_owner_approval']}")
for nr, onet, fnet in past: print(f"past project {nr}: offer net {chf(onet)}, final net {chf(fnet)}")
print("VALIDATION:", "ok" if not errors else "\n  " + "\n  ".join(errors))
