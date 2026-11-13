"""Build CRM and ERP exports of the Company world (base history, before any Case).

Canonical sources written to generation/world/:
  customers.csv, offers.csv, projects.csv
Exports as staff see them, in company-world/systems/:
  KundenCockpit_Export_Kontakte_2026-04-30.csv    CRM, UTF-8, comma
  KundenCockpit_Export_Leads_2026-04-30.csv       CRM, UTF-8, comma
  MetaBau_Export_Debitoren_2026-04-30.csv         ERP, Windows-1252, semicolon
  MetaBau_Export_Offerten_2024-2026.csv           ERP
  MetaBau_Export_Projekte_2018-2026.csv           ERP
  MetaBau_Export_Stunden_2025-2026.csv            ERP, hours per project, person and month
Cases add rows through generation/world/systems_additions.yaml (filled when a Case is generated).
Deterministic (seeded).
"""
import csv, datetime as dt, math, random, yaml
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SNAP = "2026-04-30"          # export snapshot: before the first contact of any Case (DL-158)
SNAP_D = dt.date.fromisoformat(SNAP)
def _easter(y):
    a, b, c = y % 19, y // 100, y % 100; d, e = b // 4, b % 4; f = (b + 8) // 25; g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30; i, k = c // 4, c % 4; l = (32 + 2 * e + 2 * i - h - k) % 7; m = (a + 11 * h + 22 * l) // 451
    return dt.date(y, (h + l - 7 * m + 114) // 31, (h + l - 7 * m + 114) % 31 + 1)
def _closed(d):
    """Public holidays in the canton of Schwyz and company holidays (calendar weeks 29 to 31), DL-187"""
    e = _easter(d.year); fixed = {(1, 1), (1, 6), (3, 19), (8, 1), (8, 15), (11, 1), (12, 8), (12, 25), (12, 26)}
    mov = {e - dt.timedelta(days=2), e + dt.timedelta(days=1), e + dt.timedelta(days=39), e + dt.timedelta(days=50), e + dt.timedelta(days=60)}
    return d.weekday() > 4 or (d.month, d.day) in fixed or d in mov or 29 <= d.isocalendar()[1] <= 31
def weekday(d):              # business documents fall on working days
    while _closed(d): d += dt.timedelta(days=1)
    return d
ORG = yaml.safe_load(open(ROOT / "generation/world/organisation.yaml", encoding="utf-8"))
rng = random.Random(1988)
OUT = ROOT / "company-world/systems"; OUT.mkdir(parents=True, exist_ok=True)
WORLD = ROOT / "generation/world"

TOWNS = [("Schwyz", "6430", 1), ("Ibach", "6438", 1), ("Brunnen", "6440", 1), ("Seewen", "6423", 1), ("Steinen", "6422", 1),
         ("Morschach", "6443", 1), ("Lauerz", "6424", 1), ("Gersau", "6442", 2), ("Goldau", "6410", 2), ("Arth", "6415", 2),
         ("Einsiedeln", "8840", 2), ("Küssnacht", "6403", 2), ("Rothenthurm", "6418", 2), ("Sattel", "6417", 2),
         ("Altdorf", "6460", 2), ("Zug", "6300", 2), ("Oberiberg", "8843", 2), ("Unteriberg", "8842", 2),
         ("Baar", "6340", 3), ("Luzern", "6004", 3), ("Pfäffikon SZ", "8808", 3), ("Wädenswil", "8820", 3), ("Zürich", "8003", 3)]
STREETS = ["Bergstrasse", "Dorfstrasse", "Seestrasse", "Sonnenhalde", "Mattenweg", "Rebbergstrasse", "Gartenweg",
           "Feldstrasse", "Kirchweg", "Lindenstrasse", "Bachstrasse", "Hofmatt", "Rosenweg", "Riedstrasse", "Bühlweg"]
FIRST_M = ["Thomas", "Daniel", "Markus", "Peter", "Martin", "Stefan", "Andreas", "Christian", "Beat", "Urs", "Roland", "Josef", "Lukas", "Michael"]
FIRST_F = ["Monika", "Sandra", "Barbara", "Claudia", "Andrea", "Nicole", "Ruth", "Esther", "Karin", "Priska", "Silvia", "Regula", "Franziska"]
LAST = ["Zwyssig", "Huwyler", "Inglin", "Ott", "Zehnder", "Wüthrich", "Schnüriger", "Abegg", "Kistler", "Marty", "Bucher",
        "Aschwanden", "Arnold", "Gisler", "Truttmann", "Ulrich", "Kälin", "Ochsner", "Schuler", "Beeler", "Reding", "Fuchs", "Weber", "Meier",
        "Suter", "Betschart", "Gwerder", "Bürgler", "Lüönd", "Heinzer", "Föhn", "Schelbert", "Imhof", "Camenzind", "Annen", "Niederöst",
        "Kennel", "Steiner", "Mettler", "Rickenbacher", "Fässler", "Baggenstos", "Schmidig", "Auf der Maur", "Holdener", "Kryenbühl"]

B2B = [  # name, category, town
    ("Arnold & Gisler Immobilien AG", "Verwaltung", "Brunnen"), ("Seeblick Liegenschaften GmbH", "Verwaltung", "Küssnacht"),
    ("Talkessel Immobilien-Treuhand AG", "Verwaltung", "Schwyz"), ("Rigi Verwaltungen AG", "Verwaltung", "Goldau"),
    ("Hirschgarten Immobilien AG", "Verwaltung", "Zürich"), ("Ägeri Haus & Wohnen GmbH", "Verwaltung", "Zug"),
    ("Wohnbaugenossenschaft Talkessel", "Genossenschaft", "Schwyz"), ("Baugenossenschaft Seewen-Hof", "Genossenschaft", "Seewen"),
    ("Wohngenossenschaft Rossberg", "Genossenschaft", "Goldau"), ("Genossenschaft Sonnmatt Einsiedeln", "Genossenschaft", "Einsiedeln"),
    ("Bauatelier Lauerz GmbH", "Architektur", "Lauerz"), ("Huwyler Partner Architekten AG", "Architektur", "Küssnacht"),
    ("Studio Riedmatt Architektur GmbH", "Architektur", "Zug"), ("Zwyssig Architekten AG", "Architektur", "Altdorf"),
    ("Urschweiz GU AG", "GU", "Brunnen"), ("Zentralbau Generalunternehmung AG", "GU", "Luzern"),
    ("Genossame Fronbach", "Öffentlich", "Steinen"), ("Gemeindewerke Seewen-Steinen", "Öffentlich", "Seewen"),
    ("Hotel Bürgenblick AG", "Gewerbe", "Gersau"), ("Garage Sattelegg AG", "Gewerbe", "Sattel"),
    ("Bergrestaurant Holzegg GmbH", "Gewerbe", "Schwyz"), ("Käserei Muota AG", "Gewerbe", "Schwyz"),
    ("Physiozentrum Mythen GmbH", "Gewerbe", "Schwyz"), ("Camping Lauerzersee GmbH", "Gewerbe", "Lauerz")]

def town(name=None):
    t = [x for x in TOWNS if x[0] == name][0] if name else rng.choices(TOWNS, weights=[6 if z == 1 else 3 if z == 2 else 1 for *_, z in TOWNS])[0]
    return t

def phone():
    return f"+41 {rng.choice(['41', '79', '78', '76'])} 000 {rng.randint(10, 99)} {rng.randint(10, 99)}"

def mail(local, dom):
    for a, b in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), (" ", ""), ("&", "")):
        local = local.replace(a, b); dom = dom.replace(a, b)
    return f"{local.lower()}@{dom.lower()}"

# ------------------------------------------------------------------ customers
customers = []
deb_seq = iter(range(10101, 19999, 7)); kc_seq = iter(range(20011, 29999, 3))
for name, cat, tn in B2B:
    t = town(tn); fn = rng.choice(FIRST_M + FIRST_F); ln = rng.choice(LAST)
    dom = name.split()[0].replace("&", "").lower() + "-" + t[0].split()[0].lower() + ".ch"
    customers.append(dict(type="Firma", company=name, category=cat, salutation="Herr" if fn in FIRST_M else "Frau",
                          first=fn, last=ln, street=f"{rng.choice(STREETS)} {rng.randint(1, 60)}", zip=t[1], town=t[0],
                          phone=phone(), email=mail(f"{fn}.{ln}", dom), since=rng.randint(1995, 2025)))
GEN_B2B = [("{} Immobilien GmbH", "Verwaltung"), ("{} Holzbau AG", "Gewerbe"), ("Garage {} AG", "Gewerbe"),
           ("{} Treuhand AG", "Verwaltung"), ("Bäckerei {}", "Gewerbe"), ("{} Architektur GmbH", "Architektur"),
           ("Erbengemeinschaft {}", "Privat"), ("Stockwerkeigentümergemeinschaft {}", "Verwaltung")]
for i in range(44):
    pat, cat = rng.choice(GEN_B2B); ln = rng.choice(LAST); t = town()
    name = pat.format(ln) if "Stockwerk" not in pat else pat.format(f"{rng.choice(STREETS)} {rng.randint(1, 40)}, {t[0]}")
    if any(c["company"] == name for c in customers): continue
    fn = rng.choice(FIRST_M + FIRST_F)
    customers.append(dict(type="Firma", company=name, category=cat, salutation="Herr" if fn in FIRST_M else "Frau",
                          first=fn, last=rng.choice(LAST), street=f"{rng.choice(STREETS)} {rng.randint(1, 60)}", zip=t[1], town=t[0],
                          phone=phone(), email=mail("info", ln.replace(" ", "") + "-" + t[0].split()[0] + ".ch"), since=rng.randint(1995, 2025)))
for i in range(650):
    female = rng.random() < 0.45; fn = rng.choice(FIRST_F if female else FIRST_M); ln = rng.choice(LAST); t = town()
    customers.append(dict(type="Person", company="", category="Privat", salutation="Frau" if female else "Herr",
                          first=fn, last=ln, street=f"{rng.choice(STREETS)} {rng.randint(1, 80)}", zip=t[1], town=t[0],
                          phone=phone(), email=mail(f"{fn[0]}.{ln}", rng.choice(["talnet-mail.ch", "postfach-sz.ch", "hausnet-mail.ch", "seenet-mail.ch"])) if rng.random() < 0.85 else "",
                          since=rng.randint(1992, 2026)))
RESERVED_ADDR = {"Seestrasse 18", "Hofmattweg 2", "Hofmattweg 4", "Wylenstrasse 12", "Industriestrasse 9"}       # Case object addresses must not appear in the base history (DL-163)
for c in customers:
    if c["street"] in RESERVED_ADDR: c["street"] = c["street"].replace("18", "81") if "18" in c["street"] else c["street"] + "a"
for c in customers:
    c["deb"] = f"D-{next(deb_seq)}" if c["since"] < 2026 or rng.random() < 0.5 else ""
    c["alt"] = f"K{rng.randint(1000, 9899)}" if c["since"] < 2021 and c["deb"] else ""
    c["kc"] = f"KC-{next(kc_seq)}" if (c["since"] >= 2012 or c["type"] == "Firma") else ""   # very old private customers never made it into the CRM
    c["discount"] = 0.02 if c["category"] in ("Verwaltung", "Genossenschaft", "GU") and rng.random() < 0.6 else 0.0
    c["id"] = c["deb"] or c["kc"]

# ------------------------------------------------------------------ offers and projects
LINES = {"Beratung": (0.15, 2500, 16000, ["Zustandsaufnahme Flachdach", "Sanierungskonzept Fassade", "Geländerprüfung",
                                         "Machbarkeitsstudie Dachaufbau", "Variantenstudie Terrasse"]),
         "Projektgeschäft": (0.55, 6000, 180000, ["Terrassensanierung", "Flachdachsanierung", "Spenglerarbeiten Steildach",
                                                 "Aussentreppe Stahl", "Balkongeländer", "Fassadenbekleidung", "Attikaabschluss",
                                                 "Treppengeländer innen", "Dachrandsanierung"]),
         "Systemprodukte": (0.30, 4500, 32000, ["Velounterstand", "Carport Einzel", "Carport Doppel", "Vordach Hauseingang"])}
TYPES = {"Beratung": ["Honorarofferte"], "Systemprodukte": ["Konfigurationsofferte", "Konfigurationsofferte", "Richtofferte"],
         "Projektgeschäft": ["Einheitspreisofferte", "Einheitspreisofferte", "Pauschalofferte", "Eingabe LV", "Richtofferte"]}
WIN = {"Beratung": 0.55, "Projektgeschäft": 0.35, "Systemprodukte": 0.6}
PL = {"Beratung": ["CSc", "NSu"], "Projektgeschäft": ["RBe", "RBe", "SKae", "DIn", "PKo"], "Systemprodukte": ["DIn"]}

def disp(ini):
    s = [x for x in ORG["staff"] if x["ini"] == ini][0]; return s.get("display_ini", ini)

offers, projects = [], []
proj_seq = {}
for year in range(2018, 2027):
    n = {2018: 255, 2019: 262, 2020: 248, 2021: 281, 2022: 304, 2023: 318, 2024: 331, 2025: 342, 2026: 110}[year]
    last_month = 4 if year == 2026 else 12
    for k in range(1, n + 1):
        line = rng.choices(list(LINES), weights=[v[0] for v in LINES.values()])[0]
        share, lo, hi, objs = LINES[line]
        typ = rng.choice(TYPES[line])
        if typ == "Eingabe LV": lo, hi = 40000, 350000
        amt = round(rng.lognormvariate(0, 0.6) * (lo + hi) / 4 / 5) * 5
        amt = max(lo, min(hi, amt))
        cust = rng.choice([c for c in customers if (c["category"] == "Architektur") == (typ == "Eingabe LV") and c["since"] <= year])
        pl = rng.choice(PL[line])
        if pl == "PKo" and year > 2023: pl = "DIn"
        if pl == "DIn" and line == "Projektgeschäft" and year < 2023: pl = "PKo"
        date = weekday(dt.date(year, rng.randint(1, last_month), rng.randint(1, 28)))
        status = "A" if rng.random() < WIN[line] else rng.choice(["V", "V", "V", "Z"])
        if year == 2026 and date.month >= 3: status = rng.choice(["O", "O", "A", "V"])
        obj = f"{rng.choice(objs)} {cust['town']}"
        if line == "Projektgeschäft":
            o0 = obj.split(" ")[0]
            if o0 in ("Terrassensanierung", "Flachdachsanierung", "Spenglerarbeiten", "Dachrandsanierung", "Attikaabschluss"): pl = "RBe"
            elif o0 in ("Aussentreppe", "Balkongeländer", "Treppengeländer"): pl = "SKae"
            elif o0 == "Fassadenbekleidung": pl = "PKo" if year <= 2023 else "DIn"
        of = dict(nr=f"OF-{year % 100:02d}-{k:04d}" if year >= 2021 else f"{year}-{k:03d}", date=date.isoformat(), year=year,
                  cust=cust["id"], obj=obj, line=line, type=typ, net=amt, status=status, pl=disp(pl), proj="")
        if status == "A":
            proj_seq[year] = proj_seq.get(year, 0) + 1
            pnr = f"P-{year % 100:02d}-{proj_seq[year]:03d}"
            of["proj"] = pnr
            if typ in ("Einheitspreisofferte", "Eingabe LV"): f = rng.uniform(0.92, 1.18)
            elif typ == "Richtofferte": f = rng.uniform(0.85, 1.30)
            elif typ == "Honorarofferte":            # billed by effort under a cost ceiling; overruns only with written consent (DL-187)
                u = rng.uniform(0, 1); f = 0.80 + 0.20 * u / 0.85 if u < 0.85 else 1.02 + 0.08 * (u - 0.85) / 0.15
            else: f = 1.0 if rng.random() < 0.7 else rng.uniform(1.01, 1.08)
            start = weekday(date + dt.timedelta(days=rng.randint(20, 120))); end = weekday(start + dt.timedelta(days=rng.randint(3, 45)))
            need_wd = amt * (0.8 if line == "Beratung" else 0.45) / 130 / (6 * (2 if line == "Beratung" else 4))   # people x 6 h per day
            if (end - start).days < need_wd * 7 / 5: end = weekday(start + dt.timedelta(days=math.ceil(need_wd * 7 / 5)))
            done = end < SNAP_D
            projects.append(dict(nr=pnr, alt=f"{year % 100}{proj_seq[year]:02d}" if year < 2021 else "", cust=cust["id"], obj=obj,
                                 town=cust["town"], line=line, offer=of["nr"], offer_net=amt,
                                 final_net=round(amt * f / 5) * 5 if done else "", start=start.isoformat(),
                                 end=end.isoformat() if done else "", pl=of["pl"], status="abgeschlossen" if done else ("in Ausführung" if start <= SNAP_D else "geplant")))
        offers.append(of)

# ------------------------------------------------------------------ Case additions (DL-133) and date ordered numbering
ADD = yaml.safe_load(open(WORLD / "systems_additions.yaml", encoding="utf-8")) if (WORLD / "systems_additions.yaml").exists() else {}
ADD_BY_CASE = ADD or {}
ADD = {"customers": [], "offers": [], "projects": [], "leads": [], "hours": []}
for case_add in ADD_BY_CASE.values():
    for k in ADD:
        ADD[k] += case_add.get(k, [])
reserved_of = {o["nr"] for o in ADD.get("offers", [])}
reserved_p = {p["nr"] for p in ADD.get("projects", [])}
# a Case offer or project before the snapshot replaces the nearest base record of the same year, so that adding
# a Case never shifts the numbers of earlier Cases (DL-190)
_protect = {o["cust"] for o in ADD["offers"]} | {p["cust"] for p in ADD["projects"]}
def _days(a, b): return abs((dt.date.fromisoformat(a) - dt.date.fromisoformat(b)).days)
GHOSTS = []                 # removed base projects: their hour draws are still taken, so the random stream stays stable (DL-196)
for rp in [p for p in ADD["projects"] if p.get("adopt")]:   # a Case past project takes over a base record of the same customer (DL-196)
    gone = next(p for p in projects if p["cust"] == rp["cust"] and p["obj"] == rp["adopt"])
    projects.remove(gone); GHOSTS.append(gone)
    offers.remove(next(o for o in offers if o["proj"] == gone["nr"]))
for ro in sorted([o for o in ADD["offers"] if o["date"] <= SNAP and not o.get("adopt")], key=lambda o: o["date"]):
    cand = [o for o in offers if o["year"] == int(ro["date"][:4]) and o["status"] != "A" and o["cust"] not in _protect and o["date"] > ro["date"]]
    if cand: offers.remove(min(cand, key=lambda o: (_days(o["date"], ro["date"]), o["date"])))
for rp in sorted([p for p in ADD["projects"] if p["start"] <= SNAP and not p.get("adopt")], key=lambda p: p["start"]):
    cand = [p for p in projects if p["nr"][2:4] == rp["nr"][2:4] and p["cust"] not in _protect and rp["start"] < p["start"] <= SNAP]
    if cand:
        gone = min(cand, key=lambda p: (_days(p["start"], rp["start"]), p["start"])); projects.remove(gone); GHOSTS.append(gone)
        go = next(o for o in offers if o["proj"] == gone["nr"]); go["status"], go["proj"] = "V", ""   # its offer stays, as lost
old_to_new = {}
for year in range(2018, 2027):
    yo = sorted([o for o in offers if o["year"] == year], key=lambda o: o["date"])
    k = 0
    for o in yo:
        while True:
            k += 1
            nr = f"OF-{year % 100:02d}-{k:04d}" if year >= 2021 else f"{year}-{k:03d}"
            if nr not in reserved_of: break
        old_to_new[o["nr"]] = nr; o["nr"] = nr
p_old_to_new = {}
for year in range(2018, 2027):
    yp = sorted([p for p in projects if p["nr"][2:4] == f"{year % 100:02d}"], key=lambda p: p["start"])
    k = 0
    for p in yp:
        while True:
            k += 1
            nr = f"P-{year % 100:02d}-{k:03d}"
            if nr not in reserved_p: break
        p_old_to_new[p["nr"]] = nr; p["nr"] = nr
        if p["alt"]: p["alt"] = f"{year % 100}{k:02d}"
for p in projects:
    p["offer"] = old_to_new.get(p["offer"], p["offer"])
for o in offers:
    o["proj"] = p_old_to_new.get(o["proj"], o["proj"])
def _crm_iso(c):
    d, m, y = c.get("crm_created", "01.01.1900").split("."); return f"{y}-{m}-{d}"
for c in ADD.get("customers", []):
    if _crm_iso(c) <= SNAP: customers.append(dict(c, _case=True))
for o in ADD.get("offers", []):
    if o["date"] <= SNAP: offers.append(dict(o, year=int(o["date"][:4]), _case=True))
for p in ADD.get("projects", []):
    if p["start"] <= SNAP: projects.append(dict(p, _case=True))
ADD["leads"] = [l for l in ADD.get("leads", []) if l[7] <= SNAP]
offers.sort(key=lambda o: o["date"]); projects.sort(key=lambda p: p["start"])
reserved_leads = {l[0] for l in ADD.get("leads", [])}

# ------------------------------------------------------------------ canonical sources
def write(path, rows, enc="utf-8", delim=",", header=None):
    with open(path, "w", newline="", encoding=enc) as f:
        w = csv.writer(f, delimiter=delim)
        header = header or list(rows[0].keys())
        w.writerow(header)
        for r in rows: w.writerow([r.get(k, "") for k in header] if isinstance(r, dict) else r)
strip = lambda rows: [{k: v for k, v in r.items() if k != "_case"} for r in rows]
write(WORLD / "customers.csv", strip(customers)); write(WORLD / "offers.csv", strip(offers)); write(WORLD / "projects.csv", strip(projects))

# ------------------------------------------------------------------ CRM exports (with duplicates)
crm_rows = []
for c in customers:
    if not c["kc"] or c.get("_case"): continue
    crm_rows.append([c["kc"], c["type"], c["company"], c["salutation"], c["first"], c["last"], c["street"], c["zip"], c["town"],
                     c["phone"], c["email"], c["deb"] if rng.random() < 0.8 else "", c["category"],
                     f'{rng.randint(1, 28):02d}.{rng.randint(1, 12):02d}.{max(2022, c["since"])}', ""])
dups = [("Arnold & Gisler Immobilien AG", "Arnold+Gisler Immobilien A.G."), ("Wohnbaugenossenschaft Talkessel", "WBG Talkessel"),
        ("Huwyler Partner Architekten AG", "Huwyler Partner Architekten"), ("Rigi Verwaltungen AG", "Rigi-Verwaltungen")]
for orig, alias in dups:
    r = next(x for x in crm_rows if x[2] == orig)
    crm_rows.append([f"KC-{rng.randint(30000, 39999)}", "Firma", alias, r[3], r[4], r[5], r[6], r[7], r[8], phone(), "", "",
                     r[12], f"{rng.randint(1, 28):02d}.0{rng.randint(1, 9)}.2022", "Import 2022"])
for c in customers:   # Case customers: all fields given, no randomness (keeps base stable)
    if c.get("_case") and c.get("kc"):
        crm_rows.append([c["kc"], c["type"], c["company"], c["salutation"], c["first"], c["last"], c["street"], c["zip"],
                         c["town"], c["phone"], c["email"], c.get("crm_deb", c["deb"]), c["category"], c["crm_created"], c.get("crm_note", "")])
write(OUT / "KundenCockpit_Export_Kontakte_2026-04-30.csv", crm_rows, enc="utf-8-sig",
      header=["KC_ID", "Typ", "Firma", "Anrede", "Vorname", "Nachname", "Strasse", "PLZ", "Ort", "Telefon", "E-Mail",
              "ERP_Debitor", "Kategorie", "Erstellt", "Notiz"])

NOTES = ["fup nxt wk", "sent offer, wait feedback", "Kunde will Termin vor Ostern", "callback CSc", "budget tight",
         "Besichtigung done", "lost to competitor (price)", "", "", "Plan fehlt noch", "follow up tel", "Offerte nachgefasst, kein Feedback"]
leads = []
lead_seq = {}
def next_lead(year):
    while True:
        lead_seq[year] = lead_seq.get(year, 0) + 1
        lid = f"L-{year}-{lead_seq[year]:03d}"
        if lid not in reserved_leads: return lid
case_offers = {o["nr"] for o in ADD.get("offers", [])}
for of in offers:
    if of["year"] < 2022: continue
    r_upd, r_val, r_note = rng.random(), rng.uniform(0.7, 1.3), rng.choice(NOTES)   # drawn for Case offers too, so the stream stays stable (DL-193)
    lid = next_lead(of["year"])
    kc = next((c["kc"] for c in customers if c["id"] == of["cust"]), "")
    if of["nr"] in case_offers:                # past project offer of a Case: won, value from the offer, no randomness (DL-141)
        leads.append([lid, kc, of["obj"], of["line"], "gewonnen", round(of["net"] / 1000) * 1000, of["pl"], of["date"], ""]); continue
    st = {"A": "gewonnen", "V": "verloren", "Z": "verloren", "O": "offeriert"}[of["status"]]
    if r_upd < 0.12: st = "offeriert"          # status not updated in CRM
    leads.append([lid, kc, of["obj"], of["line"], st, round(of["net"] * r_val / 1000) * 1000, of["pl"], of["date"], r_note])
for k in range(37):   # leads that never became offers
    leads.append([next_lead(2026), rng.choice([c["kc"] for c in customers if c["kc"]]), "Anfrage " + rng.choice(["Carport", "Vordach", "Geländer", "Dach"]),
                  rng.choice(list(LINES)), "neu", "", rng.choice(["EFö", "DIn", "RBe"]), weekday(dt.date(2026, rng.randint(3, 4), rng.randint(1, 28))).isoformat(),
                  rng.choice(["", "Rückruf ausstehend", "Unterlagen angefordert"])])
leads += [list(l) for l in ADD.get("leads", [])]
leads.sort(key=lambda l: (l[7], l[0]))
for l in leads:
    if len(l[7]) == 10 and l[7][4] == "-": l[7] = dt.date.fromisoformat(l[7]).strftime("%d.%m.%Y")
write(OUT / "KundenCockpit_Export_Leads_2026-04-30.csv", leads, enc="utf-8-sig",
      header=["LEAD_ID", "KC_ID", "Titel", "Bereich", "Status", "Wert_CHF", "Verantwortlich", "Erstellt", "Notiz"])

# ------------------------------------------------------------------ ERP exports
def chf(x): return "" if x == "" else f"{x:.2f}"
deb = [[c["deb"], c["alt"], c["company"] or f'{c["last"]} {c["first"]}', f'{c["salutation"]} {c["first"]} {c["last"]}' if c["company"] else "",
        c["street"], c["zip"], c["town"], c["phone"], c["email"], "30T NETTO", f'{c["discount"]*100:.1f}',
        "J" if c["company"] == "Arnold & Gisler Immobilien AG" else "N"] for c in customers if c["deb"]]
write(OUT / "MetaBau_Export_Debitoren_2026-04-30.csv", deb, enc="cp1252", delim=";",
      header=["DEB_NR", "ALT_NR", "NAME_1", "NAME_2", "STRASSE", "PLZ", "ORT", "TEL", "EMAIL", "ZAHLKOND", "RABATT_PROZ", "SKONTO_SPERRE"])
write(OUT / "MetaBau_Export_Offerten_2024-2026.csv",
      [[o["nr"], dt.date.fromisoformat(o["date"]).strftime("%d.%m.%Y"), o["cust"], o["obj"][:40], o["line"][:4].upper(), o["type"],
        chf(o["net"]), o["status"], o["pl"].upper(), o["proj"]] for o in offers if o["year"] >= 2024],
      enc="cp1252", delim=";", header=["OF_NR", "DATUM", "DEB_NR", "OBJEKT", "BER", "OFF_ART", "BETRAG_NETTO", "STAT", "PL", "PROJ_NR"])
write(OUT / "MetaBau_Export_Projekte_2018-2026.csv",
      [[p["nr"], p["alt"], p["cust"], p["obj"][:40], p["line"][:4].upper(), p["offer"], chf(p["offer_net"]), chf(p["final_net"]),
        dt.date.fromisoformat(p["start"]).strftime("%d.%m.%Y"), dt.date.fromisoformat(p["end"]).strftime("%d.%m.%Y") if p["end"] else "",
        p["pl"].upper(), p["status"]] for p in projects],
      enc="cp1252", delim=";", header=["PROJ_NR", "ALT_PROJ", "DEB_NR", "OBJEKT", "BER", "OF_NR", "OFF_NETTO", "SR_NETTO", "START", "ENDE", "PL", "STATUS"])

# hours per project, person and month for projects started 2025 or later
staff_by_cat = {}
for s in ORG["staff"]:
    if s.get("until") or s.get("board_only"): continue
    staff_by_cat.setdefault(s["rate_category"], []).append(s.get("display_ini", s["ini"]).upper())
RATE = {"fitter_a": 125, "fitter_b": 115, "workshop": 118, "apprentice": 65, "project_manager": 145, "advisory_owner": 185, "advisory_pm": 155, "advisory_technician": 125}
staff_by_cat["advisory_pm"] = ["RBE", "SKÄ"]           # project managers bill Advisory hours as B2 (cost estimates, railings)
MIX = {"Projektgeschäft": {"fitter_a": .40, "fitter_b": .25, "workshop": .15, "apprentice": .05, "project_manager": .15},
       "Systemprodukte": {"workshop": .45, "fitter_a": .25, "fitter_b": .20, "apprentice": .05, "project_manager": .05},
       "Beratung": {"advisory_owner": .15, "advisory_pm": .10, "advisory_technician": .75}}
hours = []
for p in sorted(projects + [dict(g, _ghost=True) for g in GHOSTS], key=lambda p: p["start"]):
    if p.get("_case") or (p["start"] < "2025-01-01" and (not p["end"] or p["end"] < "2025-01-01")): continue   # hours booked from 2025
    if p["start"] > SNAP: continue
    labour = (p["final_net"] or p["offer_net"]) * (0.45 if p["line"] != "Beratung" else 0.8)   # rest: flat fees, travel, material
    s0 = dt.date.fromisoformat(p["start"]); e0 = min(dt.date.fromisoformat(p["end"]) if p["end"] else SNAP_D, SNAP_D)
    months = sorted({(s0 + dt.timedelta(days=d)).strftime("%Y-%m") for d in range(0, (e0 - s0).days + 1, 10)}) or [s0.strftime("%Y-%m")]
    mix = MIX[p["line"]]; b2 = None
    if p["line"] == "Beratung":                  # B2 by trade: SKä for railings, RBe for roofs, terraces and facades
        o0 = p["obj"].split(" ")[0]
        if o0 == "Zustandsaufnahme": mix = {"advisory_owner": .15, "advisory_technician": .85}
        elif o0 == "Geländerprüfung": mix = {"advisory_owner": .10, "advisory_pm": .30, "advisory_technician": .60}; b2 = ["SKÄ"]
        else: mix = {"advisory_owner": .12, "advisory_pm": .13, "advisory_technician": .75}; b2 = ["RBE"]
    wd = max(1, sum(1 for k in range((e0 - s0).days + 1) if not _closed(s0 + dt.timedelta(days=k))))
    for cat, share in mix.items():
        h_total = labour * share / RATE[cat]
        pool = b2 if (cat == "advisory_pm" and b2) else staff_by_cat[cat]
        people = rng.sample(pool, k=min(len(pool), rng.randint(1, 3)))
        hp = min(h_total / len(people), wd * 7.0)              # nobody logs more than a working day per day (DL-187)
        for person in people:
            for m in months:
                h = round(hp / len(months) * rng.uniform(0.85, 1.05) * 2) / 2
                if h > 0 and m >= "2025-01" and not p.get("_ghost"): hours.append([p["nr"], person, {"fitter_a": "MA", "fitter_b": "MB", "workshop": "WS", "apprentice": "LE",
                                        "project_manager": "PL", "advisory_owner": "B1", "advisory_pm": "B2", "advisory_technician": "B3"}[cat], m, f"{h:.1f}"])
# capacity per person and month: about 150 h at 100 % (DL-187)
FTE = {s.get("display_ini", s["ini"]).upper(): s.get("fte", 100) for s in ORG["staff"]}
load = {}
for h in hours: load[(h[1], h[3])] = load.get((h[1], h[3]), 0) + float(h[4])
for h in hours:
    cap = 150 * FTE.get(h[1], 100) / 100
    if load[(h[1], h[3])] > cap: h[4] = f"{max(0.5, round(float(h[4]) * cap / load[(h[1], h[3])] * 2) / 2):.1f}"
hours += [h[:4] + [f"{float(h[4]):.1f}"] for h in ADD["hours"] if h[3] <= SNAP[:7]]   # past project hours of Cases
write(OUT / "MetaBau_Export_Stunden_2025-2026.csv", hours, enc="cp1252", delim=";", header=["PROJ_NR", "PERS_KZ", "KAT", "MONAT", "STD"])

print(f"{len(customers)} customers, {len(crm_rows)} CRM contacts (incl. {len(dups)} duplicates), {len(leads)} leads, "
      f"{len(offers)} offers, {len(projects)} projects, {len(hours)} hour rows")
