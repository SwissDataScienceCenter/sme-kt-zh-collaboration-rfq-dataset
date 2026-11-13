"""Build the Item master (Artikelstamm) of the Company world.

Writes
  generation/world/item_master.csv                      canonical source (UTF-8), used by pricing.py
  company-world/systems/MetaBau_Export_Artikelstamm_2026-04-30.csv   ERP export as staff see it
    (Windows-1252, semicolon separated, Swiss number format, cryptic headers, legacy numbers, blocked items)
Deterministic: same input, same output.
"""
import csv, random, yaml
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
OPS = yaml.safe_load(open(ROOT / "generation/world/operations.yaml", encoding="utf-8"))
CAT = yaml.safe_load(open(ROOT / "generation/world/catalogue.yaml", encoding="utf-8"))
rng = random.Random(1962)

GROUPS = {  # code: (name, index family, kind default)
    "10": ("Baustelle und Geräte", "other"), "20": ("Abbruch und Entsorgung", "other"),
    "30": ("Abdichtung", "bitumen_insulation"), "40": ("Wärmedämmung", "bitumen_insulation"),
    "50": ("Beläge und Nutzschichten", "other"), "60": ("Spenglerei", "sheet_metal"),
    "70": ("Stahl und Halbzeug", "steel"), "80": ("Geländer und Treppenkomponenten", "steel"),
    "85": ("Oberflächen", "other"), "90": ("Befestigung und Kleinmaterial", "other"),
    "95": ("Fassade", "other"), "97": ("Systemprodukte", "system_products"), "98": ("Beratungsleistungen", "labour"), "99": ("Regie und Leistungen", "labour"),
}
items = []
def add(code, text, unit, vk, kind="L", ek=None, supplier="", kg_m=None, text2=""):
    items.append(dict(code=code, group=code[:2], kind=kind, text=text, text2=text2, unit=unit,
                      vk=round(vk, 2), ek=None if ek is None else round(ek, 2), supplier=supplier, kg_m=kg_m))

# 10 site and equipment
for c, t, u, p, s in [
    ("10.100.01", "Baustelleninstallation klein, Zu- und Wegtransport Werkzeug und Geräte", "Pau", 450, ""),
    ("10.100.02", "Baustelleninstallation mittel", "Pau", 950, ""),
    ("10.100.03", "Baustelleninstallation gross inkl. Container", "Pau", 2200, ""),
    ("10.110.01", "Baustellen-WC, Miete pro Monat", "Mt", 445, "S10"),
    ("10.120.01", "Mobilkran inkl. An- und Rückfahrt, Auf- und Abbau", "Std", 340, "S10"),
    ("10.120.02", "Bewilligung Benützung öffentlicher Grund (Kran, Mulde)", "Pau", 180, "S10"),
    ("10.130.01", "Leiterlift, Miete pro Tag", "Tag", 210, "S10"),
    ("10.130.02", "Treppenturm, Miete pro Woche", "Wo", 390, "S10"),
    ("10.140.01", "Fassadengerüst, Weiterverrechnung Subunternehmer", "m2", 24, ""),
    ("10.150.01", "Mulde 7 m3 inkl. Entsorgung Bauschutt", "Stk", 890, "S10"),
    ("10.150.02", "Kranmulde 1 m3 für Sperrgut inkl. Entsorgung", "Stk", 435, "S10"),
    ("10.160.01", "Seitenschutz temporär, Absturzsicherung", "m", 38, ""),
    ("10.180.01", "Hubarbeitsbühne, Miete pro Tag", "Tag", 420, "S10")]:
    add(c, t, u, p, supplier=s)
for z in OPS["transport_zones"]:
    add(f"10.170.0{z['zone']}", f"Transportpauschale Zone {z['zone']}", "Pau", z["flat"])

# 20 demolition
for c, t, u, p in [
    ("20.100.01", "Überzug abspitzen, ohne Armierung, Dicke bis 70 mm", "m2", 66),
    ("20.100.02", "Überzug abspitzen, ohne Armierung, Dicke bis 120 mm", "m2", 89),
    ("20.110.01", "Plattenbelag inkl. Splitt aufnehmen, Platten reinigen und palettieren", "m2", 35),
    ("20.110.02", "Plattenbelag inkl. Splitt aufnehmen und entsorgen", "m2", 42),
    ("20.120.01", "Flachdachaufbau rückbauen: Bitumenbahnen, Dämmung, Dampfsperre", "m2", 49),
    ("20.130.01", "Spenglerarbeiten demontieren und entsorgen", "m", 28.5),
    ("20.140.01", "Geländer demontieren und entsorgen", "m", 42),
    ("20.150.01", "Kies abräumen und entsorgen", "m2", 18),
    ("20.160.01", "Bestehende Stahlbauteile trennen und entsorgen", "kg", 1.4)]:
    add(c, t, u, p)

# 30 waterproofing
for c, t, u, p in [
    ("30.100.01", "Untergrund reinigen und grundieren", "m2", 5.25),
    ("30.110.01", "Dampfsperre Bitumen mit Alu-Einlage, vollflächig aufgeschweisst", "m2", 18),
    ("30.120.01", "Abdichtung 2-lagig Polymerbitumen", "m2", 43),
    ("30.120.02", "Abdichtung 1-lagig Kunststoffbahn FPO", "m2", 39),
    ("30.130.01", "Aufbordung inkl. Keileinlage, Höhe bis 30 cm", "m", 52),
    ("30.130.02", "Aufbordung inkl. Keileinlage, Höhe über 30 cm", "m", 74),
    ("30.140.01", "Fenstertür mit Flüssigkunststoff abdichten, bis 350 mm", "Stk", 520),
    ("30.150.01", "Dachablauf inkl. Einbinden in Abdichtung", "Stk", 225),
    ("30.150.02", "Notüberlauf mit Tablettanschluss inkl. Kernbohrung", "Stk", 885),
    ("30.150.03", "Sanierungseinlauf mit Aufstockelement und Kontrollschacht", "Stk", 690),
    ("30.220.01", "Entwässerungsrinne vor Fenstertür inkl. Anschluss", "Stk", 480),
    ("30.160.01", "Abschottung", "Stk", 250),
    ("30.170.01", "Drainagematte", "m2", 17),
    ("30.180.01", "Untergrundabnahme mit Protokoll", "Pau", 600),
    ("30.190.01", "Dichtheitsprüfung der Abdichtung (elektrische Leckortung)", "Pau", 680),
    ("30.200.01", "Schutz- und Trennlage", "m2", 6.5)]:
    add(c, t, u, p, supplier="" if c in ("30.180.01", "30.190.01", "30.130.01", "30.130.02", "30.140.01", "30.150.03", "30.220.01") else "S02")

# 40 insulation
for c, t, u, p in [
    ("40.100.01", "Dämmung PIR alukaschiert 80 mm", "m2", 46), ("40.100.02", "Dämmung PIR alukaschiert 120 mm", "m2", 58),
    ("40.100.03", "Dämmung PIR alukaschiert 160 mm", "m2", 71), ("40.100.04", "Dämmung PIR alukaschiert 200 mm", "m2", 84),
    ("40.110.01", "Gefälledämmung EPS kaschiert, im Mittel 60 mm", "m2", 42), ("40.110.02", "Gefälledämmung EPS kaschiert, im Mittel 100 mm", "m2", 55),
    ("40.120.01", "Dämmung XPS Umkehrdach 120 mm", "m2", 64),
    ("40.130.01", "Mineralwolle Fassade 140 mm", "m2", 52), ("40.130.02", "Mineralwolle Fassade 180 mm", "m2", 61)]:
    add(c, t, u, p, supplier="S02")

# 50 coverings
for c, t, u, p, s in [
    ("50.100.01", "Steinplatten in Splitt verlegen, inkl. Splitt", "m2", 260, "S08"),
    ("50.100.02", "Feinsteinzeugplatten 60/60/2 auf Stelzlager, inkl. Lager", "m2", 185, "S08"),
    ("50.100.03", "Betonplatten 50/50 in Splitt verlegen", "m2", 128, "S08"),
    ("50.100.04", "Bestehende Platten in neuem Splitt wieder verlegen", "m2", 115, ""),
    ("50.150.01", "Ersatzplatten liefern, passend zum Bestand (ohne Verlegen)", "m2", 145, "S08"),
    ("50.110.01", "Holzrost Lärche auf Unterkonstruktion", "m2", 295, ""),
    ("50.120.01", "Rundkies 16/32, Schicht 5 cm", "m2", 14, ""),
    ("50.130.01", "Extensivbegrünung, Substrat 8 cm, inkl. Ansaat", "m2", 58, ""),
    ("50.140.01", "Kiesstreifen als Randabschluss", "m", 24, "")]:
    add(c, t, u, p, supplier=s)

# 60 sheet metal: variants by material
MATS = [("01", "Kupfer"), ("02", "Uginox"), ("03", "Titanzink"), ("04", "Aluminium")]
for base, t, u, prices in [
    ("60.100", "Dachrandabschluss Winkelblech, Zuschnitt bis 333 mm", "m", [128, 104, 86, 74]),
    ("60.110", "Mauerabdeckung, Zuschnitt bis 400 mm", "m", [168, 138, 112, 98]),
    ("60.120", "Deckstreifen inkl. Kittfuge", "m", [58, 52, 44, 40]),
    ("60.130", "Hängerinne halbrund, Zuschnitt 333 mm", "m", [112, 104, 82, None]),
    ("60.140", "Ablaufrohr Ø 100 mm", "m", [96, 88, 68, None]),
    ("60.150", "Einlaufstutzen oder Speier", "Stk", [145, 138, 115, None])]:
    for (v, m), p in zip(MATS, prices):
        if p:
            add(f"{base}.{v}", f"{t}, {m}", u, p, supplier="S03")
add("60.160.02", "Türschwellenblech, Uginox", "Stk", 185, supplier="S03")
add("60.170.04", "Fensterbank Aluminium, pulverbeschichtet", "m", 78, supplier="S03")

# 70 steel and semi finished (material articles, priced per kg)
ST = (1.85, 2.20); S355 = 0.15; INOX = (6.40, 7.60)
def steel(code, t, kg, grade="S235", inox=False):
    ek, vk = INOX if inox else ST
    if grade == "S355": ek, vk = ek + S355, vk + S355
    add(code, t, "kg", vk, kind="M", ek=ek, supplier="S01", kg_m=kg, text2=f"{grade}, {kg} kg/m")
series = [("70.100", "HEA", [100, 120, 140, 160, 180, 200], [16.7, 19.9, 24.7, 30.4, 35.5, 42.3]),
          ("70.110", "IPE", [100, 120, 140, 160, 180, 200], [8.1, 10.4, 12.9, 15.8, 18.8, 22.4]),
          ("70.120", "UNP", [100, 120, 140, 160, 180, 200], [10.6, 13.4, 16.0, 18.8, 22.0, 25.3])]
for base, n, sizes, kgs in series:
    for i, (sz, kg) in enumerate(zip(sizes, kgs), 1):
        steel(f"{base}.{i:02d}", f"{n} {sz}", kg)
        if n == "HEA" and sz >= 160:
            steel(f"{base}.{i+10:02d}", f"{n} {sz}", kg, grade="S355")
for i, (dim, kg) in enumerate([("60x40x3", 4.35), ("80x40x3", 5.29), ("100x50x4", 8.59), ("120x60x4", 10.5)], 1):
    steel(f"70.130.{i:02d}", f"Rechteckrohr {dim}", kg)
for i, (dim, kg) in enumerate([("40x40x3", 3.41), ("50x50x3", 4.35), ("60x60x4", 6.97), ("80x80x4", 9.22), ("100x100x5", 14.4)], 1):
    steel(f"70.140.{i:02d}", f"Quadratrohr {dim}", kg)
for i, (dim, kg) in enumerate([("40x8", 2.51), ("50x10", 3.93), ("60x10", 4.71), ("80x10", 6.28)], 1):
    steel(f"70.150.{i:02d}", f"Flachstahl {dim}", kg)
for i, (dim, kg) in enumerate([("Ø 12", 0.888), ("Ø 16", 1.58), ("Ø 20", 2.47)], 1):
    steel(f"70.160.{i:02d}", f"Rundstahl {dim}", kg)
for i, (dim, kg) in enumerate([("Ø 42.4 x 2.6", 2.55), ("Ø 48.3 x 2.6", 2.93)], 1):
    steel(f"70.170.{i:02d}", f"Stahlrohr {dim}", kg)
for i, (t, kg) in enumerate([(3, 23.6), (4, 31.4), (5, 39.3), (6, 47.1), (8, 62.8), (10, 78.5)], 1):
    add(f"70.180.{i:02d}", f"Stahlblech {t} mm", "kg", ST[1], kind="M", ek=ST[0], supplier="S01", text2=f"S235, {kg} kg/m2")
for i, (dim, kg) in enumerate([("Rohr Ø 42.4 x 2.0", 2.02), ("Flach 40x8", 2.56), ("Rohr Ø 33.7 x 2.0", 1.58)], 1):
    steel(f"70.190.{i:02d}", f"Chromstahl 1.4301 {dim}", kg, grade="1.4301", inox=True)

# 80 railing and stair components (bought parts)
for c, t, u, ek, vk, s in [
    ("80.100.01", "Gitterrost 30/30, verzinkt", "m2", 118, 145, "S01"),
    ("80.110.01", "Treppenstufe Gitterrost 800 mm, mit Rutschkante", "Stk", 79, 98, "S01"),
    ("80.110.02", "Treppenstufe Gitterrost 1000 mm, mit Rutschkante", "Stk", 95, 118, "S01"),
    ("80.120.01", "Glasfüllung VSG 2 x 6 mm, Zuschnitt", "m2", 180, 220, "S07"),
    ("80.120.02", "Glasklemmhalter Chromstahl", "Stk", 14, 19, "S07"),
    ("80.130.01", "Lochblechfüllung 3 mm, verzinkt", "m2", 92, 115, "S01"),
    ("80.140.01", "Handlaufhalter Chromstahl", "Stk", 18, 24, "S01"),
    ("80.140.02", "Endkappe Handlauf", "Stk", 6.5, 9, "S01"),
    ("80.150.01", "Pfostenfussplatte 150 x 150 x 10 mit Bohrungen", "Stk", 22, 29, "S01"),
    ("80.160.01", "Edelstahlseil Ø 4 mm inkl. Spanner, Seilfüllung", "m", 24, 31, "S01")]:
    add(c, t, u, vk, kind="M", ek=ek, supplier=s)

# 85 surfaces
for c, t, u, ek, vk, s in [
    ("85.100.01", "Feuerverzinken", "kg", 0.95, 1.15, "S04"),
    ("85.110.01", "Pulverbeschichten RAL Standard", "m2", 28, 34, "S05"),
    ("85.110.02", "Pulverbeschichten Sonderfarbe oder Feinstruktur", "m2", 32, 39, "S05"),
    ("85.120.01", "Duplex: verzinkt und pulverbeschichtet", "m2", 58, 70, "S05"),
    ("85.130.01", "Mindermengenzuschlag Beschichtung", "Pau", 100, 120, "S05")]:
    add(c, t, u, vk, kind="M", ek=ek, supplier=s)

# 90 fasteners
for c, t, u, ek, vk in [
    ("90.100.01", "Schwerlastanker M12", "Stk", 5.1, 6.8), ("90.100.02", "Schwerlastanker M16", "Stk", 7.2, 9.6),
    ("90.110.01", "Injektionsanker M12 inkl. Mörtel", "Stk", 10.9, 14.5), ("90.110.02", "Injektionsanker M16 inkl. Mörtel", "Stk", 14.2, 18.9),
    ("90.120.01", "Konsole thermisch getrennt für Aussenwärmedämmung", "Stk", 51, 68),
    ("90.130.01", "Schrauben Chromstahl, Paket", "Pak", 26, 34), ("90.140.01", "Montagekleinmaterial", "Pau", 60, 80),
    ("90.150.01", "Fugendichtband vorkomprimiert", "m", 4.2, 5.6)]:
    add(c, t, u, vk, kind="M", ek=ek, supplier="S06")

# 95 facade
for c, t, u, p in [
    ("95.100.01", "Unterkonstruktion Aluminium, justierbar", "m2", 48),
    ("95.110.01", "Fassadenbekleidung Aluminium-Verbundplatte", "m2", 165),
    ("95.110.02", "Fassadenbekleidung Trapezblech", "m2", 72),
    ("95.110.03", "Fassadenbekleidung Faserzementplatten", "m2", 118),
    ("95.120.01", "Fensteranschlussblech Aluminium", "m", 64),
    ("95.130.01", "Lüftungsgitter Sockel und Attika", "m", 22),
    ("95.140.01", "Eckausbildung Fassadenbekleidung", "m", 46)]:
    add(c, t, u, p, supplier="S09")

# 97 system products from the catalogue
for fam in CAT["families"].values():
    pre = fam["code_prefix"]
    for i, m in enumerate(fam["models"], 1):
        add(f"{pre}.{i:02d}", m["de"] if m["de"].split()[0] in fam["name"] else f'{fam["name"]} {m["de"]}', m["unit"], m["price"], text2=m["code"])
    for i, o in enumerate(fam["options"], 51):
        price = o.get("price", 0)
        add(f"{pre}.{i:02d}", f'{fam["name"]} Option: {o["de"]}', o["unit"] if "price" in o else "%",
            price, text2=o["code"] + (f' ({o["pct"]*100:+.0f} %)' if "pct" in o else ""))

# assembly flat fees of the catalogue as articles (DL-168)
ASM = {"velo": [("81", "Montage erstes Modul", "first_module"), ("82", "Montage jedes weitere Modul", "each_additional")],
       "carport": [("81", "Montage Carport CP-E", "CP-E"), ("82", "Montage Carport CP-EL", "CP-EL"), ("83", "Montage Carport CP-D", "CP-D")],
       "vordach": [("81", "Montage Vordach", "flat")]}
for key, fam in CAT["families"].items():
    for nr, txt, k in ASM[key]:
        add(f'{fam["code_prefix"]}.{nr}', f'{fam["name"]} {txt}', "Pau", fam["assembly"][k], text2="Montagepauschale")

# 99 time and material
for i, (k, v) in enumerate(OPS["rate_categories"].items(), 1):
    add(f"99.100.{i:02d}", f'Regie {v["de"]}', "Std", v["rate"], kind="R", text2=v["code"])
add("99.110.01", "Fahrzeugpauschale pro Tag", "Tag", OPS["surcharges_regie"]["vehicle_day"], kind="R")
add("99.110.02", "Zuschlag Nacht- und Wochenendarbeit", "%", 50, kind="R")

# additional service and material articles
for c, t, u, p, s in [
    ("10.190.01", "Baustrom, Anschluss und Verteiler", "Pau", 260, ""),
    ("10.190.02", "Abdecken und Schützen angrenzender Bauteile", "Pau", 780, ""),
    ("30.120.03", "Flüssigabdichtung Balkon, 2-komponentig, inkl. Vlies", "m2", 98, "S02"),
    ("30.130.03", "Anschluss Abdichtung an Fassade mit Klemmprofil", "m", 42, "S02"),
    ("30.200.02", "Wurzelschutzbahn", "m2", 12, "S02"),
    ("30.210.01", "Flachdachentlüfter", "Stk", 165, "S02"),
    ("60.170.01", "Fensterbank Kupfer", "m", 94, "S03"), ("60.170.02", "Fensterbank Uginox", "m", 88, "S03"),
    ("60.180.01", "Kamineinfassung, Kupfer", "Stk", 640, "S03"), ("60.180.02", "Kamineinfassung, Uginox", "Stk", 610, "S03"),
    ("60.185.02", "Dunstrohreinfassung, Uginox", "Stk", 145, "S03"),
    ("60.190.03", "Schneefang Rundrohr, Titanzink", "m", 58, "S03"),
    ("60.195.01", "Laubfanggitter", "Stk", 28, "S03"),
    ("80.200.01", "Montage Geländer, Pfosten aufgedübelt", "m", 68, ""),
    ("80.200.02", "Montage Geländer, Pfosten seitlich befestigt", "m", 79, ""),
    ("80.200.03", "Montage Handlauf an Wand", "m", 32, ""),
    ("80.200.04", "Kernbohrung Ø 50 bis 100 mm", "Stk", 85, ""),
    ("80.200.05", "Ankerplatte versetzen und vergiessen", "Stk", 45, ""),
    ("80.200.06", "Geländer demontieren, zwischenlagern und wieder montieren", "m", 95, ""),
    ("80.200.07", "Geländer erhöhen, Aufsatz an bestehenden Pfosten", "m", 145, ""),
    ("95.150.01", "Montage Fassadenbekleidung, Zuschlag Höhe über 8 m", "m2", 14, "")]:
    add(c, t, u, p, supplier=s)
for c, t, u, ek, vk in [
    ("90.160.01", "Gewindestange M12 verzinkt", "m", 3.1, 4.2), ("90.160.02", "Gewindestange M16 verzinkt", "m", 5.4, 7.2),
    ("90.170.01", "Holzschrauben Chromstahl 6 x 80, Paket", "Pak", 31, 41),
    ("90.180.01", "Fugendichtstoff MS-Polymer, Kartusche", "Stk", 9.8, 13.5)]:
    add(c, t, u, vk, kind="M", ek=ek, supplier="S06")
for c, t, u, p in [
    ("98.100.01", "Zustandsaufnahme vor Ort, Grundpauschale bis 2 Std", "Pau", 390),
    ("98.100.02", "Drohnenaufnahme Dach und Fassade", "Pau", 650),
    ("98.100.03", "Thermografie Gebäudehülle", "Pau", 480),
    ("98.100.04", "Feuchtemessung Dachaufbau, bis 5 Messstellen", "Pau", 320),
    ("98.100.05", "Sondierung Dachaufbau mit Probeöffnung und Verschluss", "Stk", 280),
    ("98.110.01", "Kurzbericht Zustand (bis 5 Seiten)", "Pau", 750),
    ("98.110.02", "Variantenstudie Sanierung mit Kostenschätzung", "Pau", 1850),
    ("98.120.01", "Spesen Beratung, km-Pauschale", "km", 0.9)]:
    add(c, t, u, p, kind="L")
    if c.startswith("98.110"): items[-1]["blocked"] = True    # report lump sums from 2016, replaced by fees by effort (DL-187)

# legacy articles blocked after the 2021 migration (realistic artefact)
legacy = [("30.120.91", "Abdichtung 2-lagig (ALT, nicht mehr verwenden)", "m2", 39),
          ("40.100.91", "PIR 80mm alukasch.", "m2", 41), ("50.100.91", "Steinplatten Splitt", "m2", 235),
          ("60.100.91", "Winkelblech CU", "m", 86), ("70.130.91", "RR 60/40/3", "kg", 1.95),
          ("85.100.91", "Verzinken", "kg", 0.98), ("97.100.91", "Velounterstand Modul (alt)", "Stk", 3100),
          ("97.200.91", "Carport Einzel (alt)", "Stk", 7600), ("99.100.91", "Monteur", "Std", 108),
          ("10.120.91", "Pneukran", "Std", 395), ("20.120.91", "Rückbau Flachdach", "m2", 44),
          ("30.150.91", "Ablauf", "Stk", 190)]
for c, t, u, p in legacy:
    add(c, t, u, p, kind="L"); items[-1]["blocked"] = True
steel("70.150.05", "Flachstahl 200x12", 18.84)          # stair stringers, added 2026 (C2); appended so earlier random draws stay stable
add("80.110.03", "Treppenstufe Gitterrost 1000 mm mit Seitenblechen, zum Anschrauben", "Stk", 165, kind="M", ek=135, supplier="S01")
add("80.170.01", "Torbeschlag selbstschliessend mit Kindersicherung", "Stk", 230, kind="M", ek=180, supplier="S06")
add("10.120.03", "Lastwagen mit Ladekran (eigener), inkl. Chauffeur", "Std", 165, kind="L")
for c, t in [("97.200.61", "Fronalp Carport Zuschlag Überbreite (2.5 % je angefangene 10 cm)"),
             ("97.200.62", "Fronalp Carport Zuschlag Durchfahrtshöhe 2.40 m (6 %)"),
             ("97.200.63", "Fronalp Carport Zuschlag Ausführung Verstärkt 801 bis 1200 m ü. M. (18 %)"),
             ("97.200.64", "Fronalp Carport Farbzuschlag andere RAL-Farbe (8 %)"),
             ("97.200.65", "Fronalp Carport Farbzuschlag DB-Farbe (12 %)")]:
    add(c, t, "Pau", 0, kind="L")                      # amount set per offer from the price list rule (S2)

# legacy KALKO numbers: about 60 % of articles that existed before 2021
idx_year = OPS["price_index"]["years"]
for it in items:
    it["family"] = GROUPS[it["group"]][1]
    old = it.get("blocked") or (it["group"] not in ("97",) and rng.random() < 0.6) or (it["group"] == "97" and "Option" not in it["text"] and rng.random() < 0.7)
    it["alt_nr"] = f"K{rng.randint(1000, 9899)}" if old else ""
    if it["alt_nr"] and rng.random() < 0.05:
        it["alt_nr"] = it["alt_nr"][:-1]          # truncated during migration
    it["valid_from"] = "01.01.2026" if not it.get("blocked") else "01.03.2021"

# canonical source
src = ROOT / "generation/world/item_master.csv"
with open(src, "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=["code", "group", "kind", "family", "text", "text2", "unit", "vk", "ek",
                                      "supplier", "kg_m", "alt_nr", "blocked"])
    w.writeheader()
    for it in items:
        w.writerow({k: ("" if it.get(k) is None else it.get(k, "")) for k in w.fieldnames})

# ERP export as staff see it
def chf(x):
    return "" if x is None else f"{x:.2f}".replace(".", ".")
out = ROOT / "company-world/systems"; out.mkdir(parents=True, exist_ok=True)
sup = {s["id"]: s["name"] for s in OPS["suppliers"]}
with open(out / "MetaBau_Export_Artikelstamm_2026-04-30.csv", "w", newline="", encoding="cp1252") as f:
    w = csv.writer(f, delimiter=";")
    w.writerow(["ART_NR", "WG", "ART_TYP", "BEZ_1", "BEZ_2", "ME", "VK_PREIS", "EK_PREIS", "KG_M",
                "LIEF_NR", "LIEFERANT", "ALT_NR", "GUELTIG_AB", "STATUS"])
    for it in sorted(items, key=lambda x: x["code"]):
        w.writerow([it["code"], it["group"], it["kind"], it["text"][:60], it["text2"], it["unit"],
                    chf(it["vk"]), chf(it["ek"]), "" if it["kg_m"] is None else it["kg_m"],
                    it["supplier"], sup.get(it["supplier"], ""), it["alt_nr"], it["valid_from"],
                    "GESPERRT" if it.get("blocked") else "AKTIV"])
print(len(items), "articles,", sum(1 for i in items if i.get("blocked")), "blocked,",
      sum(1 for i in items if i["alt_nr"]), "with legacy number")
