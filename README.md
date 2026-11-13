# RFQ Datensatz: Von der Anfrage zur Offerte

Synthetischer Datensatz für das KMU KI Programm des Kantons Zürich (SDSC). Er zeigt, wie in einem Schweizer KMU aus einer Kundenanfrage (Request for Quotation, RFQ) eine Offerte entsteht, mit allen Dokumenten, die dabei typischerweise anfallen: E-Mails, Formulare, Skizzen, Pläne, Aktennotizen, CRM und ERP Exporte, Preislisten, interne Richtlinien und frühere Projekte.

Firma, Personen, Kunden, Adressen, Telefonnummern, Bankverbindungen und Preise sind frei erfunden. Ähnlichkeiten mit realen Firmen oder Personen sind nicht beabsichtigt.

## Wozu dient der Datensatz?

Im Workshop nutzen wir den Datensatz, um

1. den Ablauf von der Anfrage bis zur Offerte in einem KMU nachzuvollziehen,
2. selbst eine Offerte aus den vorhandenen Unterlagen herzuleiten,
3. zu erkennen, welche Schritte sich mit KI unterstützen oder automatisieren lassen und wo Erfahrungswissen, Rückfragen oder Entscheide nötig bleiben.

Zu jedem Fall gibt es eine Musterlösung. Damit lässt sich prüfen, ob eine von Hand oder mit KI erstellte Offerte stimmt.

## Die Firma: Fronalp Metall & Gebäudehülle AG

Fronalp ist ein fiktives Familienunternehmen in Ibach SZ, gegründet 1962, mit rund 35 Mitarbeitenden. Es hat drei Geschäftsbereiche:

| Bereich | Was | Typische Offerte |
|---|---|---|
| Beratung | Zustandsaufnahmen, Sanierungskonzepte | Honorarofferte |
| Projektgeschäft | Massanfertigungen: Terrassen, Flachdächer, Spenglerarbeiten, Geländer, Treppen | Einheitspreis oder Pauschalofferte |
| Systemprodukte | Katalogprodukte mit kleinen Anpassungen: Velounterstand, Carport, Vordach | Konfigurationsofferte |

Seit 2021 arbeitet Fronalp mit dem ERP MetaBau, seit 2022 mit dem CRM KundenCockpit. Ältere Unterlagen stammen aus der Zeit davor und sehen entsprechend anders aus. Wie in jedem gewachsenen Betrieb sind nicht alle Unterlagen aktuell und widerspruchsfrei. Das ist gewollt.

## Aufbau des Repositorys

**Für den Workshop genügt ein einzelner Fall.** Wichtig sind die Unterlagen in `cases/<Fall>/request/` und `cases/<Fall>/operations/`. Aus `company-world/` braucht man nur, was der Fall verlangt, zum Beispiel die Preisliste zum angefragten Produkt oder den Kunden im Debitorenexport. Bei S1 führt die Anleitung weiter unten Schritt für Schritt zu diesen Unterlagen.

Die Firma Fronalp muss man nicht kennenlernen. `company-world/` ist zum Nachschlagen da, nicht zwingend zum Durchlesen. `docs/` und `generation/` braucht man für den Workshop nicht.

| Ordner | Inhalt |
|---|---|
| `company-world/organisation/` | Firmenportrait, Organigramm, Personalliste |
| `company-world/operations/` | Preislisten Systemprodukte, Kalkulationsrichtlinie, Faustregeln des Kalkulators, Regieansätze, AGB, Betriebshandbuch, Vorlagen |
| `company-world/systems/` | Exporte aus CRM und ERP per 30.04.2026: Kontakte, Leads, Debitoren, Artikelstamm, Offerten, Projekte, Stunden |
| `company-world/regulations/` | Normenverzeichnis und Auszüge aus OR, MWSTG und BauAV |
| `company-world/past-projects/` | Abgeschlossene Projekte mit Offerte und Abrechnung |
| `cases/<Fall>/request/` | Die Anfrage des Kunden und interne Weiterleitungen |
| `cases/<Fall>/operations/` | Interne Unterlagen während der Bearbeitung, zum Beispiel ein CRM Ausdruck |
| `cases/<Fall>/offer/` | Die versendete Offerte und das Begleitmail |
| `cases/<Fall>/case.yaml` | Story, Fakten, jede Offertposition mit Quelle, erwartete Beträge |
| `docs/` | Glossar der Begriffe, Entscheide (ADR), Designlog, Qualitätsprüfungen |
| `generation/` | Quelldaten und Skripte, aus denen alle Dokumente erzeugt werden |

**Wichtig:** Wer die Offerte selbst herleiten will, öffnet `offer/` und `case.yaml` erst am Schluss.

### Hinweise zu den Dateiformaten

* E-Mails sind `.eml` Dateien. Sie lassen sich mit Outlook, Apple Mail oder Thunderbird öffnen. Anhänge (PDF, Bilder, JSON) stecken nur im Mail selbst.
* Die CSV Exporte stammen aus einem Windows ERP: Trennzeichen Semikolon, Zeichensatz Windows 1252. In Excel über *Daten, Aus Text/CSV* importieren, sonst werden Umlaute falsch dargestellt.
* Lange Artikeltexte sind im ERP Export auf 60 Zeichen gekürzt, wie in vielen realen Systemen.

## Die Fälle

| Fall | Thema | Bereich | Schwierigkeit | Offerte |
|---|---|---|---|---|
| **S1** | Velounterstand für 20 Velos, Wohngenossenschaft Rossberg, Goldau | Systemprodukte | einfach | OF-26-0116 |
| **C1** | Terrassensanierung einer Attikawohnung, Seestrasse 18, Brunnen | Projektgeschäft | mittel | OF-26-0178 |
| **A1** | Zustandsaufnahme und Sanierungsvarianten zweier Flachdächer, Hofmattweg 2 und 4, Steinen | Beratung | mittel | OF-26-0147 (Honorarofferte) |
| **C2** | Aussentreppe Stahl zum Balkon, Einfamilienhaus Gibelweg 7, Morschach | Projektgeschäft | schwer | OF-26-0209 (Pauschalofferte) |
| **S2** | Doppelcarport mit Anpassungen, Sonnenhaldenweg 6, Oberiberg | Systemprodukte | mittel | OF-26-0228 |
| **A3** | Prüfung der Balkon- und Treppenhausgeländer, Siedlung Rütimatt, Schwyz | Beratung | mittel | OF-26-0254 (Honorarofferte nach Richtofferte) |

**S1 ist der saubere Basisfall.** Der Kunde füllt das Webformular vollständig aus und lädt eine Handskizze hoch. Alle nötigen Angaben sind vorhanden, es gibt keine absichtlich eingebauten Fehler. Die Offerte entsteht aus Katalog, Preisliste und wenigen Regeln. S1 eignet sich, um den Ablauf und die Unterlagen kennenzulernen.

**C1 ist ein realistischer, unordentlicher Fall.** Die Anfrage kommt vom Wohnungseigentümer, Besteller ist aber die Stockwerkeigentümergemeinschaft über ihre Verwaltung. Die Informationen verteilen sich auf acht Mails, eine Telefonnotiz, einen gescannten Plan und eine Besprechungsnotiz. Mengen müssen aus dem Plan bestimmt werden, und mehrere Positionen beruhen auf Erfahrungsregeln des Kalkulators. Der Fall enthält bewusst eingebaute Unstimmigkeiten, wie sie in echten Dossiers vorkommen. Die Auflösung steht in `case.yaml` unter `flaws`, also in der Lösung.

**A1 ist ein Beratungsfall.** Eine Verwaltung fragt im Auftrag einer Stiftung nach einer Zustandsaufnahme, der Umfang ändert sich im Mailverlauf zweimal. Für Beratungsstunden gibt es keine Tabelle: Das Honorar wird aus zwei früheren ähnlichen Offerten und ihren effektiven Stunden geschätzt. Geprüft wird deshalb der Inhalt der Offerte und ein Honorarband von ±20 % um die Musterlösung.

**C2 ist ein schwieriger Fall für einen Privatkunden.** Die Masse im ersten Mail und in der Handskizze sind veraltet, der Kunde nennt ein Budget, und die alte Word-Vorlage enthält veraltete Werte. Die Pauschalofferte wird von unten nach oben kalkuliert (Kalkulationsrichtlinie und Faustregeln) und enthält eine günstigere Variante.

**S2 ist ein Systemprodukt mit Anpassungen.** Ein Privatpaar wünscht einen Doppelcarport. Mehrere Regeln der Preisliste greifen gleichzeitig (Überbreite, Ausführung nach Höhenlage, Durchfahrtshöhe, Farbe), und die Angaben der Kunden stimmen nicht überall mit dem Situationsplan und den späteren Mails überein. Dazu kommt die Frage nach einer Schneelastberechnung.

**A3 ist ein Beratungsfall in zwei Schritten.** Eine Wohnbaugenossenschaft erhält zuerst eine unverbindliche Richtofferte, danach wird die verbindliche Honorarofferte verlangt. Zwischen den beiden ändert sich der Umfang: Eine telefonische Abmachung steht im Widerspruch zu einem späteren Mail, die genannte Zahl der Balkone stimmt nicht, und im CRM ist der Kunde doppelt erfasst. Das Honorar wird wie bei A1 aus früheren ähnlichen Projekten geschätzt.

Weitere Fälle (weitere Projekte und Beratungen) folgen.

## S1 Schritt für Schritt: Die Offerte selbst erstellen

Ziel: Aus der Anfrage und den Firmenunterlagen die Offerte OF-26-0116 vom 08.05.2026 herleiten. Dafür reichen Excel oder Papier und Taschenrechner.

Der Fall zeigt, welche Informationen zusammenkommen müssen, damit überhaupt eine Offerte entstehen kann. Dieselben Informationen braucht jeder Betrieb, der Offerten automatisch erstellen will.

### 1. Anfrage lesen

Öffne `cases/S1/request/2026-05-05_Webformular_Anfrage_Systemprodukte.eml`. Das Mail enthält den Formulartext, einen JSON Export des Formulars und das Foto einer Handskizze. Notiere:

* Produkt und Anzahl Velos
* verfügbare Länge und Untergrund
* gewünschte Optionen: Rückwand, Seitenwände, Veloständer, Beleuchtung, Fundamente, Dachrinne, Farbe
* Standort des Objekts
* Stand der Baubewilligung, Wunschtermin
* Bemerkungen des Kunden (es gibt eine Frage, die in der Antwort beantwortet werden muss)

Das Weiterleitungsmail `2026-05-06_Mail_EFoe_Weiterleitung.eml` zeigt, wie die Administration den Lead erfasst.

### 2. Kunde prüfen

* `cases/S1/operations/KundenCockpit_Lead_L-2026-152_Druck.pdf`: Lead und Kundenhistorie im CRM.
* `company-world/systems/MetaBau_Export_Debitoren_2026-04-30.csv`: Suche den Debitor der Wohngenossenschaft. Die Spalten `RABATT_PROZ` und `SKONTO_SPERRE` bestimmen Rabatt und Skonto.

### 3. Produkt konfigurieren

Öffne `company-world/operations/Preisliste_Systemprodukte_2026.pdf`, Abschnitt Fronalp Velo. Auf dem Server liegt auch die Preisliste 2024. Achte auf das Gültigkeitsdatum: Massgebend ist die Liste, die am Offertdatum gilt.

* Wie viele Module braucht es? Beachte die Anzahl Velos pro Modul und die Modullänge, und prüfe, ob die Gesamtlänge auf dem verfügbaren Platz Raum findet.
* Wie teilen sich die Module auf Startmodul und Anbaumodule auf?
* Welche Optionen werden pro Modul, welche pro Stück oder pro Anlage verrechnet? Lies dazu die Anpassungen unter der Preistabelle.
* Montage: Pauschale für das erste Modul und für jedes weitere.
* Transport: In welcher Zone liegt der Standort?

### 4. Artikelnummern zuordnen

Suche jede Position in `company-world/systems/MetaBau_Export_Artikelstamm_2026-04-30.csv`. Die Velounterstand Artikel beginnen mit `97.100.`, die Transportpauschalen mit `10.170.`. Die Spalte `BEZ_2` enthält den Katalogcode aus der Preisliste.

### 5. Firmenregeln anwenden

Lies in `company-world/operations/Kalkulationsrichtlinie_v3.2_2024-03.pdf` und `company-world/operations/Faustregeln_BUl.xlsx` nach:

* Rabatt: Wann gibt es wie viel? Was gilt für Genossenschaften?
* Skonto: Auf welchen Betrag wird er berechnet?
* Rundung: Welche Beträge werden auf 5 Rappen gerundet?
* Mehrwertsteuer: Welcher Satz gilt am Offertdatum?
* Zahlungskonditionen und Gültigkeit der Offerte
* Was ist bauseits zu leisten, welche Annahmen und Vorbehalte gehören in die Offerte?
* Was bietet man zusätzlich als Eventualposition an? Eventualpositionen stehen in Klammern und zählen nicht zum Total.

### 6. Offerte rechnen

Lege eine Tabelle an mit Position, Artikelnummer, Bezeichnung, Menge, Einheit, Einzelpreis und Betrag. Berechne danach:

1. Total brutto (Summe aller Positionen ohne Eventualpositionen)
2. Rabatt in % vom Total brutto
3. Skonto in % vom Total brutto
4. Zwischentotal = brutto minus Rabatt minus Skonto
5. Mehrwertsteuer auf das Zwischentotal
6. Total inkl. MWST

Wer Python nutzt, kann die Rechnung mit der Preislogik des Repositorys prüfen. Die Positionen sind hier bewusst leer gelassen:

```python
# im Ordner generation/scripts ausführen
from pricing import World, compute_offer
positionen = [
    # ("Artikelnummer", Menge),
]
r = compute_offer([{"code": c, "qty": q} for c, q in positionen],
                  "2026-05-08", discount=0.0, skonto=True, world=World())
print(r["gross"], r["discount"], r["skonto"], r["net"], r["vat"], r["total"])
```

### 7. Vergleichen

Erst jetzt öffnen:

* `cases/S1/offer/OF-26-0116_Offerte_Velounterstand_Rossberg.pdf`: die versendete Offerte
* `cases/S1/offer/2026-05-08_Mail_DIn_Offerte.eml`: das Begleitmail mit der Antwort auf die Frage des Kunden
* `cases/S1/case.yaml`: jede Position mit ihrer Quelle (`src`) und die erwarteten Beträge (`expected`)

Die fünf Stellen, an denen bei S1 am häufigsten Fehler entstehen:

1. Anzahl Module aus der Anzahl Velos (aufrunden)
2. Veloständer pro Modul statt pro Velo
3. Transportzone des Standorts
4. Umgang mit der Rabattfrage des Kunden
5. Zweite Seitenwand als Eventualposition

### Zum Weiterdenken im Workshop

* Welche dieser Schritte könnte ein KI System übernehmen, welche nur vorbereiten?
* Welche Informationen müsste das System lesen können (Mail, JSON, Bild, PDF, CSV, Excel)?
* Wo braucht es eine Regel, die heute nur im Kopf des Kalkulators existiert?
* Wie würde man prüfen, ob eine automatisch erstellte Offerte stimmt?
* Was ändert sich beim Fall C1, wo die Angaben unvollständig und widersprüchlich sind?

## Grenzen des Datensatzes

* Fotos von Begehungen fehlen. In der Praxis gehören sie oft zu Anfragen und Zustandsaufnahmen; hier beschreiben Aktennotizen die Befunde.
* Bei Beratungsfällen gibt es keine eindeutige Lösung. Das Honorar wird aus früheren ähnlichen Offerten geschätzt; geprüft werden der Inhalt und ein Honorarband um die Musterlösung.
* Alle Dokumente sind synthetisch erzeugt und sauberer gesetzt als echte Scans, Fotos oder Handnotizen.

## Hintergrund

* `GLOSSARY.md`: verbindliche Begriffe (englisch, mit deutscher Entsprechung)
* `docs/plan.md`: Gesamtplan des Datensatzes und geplante Fälle
* `docs/adr/`: grundlegende Entscheide, zum Beispiel warum Normen nur genannt und nie für Preise benötigt werden
* `docs/design-log.md`: alle Designentscheide mit Begründung
* `docs/qa/`: Ergebnisse der Qualitätsprüfungen pro Fall
* `generation/README.md`: wie die Dokumente aus den Quelldaten erzeugt werden (`bash generation/scripts/build_all.sh`, danach `python3 generation/scripts/build_case.py S1`)

Die Dokumente sind in Schweizer Hochdeutsch verfasst, ohne ß, mit Schweizer Zahlenformat (1'234.50) und Schweizer Recht. Die Dokumentation zum Aufbau des Datensatzes ist auf Englisch.

## Lizenz

© 2026 Swiss Data Science Center (SDSC), erstellt für das KMU KI Programm des Kantons Zürich. Für den Datensatz und die Dokumentation gilt [CC BY 4.0](LICENSE-DATA), für die Skripte in `generation/scripts/` die [MIT Lizenz](LICENSE-CODE). Die Gesetzesauszüge in `company-world/regulations/` und `generation/sources/law/` sind amtliche Erlasse und nach Art. 5 URG nicht urheberrechtlich geschützt. Details stehen in [LICENSE](LICENSE).

Bitte zitieren als: *RFQ Datensatz: Von der Anfrage zur Offerte*, Swiss Data Science Center (SDSC) für das KMU KI Programm des Kantons Zürich, CC BY 4.0.
