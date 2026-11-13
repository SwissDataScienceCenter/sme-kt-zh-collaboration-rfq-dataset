# Generation

Everything in `company-world/` and `cases/` is produced from the sources in `generation/world/` by the scripts in `generation/scripts/`. Do not edit generated documents by hand; change the source and re-run.

## Rebuild the Company world

```
bash generation/scripts/build_all.sh
```

Order: item master, pricing self test, CRM and ERP exports, organisation, operations, templates and manuals, regulations.

Then build the Cases and restore files whose content did not change, so that git only shows real changes:

```
python3 generation/scripts/build_case.py C1
python3 generation/scripts/build_case.py S1
python3 generation/scripts/build_case.py A1
python3 generation/scripts/build_case.py C2
python3 generation/scripts/build_case.py S2
python3 generation/scripts/build_case.py A3
python3 generation/scripts/keep_unchanged.py
```

## Folders

| Folder | Content |
|---|---|
| world/ | Sources of truth (YAML, CSV) and `known_issues.yaml` |
| scripts/ | Builders and renderers; `pricing.py` is the only place where amounts are computed |
| specs/ | `case_schema.yaml` and `writing_style.md` for generating Cases |
| sources/ | External public domain inputs (law PDFs) |

Design decisions are logged in `docs/design-log.md`.
