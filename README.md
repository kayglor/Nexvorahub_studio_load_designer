# Electrical Load Designer

Electrical load assessment and design tools for buildings, usable across 94 countries and regions (UK BS 7671 / IEC 60364, NEC, CEC, AS/NZS 3000 families).

## Contents

| Path | What it is |
|---|---|
| `Electrical Load Design Report.xlsx` | Excel design template (27 sheets): regional load schedules, NEC/CEC code calculation, equipment loads with demand/diversity/simultaneity factors, lifts, pumps, data centre, BS 7671 cable and protection, transformers and switchrooms, energy management, compliance checklist, utility application data. |
| `_build/` | Node.js scripts that generate the workbook. |
| `app/load-designer.html` | Single-file web app: load assessment and design report generator using the same design data. |
| `app/src/` | App source, data extractor, local preview server and headless tests. |

## Rebuild the workbook

```bash
cd _build
npm install exceljs@4.4.0
node assemble.js
node build_full.js
```

Open the result in Excel; it recalculates on open.

## Rebuild and test the app

```bash
cd app/src
node extract-data.js ../../_build/build_full.js designdata.json
node build-app.js designdata.json
node test-app.js
node serve.js
```

`extract-data.js` reads the reference data (countries, densities, factors, cable tables) straight from the workbook generator, so the app and the workbook stay consistent. `serve.js` previews the app at http://localhost:5179/.

## Important

Load densities, regional factors, demand factors, code unit loads and cable ratings are indicative engineering defaults. Verify them against the edition of each standard adopted locally before issuing a design. Canadian (CEC) basic loads are placeholders. Cable ratings are BS 7671 Appendix 4 copper values; use local tables for NEC, CEC and AS/NZS projects. The outputs support, and do not replace, design verification and certification by a competent person.
