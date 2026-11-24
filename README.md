# NexvoraHub Studio

*Electrical load design studio* – load assessment, design reports, electrical drawing generation and AutoCAD (DXF) drawing review.

Electrical load assessment and design tools for buildings, usable across 94 countries and regions (UK BS 7671 / IEC 60364, NEC, CEC, AS/NZS 3000 families).

## Contents

| Path | What it is |
|---|---|
| `Electrical Load Design Report.xlsx` | Excel design template (27 sheets): regional load schedules, NEC/CEC code calculation, equipment loads with demand/diversity/simultaneity factors, lifts, pumps, data centre, BS 7671 cable and protection, transformers and switchrooms, energy management, compliance checklist, utility application data. |
| `_build/` | Node.js scripts that generate the workbook. |
| `app/load-designer.html` | Single-file web app (NexvoraHub Studio): load assessment, design report, AI data-centre planning, electrical drawing generator (DXF), drawing review with legend learning, panel-schedule import and design on submitted layouts, and a facility management performance board. |
| `app/src/` | App source, AI data-centre library (`ai-dc-data.json`), data extractor, local preview server and headless tests. |
| `site/index.html` | Standalone build of the app for web hosting. Only this folder is published. |

## Facility management performance board

Stages 13–14 of the app. **FM data input** takes the reporting period, targets, and one row per hard or soft service contract (contractor, contract reference and dates, annual value, response time, SLA target, reactive calls raised / completed / completed within SLA / overdue, PPM tasks planned and completed, KPI audit score), plus assets (failures, repair hours, downtime), an incident log, SOP compliance, CAPEX and OPEX, customer satisfaction, energy and water, safety and fire, and optimisation strategies. **FM performance board** turns these into a tile grid: executive summary, SLA / PPM / MTTR / MTBF / availability / satisfaction gauges, incident summary, services and contract compliance (with contracts due for renewal within 90 days flagged), SOP compliance, MTTR and MTBF by equipment, energy and CO₂, CAPEX and OPEX, work orders, asset management, and safety and fire. The board filters to hard services, soft services, or both, exports to CSV and prints to PDF. Floor area and the CO₂ factor default to the area load schedule and the project region.

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

`build-app.js` writes both `app/load-designer.html` (claude.ai Artifact) and `site/index.html` (website). Commit and push `site/index.html` to redeploy the live site: the GitHub Actions workflow `.github/workflows/pages.yml` publishes `site/` to GitHub Pages at https://kayglor.github.io/Nexvorahub_studio_load_designer/ (one-time setup: Settings → Pages → Source: GitHub Actions).

`extract-data.js` reads the reference data (countries, densities, factors, cable tables) straight from the workbook generator, so the app and the workbook stay consistent. `serve.js` previews the app at http://localhost:5179/.

## Important

Load densities, regional factors, demand factors, code unit loads and cable ratings are indicative engineering defaults. Verify them against the edition of each standard adopted locally before issuing a design. Canadian (CEC) basic loads are placeholders. Cable ratings are BS 7671 Appendix 4 copper values; use local tables for NEC, CEC and AS/NZS projects. The outputs support, and do not replace, design verification and certification by a competent person.
