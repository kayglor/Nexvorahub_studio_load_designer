const ExcelJS = require('exceljs');
const OUT = require('path').join(__dirname, '..', 'Electrical Load Design Report.xlsx');

const wb = new ExcelJS.Workbook();
wb.creator = 'Electrical design';
wb.calcProperties.fullCalcOnLoad = true;

const FONT = { name: 'Arial', size: 10 };
const NAVY = 'FF1F3864';
const hdrFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
const inFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
const totFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
const secFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } };
const thin = { style: 'thin', color: { argb: 'FFBFBFBF' } };
const box = { top: thin, left: thin, bottom: thin, right: thin };
const BLUE = 'FF0000FF', GREEN = 'FF008000';

function base(ws) {
  ws.eachRow({ includeEmpty: true }, r => r.eachCell({ includeEmpty: true }, c => {
    c.font = Object.assign({}, FONT, c.font && c.font.bold ? { bold: true } : {}, c.font && c.font.color ? { color: c.font.color } : {}, c.font && c.font.size && c.font.size !== 11 ? { size: c.font.size } : {}, c.font && c.font.italic ? { italic: true } : {});
  }));
}
function title(ws, text, sub) {
  ws.getCell('A1').value = text;
  ws.getCell('A1').font = { name: 'Arial', size: 14, bold: true, color: { argb: NAVY } };
  if (sub) { ws.getCell('A2').value = sub; ws.getCell('A2').font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } }; }
}
function header(ws, row, labels) {
  labels.forEach((l, i) => {
    const c = ws.getRow(row).getCell(i + 1);
    c.value = l; c.fill = hdrFill; c.border = box;
    c.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    c.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
  });
  ws.getRow(row).height = 30;
}
function input(c, v, fmt) {
  if (v !== undefined) c.value = v;
  c.fill = inFill; c.border = box; c.font = { name: 'Arial', size: 10, color: { argb: BLUE } };
  if (fmt) c.numFmt = fmt;
}
function calc(c, f, fmt, color) {
  c.value = { formula: f }; c.border = box;
  c.font = { name: 'Arial', size: 10, color: { argb: color || 'FF000000' } };
  if (fmt) c.numFmt = fmt;
}
function legend(ws, row) {
  ws.getCell(`A${row}`).value = 'Legend:';
  ws.getCell(`A${row}`).font = { name: 'Arial', size: 9, bold: true };
  const b = ws.getCell(`B${row}`); b.value = 'Input – edit these cells'; b.fill = inFill; b.font = { name: 'Arial', size: 9, color: { argb: BLUE } };
  const c = ws.getCell(`C${row}`); c.value = 'Formula – do not edit'; c.font = { name: 'Arial', size: 9 };
  const d = ws.getCell(`D${row}`); d.value = 'Link from another sheet'; d.font = { name: 'Arial', size: 9, color: { argb: GREEN } };
}

// ---------------- 1. Project Info ----------------
const pi = wb.addWorksheet('Project Info');
title(pi, 'Electrical Load Assessment – Design Report', 'Early-stage maximum demand estimate for supply application / feasibility. Replace with a detailed load schedule at detailed design.');
legend(pi, 3);
const piRows = [
  [5, 'PROJECT DETAILS', null, null, true],
  [6, 'Project name', 'Example: Riverside Mixed-Use Development'],
  [7, 'Project reference', 'EX-0001'],
  [8, 'Client', 'Example Client Ltd'],
  [9, 'Site address / postcode', 'Example Street, London'],
  [10, 'DNO / IDNO', 'UKPN / SSEN / NGED / NPg / ENWL / SPEN'],
  [11, 'Report revision', 'P01'],
  [12, 'Date', new Date(2026, 9, 18)],
  [13, 'Prepared by', ''],
  [14, 'Checked by', ''],
  [16, 'DESIGN PARAMETERS', null, null, true],
  [17, 'Supply voltage, line–line (V)', 400, 'Standard UK LV: 400 V 3-phase'],
  [18, 'Supply voltage, line–neutral (V)', 230, 'Standard UK LV: 230 V'],
  [19, 'Power factor (overall)', 0.95, 'Typical 0.9–0.95 for modern LED/VSD-dominated loads'],
  [20, 'Spare / future capacity allowance', 0.2, 'Typical 15–25%; higher where heat/EV electrification is expected'],
  [21, 'Density basis for lookup', 'Typical', 'Choose Min / Typical / Max from the VA-m2 Data sheet'],
];
piRows.forEach(([r, label, val, note, sec]) => {
  const a = pi.getCell(`A${r}`); a.value = label; a.font = { name: 'Arial', size: 10, bold: !!sec };
  if (sec) { ['A', 'B', 'C'].forEach(col => pi.getCell(`${col}${r}`).fill = secFill); return; }
  a.border = box;
  input(pi.getCell(`B${r}`), val);
  if (note) { pi.getCell(`C${r}`).value = note; pi.getCell(`C${r}`).font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } }; }
});
pi.getCell('B12').numFmt = 'dd/mm/yyyy';
pi.getCell('B19').numFmt = '0.00';
pi.getCell('B20').numFmt = '0%';
pi.getCell('B21').dataValidation = { type: 'list', allowBlank: false, formulae: ['"Min,Typical,Max"'] };
pi.getCell('B19').dataValidation = { type: 'decimal', operator: 'between', formulae: [0.5, 1], showErrorMessage: true, error: 'Power factor must be between 0.5 and 1.0' };
pi.getColumn(1).width = 36; pi.getColumn(2).width = 42; pi.getColumn(3).width = 62; pi.getColumn(4).width = 24;

// ---------------- 2. VA-m2 Data ----------------
const dt = wb.addWorksheet('VA-m2 Data');
title(dt, 'Load Density Data – VA/m² GIA (after diversity)', 'Indicative UK rules of thumb compiled from BSRIA BG 9, CIBSE Guide F/K, BCO Guide and common consultant practice. Verify against current editions.');
header(dt, 4, ['Category', 'Building / space type', 'Min', 'Typical', 'Max', 'Notes']);
const data = [
  ['Offices', 'Office – naturally ventilated, gas heated', 40, 50, 60, 'LED lighting, basic small power'],
  ['Offices', 'Office – mixed mode', 60, 72, 85, ''],
  ['Offices', 'Office – fully air-conditioned', 80, 100, 120, 'Typical Cat A speculative office (VRF / 4-pipe)'],
  ['Offices', 'Office – all-electric (heat pumps)', 100, 120, 140, 'Net-zero / no gas'],
  ['Offices', 'Office – prestige / trading floor', 150, 200, 250, 'Heavy IT and cooling'],
  ['Retail', 'Retail – non-food / shell unit', 60, 80, 100, 'Tenant fit-out dependent'],
  ['Retail', 'Retail – department store', 80, 105, 130, ''],
  ['Retail', 'Retail – supermarket / food', 150, 200, 250, 'Refrigeration dominant'],
  ['Retail', 'Shopping centre – mall common areas', 40, 60, 80, ''],
  ['Hospitality', 'Hotel – budget', 50, 65, 80, 'Alt: 3–5 kVA per bedroom'],
  ['Hospitality', 'Hotel – 4/5 star, air-conditioned', 90, 115, 140, ''],
  ['Hospitality', 'Restaurant – dining area', 80, 115, 150, ''],
  ['Hospitality', 'Commercial kitchen – all electric', 500, 750, 1000, 'Size from equipment schedule'],
  ['Hospitality', 'Restaurant – whole (dining + electric kitchen)', 200, 285, 400, 'Kitchen typically 30–40% of floor area'],
  ['Education', 'School – primary', 25, 35, 45, ''],
  ['Education', 'School – secondary', 40, 52, 65, ''],
  ['Education', 'School – secondary with labs/DT/kitchens', 55, 70, 85, ''],
  ['Education', 'FE college / university teaching', 60, 80, 100, ''],
  ['Education', 'University research laboratories', 120, 185, 250, ''],
  ['Healthcare', 'GP surgery / health centre', 40, 55, 70, ''],
  ['Healthcare', 'Care home', 40, 55, 70, 'Alt: 2–3 kVA per bed'],
  ['Healthcare', 'Hospital – acute', 100, 140, 180, 'HTM 06-01; add standby generation'],
  ['Healthcare', 'Hospital – theatres / imaging / ICU', 150, 200, 250, ''],
  ['Residential', 'Residential – gas heated', 15, 22, 30, 'Prefer ADMD sheet for DNO applications'],
  ['Residential', 'Residential – all-electric (ASHP)', 35, 48, 60, 'Prefer ADMD sheet for DNO applications'],
  ['Residential', 'Student accommodation', 30, 40, 50, 'Alt: 1.5–2.5 kVA per bedroom'],
  ['Residential', 'Residential landlord / common areas', 10, 15, 20, 'Lifts, lighting, ventilation'],
  ['Industrial', 'Warehouse – unheated', 10, 18, 25, ''],
  ['Industrial', 'Warehouse – heated, high-bay LED', 25, 38, 50, ''],
  ['Industrial', 'Distribution centre – automated', 50, 75, 100, ''],
  ['Industrial', 'Chilled / cold store', 80, 140, 200, 'Temperature regime dependent'],
  ['Industrial', 'Light industrial / workshop', 40, 60, 80, ''],
  ['Industrial', 'Heavy industrial / manufacturing', 100, 300, 500, 'Size from process load'],
  ['Leisure / public', 'Sports hall (dry)', 30, 40, 50, ''],
  ['Leisure / public', 'Leisure centre with pool', 80, 115, 150, 'Pool plant and ventilation'],
  ['Leisure / public', 'Library / museum / gallery', 50, 70, 90, 'Higher with environmental control'],
  ['Leisure / public', 'Theatre / cinema', 80, 115, 150, ''],
  ['Leisure / public', 'Stadium – stands / concourse building', 60, 85, 110, 'Excl. floodlights, pitch heating, big screens – add on Specific Loads'],
  ['Leisure / public', 'Church / community hall', 20, 30, 40, ''],
  ['Specialist', 'Data centre – white space', 1000, 2000, 3000, 'IT load × PUE'],
  ['Specialist', 'Comms / server room', 500, 1000, 1500, ''],
  ['Specialist', 'Laboratory – general', 100, 150, 200, 'Fume cupboards and plant'],
  ['Car parks', 'Car park – open / surface', 2, 3, 5, 'Lighting only; EV on Specific Loads sheet'],
  ['Car parks', 'Car park – enclosed, ventilated', 10, 15, 20, 'EV on Specific Loads sheet'],
];
const D0 = 5, D1 = D0 + data.length - 1; // 5..46
data.forEach((row, i) => {
  const r = dt.getRow(D0 + i);
  row.forEach((v, j) => {
    const c = r.getCell(j + 1); c.value = v; c.border = box;
    c.font = { name: 'Arial', size: 10, color: { argb: (j >= 2 && j <= 4) ? BLUE : 'FF000000' } };
    if (j >= 2 && j <= 4) { c.numFmt = '#,##0'; c.alignment = { horizontal: 'center' }; }
  });
});
const srcRow = D1 + 2;
dt.getCell(`A${srcRow}`).value = 'Source: indicative values from BSRIA BG 9/2011 Rules of Thumb, CIBSE Guide F & K, BCO Guide to Specification and UK consultant practice. Values are VA/m² GIA after diversity. User may edit any value (blue).';
dt.getCell(`A${srcRow}`).font = { name: 'Arial', size: 9, italic: true };
// Standard ratings lookups
dt.getCell('H4').value = 'Standard main switchboard / ACB ratings (A)';
dt.getCell('J4').value = 'Standard DNO transformer sizes (kVA)';
['H4', 'J4'].forEach(a => { const c = dt.getCell(a); c.fill = hdrFill; c.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } }; c.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' }; c.border = box; });
const swb = [100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3200, 4000];
const tx = [200, 315, 500, 800, 1000, 1250, 1500, 1600, 2000];
swb.forEach((v, i) => { const c = dt.getCell(`H${5 + i}`); c.value = v; c.border = box; c.font = { name: 'Arial', size: 10, color: { argb: BLUE } }; c.numFmt = '#,##0'; c.alignment = { horizontal: 'center' }; });
tx.forEach((v, i) => { const c = dt.getCell(`J${5 + i}`); c.value = v; c.border = box; c.font = { name: 'Arial', size: 10, color: { argb: BLUE } }; c.numFmt = '#,##0'; c.alignment = { horizontal: 'center' }; });
const SWB = `'VA-m2 Data'!$H$5:$H$${4 + swb.length}`;
const TX = `'VA-m2 Data'!$J$5:$J$${4 + tx.length}`;
dt.getColumn(1).width = 18; dt.getColumn(2).width = 44; [3, 4, 5].forEach(i => dt.getColumn(i).width = 10);
dt.getColumn(6).width = 42; dt.getColumn(7).width = 3; dt.getColumn(8).width = 20; dt.getColumn(9).width = 3; dt.getColumn(10).width = 20;
dt.views = [{ state: 'frozen', ySplit: 4 }];

const TYPES = `'VA-m2 Data'!$B$${D0}:$B$${D1}`;
const DENS = `'VA-m2 Data'!$C$${D0}:$E$${D1}`;
const PF = `'Project Info'!$B$19`;

// ---------------- 3. Load Schedule ----------------
const ls = wb.addWorksheet('Load Schedule');
title(ls, 'Area-Based Load Schedule (VA/m²)', 'Select a building/space type per zone; density is looked up using the basis set on Project Info. Enter an override VA/m² to replace the lookup.');
legend(ls, 3);
header(ls, 5, ['Ref', 'Zone / description', 'Building / space type', 'GIA (m²)', 'Lookup VA/m²', 'Override VA/m²', 'Design VA/m²', 'Demand (kVA)', 'Notes']);
const L0 = 6, L1 = 35;
const examples = [
  ['Z01', 'Ground floor retail units', 'Retail – non-food / shell unit', 850, null, 'Example row – replace'],
  ['Z02', 'Levels 1–4 offices', 'Office – all-electric (heat pumps)', 4200, null, 'Example row – replace'],
  ['Z03', 'Basement car park', 'Car park – enclosed, ventilated', 1500, null, 'EV chargers on Specific Loads'],
  ['Z04', 'Residential landlord areas', 'Residential landlord / common areas', 600, null, 'Flats on Residential ADMD sheet'],
];
for (let r = L0; r <= L1; r++) {
  const ex = examples[r - L0];
  input(ls.getCell(`A${r}`), ex ? ex[0] : null);
  input(ls.getCell(`B${r}`), ex ? ex[1] : null);
  input(ls.getCell(`C${r}`), ex ? ex[2] : null);
  ls.getCell(`C${r}`).dataValidation = { type: 'list', allowBlank: true, formulae: [TYPES], showErrorMessage: true, error: 'Choose a type from the VA-m2 Data list' };
  input(ls.getCell(`D${r}`), ex ? ex[3] : null, '#,##0');
  calc(ls.getCell(`E${r}`), `IF(C${r}="","",INDEX(${DENS},MATCH(C${r},${TYPES},0),MATCH('Project Info'!$B$21,'VA-m2 Data'!$C$4:$E$4,0)))`, '#,##0', GREEN);
  input(ls.getCell(`F${r}`), ex ? ex[4] : null, '#,##0');
  calc(ls.getCell(`G${r}`), `IF(F${r}<>"",F${r},E${r})`, '#,##0');
  calc(ls.getCell(`H${r}`), `IF(OR(D${r}="",G${r}=""),0,D${r}*G${r}/1000)`, '#,##0.0;-#,##0.0;"-"');
  input(ls.getCell(`I${r}`), ex ? ex[5] : null);
}
const LT = L1 + 1;
ls.getCell(`A${LT}`).value = 'TOTAL – area-based loads';
ls.mergeCells(`A${LT}:C${LT}`);
calc(ls.getCell(`D${LT}`), `SUM(D${L0}:D${L1})`, '#,##0');
calc(ls.getCell(`G${LT}`), `IF(D${LT}=0,0,H${LT}*1000/D${LT})`, '#,##0');
calc(ls.getCell(`H${LT}`), `SUM(H${L0}:H${L1})`, '#,##0.0');
ls.getCell(`F${LT}`).value = 'Avg VA/m²:';
['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(c => { const x = ls.getCell(`${c}${LT}`); x.fill = totFill; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); x.border = box; });
ls.getColumn(1).width = 7; ls.getColumn(2).width = 32; ls.getColumn(3).width = 42; ls.getColumn(4).width = 11;
[5, 6, 7].forEach(i => ls.getColumn(i).width = 11); ls.getColumn(8).width = 12; ls.getColumn(9).width = 34;
ls.views = [{ state: 'frozen', ySplit: 5 }];

// ---------------- 4. Residential ADMD ----------------
const ra = wb.addWorksheet('Residential ADMD');
title(ra, 'Residential Load – After Diversity Maximum Demand (ADMD)', 'UK DNOs size dwelling supplies on ADMD per dwelling. Confirm values with the relevant DNO design policy.');
legend(ra, 3);
header(ra, 5, ['Dwelling type', 'ADMD per dwelling (kVA)', 'No. of dwellings', 'Total demand (kVA)', 'Notes']);
const admd = [
  ['Flat / house – gas heating', 1.75, 0, 'Typical 1.5–2.0 kVA'],
  ['Electric storage heating (legacy)', 5, 0, 'Typical 4–6 kVA'],
  ['Air-source heat pump', 4.25, 60, 'Typical 3.5–5 kVA – example quantity'],
  ['Heat pump + managed EV charger', 6.5, 20, 'Typical 5–8 kVA – example quantity'],
  ['Heat pump + unmanaged EV charger', 10, 0, 'Typical 8–12 kVA'],
  ['Communal heat network (per flat, excl. plant)', 2, 0, 'Typical 1.5–2.5 kVA; add plant to Specific Loads'],
  ['Student bedroom (cluster flat)', 2, 0, 'Typical 1.5–2.5 kVA per bedroom'],
  ['', null, null, ''],
  ['', null, null, ''],
];
const R0 = 6, R1 = R0 + admd.length - 1;
admd.forEach((row, i) => {
  const r = R0 + i;
  input(ra.getCell(`A${r}`), row[0] || null);
  input(ra.getCell(`B${r}`), row[1], '0.00');
  input(ra.getCell(`C${r}`), row[2], '#,##0');
  calc(ra.getCell(`D${r}`), `IF(OR(B${r}="",C${r}=""),0,B${r}*C${r})`, '#,##0.0;-#,##0.0;"-"');
  input(ra.getCell(`E${r}`), row[3] || null);
});
const RT = R1 + 1;
ra.getCell(`A${RT}`).value = 'TOTAL – residential ADMD';
calc(ra.getCell(`C${RT}`), `SUM(C${R0}:C${R1})`, '#,##0');
calc(ra.getCell(`D${RT}`), `SUM(D${R0}:D${R1})`, '#,##0.0');
['A', 'B', 'C', 'D', 'E'].forEach(c => { const x = ra.getCell(`${c}${RT}`); x.fill = totFill; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); x.border = box; });
ra.getCell(`A${RT + 2}`).value = 'Note: ADMD values already include diversity across the development. Do not apply further diversity unless the DNO policy allows it for large schemes.';
ra.getCell(`A${RT + 2}`).font = { name: 'Arial', size: 9, italic: true };
ra.getColumn(1).width = 44; ra.getColumn(2).width = 16; ra.getColumn(3).width = 14; ra.getColumn(4).width = 16; ra.getColumn(5).width = 46;

// ---------------- 5. Specific Loads ----------------
const sl = wb.addWorksheet('Specific Loads');
title(sl, 'Specific / Point Loads (EV charging, plant, process)', 'Loads not covered by the area densities. Demand = Qty × rating × diversity; kVA uses the power factor on Project Info.');
legend(sl, 3);
header(sl, 5, ['Item', 'Qty', 'Rating each (kW)', 'Diversity factor', 'Demand (kW)', 'Demand (kVA)', 'Notes']);
const sp = [
  ['EV charger – 7.4 kW AC (load managed)', 20, 7.4, 0.5, 'Example – diversity per load management scheme'],
  ['EV charger – 50 kW DC rapid', 1, 50, 1, 'Example'],
  ['Communal heat pump plant', 0, 0, 1, ''],
  ['Commercial kitchen equipment', 0, 0, 0.6, 'Use equipment schedule'],
  ['Lifts (additional to area density)', 0, 0, 1, ''],
  ['Sprinkler / fire pumps (life safety)', 0, 0, 1, 'Normally excluded from MD if on dedicated supply'],
  ['Stadium floodlights (LED)', 0, 0, 1, 'Broadcast-standard pitches typically 0.5–1.5 MW total'],
  ['Stadium undersoil heating / pitch lighting rigs', 0, 0, 0.8, 'Often on a separate metered supply'],
  ['Stadium big screens / broadcast compound', 0, 0, 1, ''],
  ['', null, null, null, ''], ['', null, null, null, ''], ['', null, null, null, ''],
];
const S0 = 6, S1 = S0 + sp.length - 1;
sp.forEach((row, i) => {
  const r = S0 + i;
  input(sl.getCell(`A${r}`), row[0] || null);
  input(sl.getCell(`B${r}`), row[1], '#,##0');
  input(sl.getCell(`C${r}`), row[2], '#,##0.0');
  input(sl.getCell(`D${r}`), row[3], '0.00');
  calc(sl.getCell(`E${r}`), `IF(OR(B${r}="",C${r}="",D${r}=""),0,B${r}*C${r}*D${r})`, '#,##0.0;-#,##0.0;"-"');
  calc(sl.getCell(`F${r}`), `IF(${PF}=0,0,E${r}/${PF})`, '#,##0.0;-#,##0.0;"-"');
  input(sl.getCell(`G${r}`), row[4] || null);
});
const ST = S1 + 1;
sl.getCell(`A${ST}`).value = 'TOTAL – specific loads';
calc(sl.getCell(`E${ST}`), `SUM(E${S0}:E${S1})`, '#,##0.0');
calc(sl.getCell(`F${ST}`), `SUM(F${S0}:F${S1})`, '#,##0.0');
['A', 'B', 'C', 'D', 'E', 'F', 'G'].forEach(c => { const x = sl.getCell(`${c}${ST}`); x.fill = totFill; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); x.border = box; });
sl.getColumn(1).width = 40; sl.getColumn(2).width = 8; sl.getColumn(3).width = 14; sl.getColumn(4).width = 12; sl.getColumn(5).width = 13; sl.getColumn(6).width = 13; sl.getColumn(7).width = 46;

// ---------------- 6. Office Breakdown (cross-check) ----------------
const ob = wb.addWorksheet('Office Breakdown');
title(ob, 'Office Load Breakdown by Service – Cross-Check', 'Use to sanity-check the office VA/m² chosen on the Load Schedule. NOT added to the Summary total (avoids double counting).');
legend(ob, 3);
ob.getCell('A5').value = 'Office net internal area, NIA (m²)'; ob.getCell('A5').font = { name: 'Arial', size: 10, bold: true };
input(ob.getCell('B5'), 3600, '#,##0');
ob.getCell('C5').value = 'Example – enter the office NIA'; ob.getCell('C5').font = { name: 'Arial', size: 9, italic: true };
header(ob, 7, ['Service', 'W/m² (NIA)', 'Demand (kW)', 'Demand (kVA)', 'Typical range / notes']);
const os = [
  ['General lighting (LED)', 8, '6–10 W/m² (fluorescent 10–15)'],
  ['Small power (desks / IT)', 15, '15–25 design; BCO ~15 W/m² diversified'],
  ['HVAC – cooling, fans, pumps', 40, '30–60 W/m²'],
  ['Heating (heat pumps)', 25, '20–40 W/m²; 0 if gas heated'],
  ['Lifts / escalators', 5, '3–8 W/m²'],
  ['Domestic hot water (electric)', 3, '2–5 W/m²'],
  ['Catering / tea points', 3, '2–5 W/m²'],
  ['Comms / small power to plant', 7, '5–10 W/m²'],
  ['Life safety (fire, emergency lighting)', 2, '1–3 W/m²'],
];
const O0 = 8, O1 = O0 + os.length - 1;
os.forEach((row, i) => {
  const r = O0 + i;
  const a = ob.getCell(`A${r}`); a.value = row[0]; a.border = box;
  input(ob.getCell(`B${r}`), row[1], '#,##0.0');
  calc(ob.getCell(`C${r}`), `B${r}*$B$5/1000`, '#,##0.0');
  calc(ob.getCell(`D${r}`), `IF(${PF}=0,0,C${r}/${PF})`, '#,##0.0');
  const e = ob.getCell(`E${r}`); e.value = row[2]; e.border = box;
});
const OT = O1 + 1;
ob.getCell(`A${OT}`).value = 'TOTAL';
calc(ob.getCell(`B${OT}`), `SUM(B${O0}:B${O1})`, '#,##0.0');
calc(ob.getCell(`C${OT}`), `SUM(C${O0}:C${O1})`, '#,##0.0');
calc(ob.getCell(`D${OT}`), `SUM(D${O0}:D${O1})`, '#,##0.0');
['A', 'B', 'C', 'D', 'E'].forEach(c => { const x = ob.getCell(`${c}${OT}`); x.fill = totFill; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); x.border = box; });
ob.getCell(`A${OT + 1}`).value = 'Equivalent VA/m² (NIA)';
calc(ob.getCell(`B${OT + 1}`), `IF($B$5=0,0,D${OT}*1000/$B$5)`, '#,##0');
ob.getCell(`C${OT + 1}`).value = 'Compare with the office density used on the Load Schedule';
ob.getCell(`C${OT + 1}`).font = { name: 'Arial', size: 9, italic: true };
ob.getColumn(1).width = 40; ob.getColumn(2).width = 13; ob.getColumn(3).width = 13; ob.getColumn(4).width = 13; ob.getColumn(5).width = 44;

// ---------------- 7. Summary ----------------
const sm = wb.addWorksheet('Summary');
title(sm, 'Maximum Demand Summary');
calc(sm.getCell('A2'), `'Project Info'!B6&"  |  Ref: "&'Project Info'!B7&"  |  Rev "&'Project Info'!B11`, null, GREEN);
header(sm, 4, ['Item', 'Value', 'Unit', 'Basis / notes']);
const sumRows = [
  ['Area-based loads (Load Schedule)', `'Load Schedule'!H${LT}`, 'kVA', 'Σ GIA × design VA/m²', GREEN],
  ['Residential dwellings (ADMD)', `'Residential ADMD'!D${RT}`, 'kVA', 'Σ dwellings × ADMD', GREEN],
  ['Specific / point loads', `'Specific Loads'!F${ST}`, 'kVA', 'EV, plant, process', GREEN],
  ['Subtotal – estimated maximum demand', 'SUM(B5:B7)', 'kVA', '', null, true],
  ['Spare / future capacity allowance', `'Project Info'!B20`, '%', 'Set on Project Info', GREEN],
  ['Spare capacity', 'B8*B9', 'kVA', ''],
  ['TOTAL DESIGN MAXIMUM DEMAND', 'B8+B10', 'kVA', 'Use for DNO capacity application', null, true],
  ['Total demand (kW)', `B11*${PF}`, 'kW', 'kVA × power factor'],
  ['Design current per phase', `IF('Project Info'!B17=0,0,B11*1000/(SQRT(3)*'Project Info'!B17))`, 'A', 'I = S ÷ (√3 × V L-L)'],
  ['Recommended main switchboard rating', `IF(B13>MAX(${SWB}),"> "&MAX(${SWB})&" A – split supplies",INDEX(${SWB},COUNTIF(${SWB},"<"&B13)+1))`, 'A', 'Next standard rating ≥ design current'],
  ['Indicative transformer size', `IF(B11>MAX(${TX}),"Multiple transformers / HV",INDEX(${TX},COUNTIF(${TX},"<"&B11)+1))`, 'kVA', 'Next standard size ≥ total MD; DNO to confirm'],
  ['Indicative supply arrangement', `IF(B11<=69,"LV whole-current service (≤100 A 3-phase)",IF(B11<=275,"LV CT-metered supply from network",IF(B11<=1000,"LV supply from dedicated substation (likely)","HV supply / customer-owned substation")))`, '', 'Indicative only – thresholds vary by DNO'],
  ['Total floor area (GIA, area-based zones)', `'Load Schedule'!D${LT}`, 'm²', '', GREEN],
  ['Overall load density (area-based zones)', `'Load Schedule'!G${LT}`, 'VA/m²', '', GREEN],
];
sumRows.forEach(([label, f, unit, note, color, bold], i) => {
  const r = 5 + i;
  const a = sm.getCell(`A${r}`); a.value = label; a.border = box; a.font = { name: 'Arial', size: 10, bold: !!bold };
  calc(sm.getCell(`B${r}`), f, '#,##0.0', color);
  const u = sm.getCell(`C${r}`); u.value = unit; u.border = box; u.font = FONT; u.alignment = { horizontal: 'center' };
  const n = sm.getCell(`D${r}`); n.value = note; n.border = box; n.font = { name: 'Arial', size: 9, italic: true };
  if (bold) ['A', 'B', 'C', 'D'].forEach(c => { sm.getCell(`${c}${r}`).fill = totFill; sm.getCell(`${c}${r}`).font = Object.assign({}, sm.getCell(`${c}${r}`).font, { bold: true }); });
});
sm.getCell('B9').numFmt = '0%';
sm.getCell('B14').numFmt = '#,##0'; sm.getCell('B15').numFmt = '#,##0'; sm.getCell('B17').numFmt = '#,##0'; sm.getCell('B18').numFmt = '#,##0';
sm.getCell('B16').alignment = { wrapText: true };
sm.getCell('A20').value = 'Important: this is an early-stage estimate using rule-of-thumb densities. Replace with a detailed load schedule (actual plant and equipment ratings, BS 7671 / IET diversity) at detailed design stage.';
sm.getCell('A20').font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FFC00000' } };
sm.getColumn(1).width = 42; sm.getColumn(2).width = 40; sm.getColumn(3).width = 8; sm.getColumn(4).width = 44;

// ---------------- 8. References ----------------
const rf = wb.addWorksheet('References & Notes');
title(rf, 'References, Assumptions and Notes');
const refs = [
  ['REFERENCES', true],
  ['BSRIA BG 9/2011 – Rules of Thumb: Guidelines for Building Services (5th ed.)'],
  ['CIBSE Guide F – Energy Efficiency in Buildings'],
  ['CIBSE Guide K – Electricity in Buildings'],
  ['CIBSE TM54 – Evaluating Operational Energy Performance of Buildings at the Design Stage'],
  ["BCO Guide to Specification (2019 / 2023)"],
  ["Spon's Mechanical & Electrical Services Price Book"],
  ['BS 7671:2018+A2 Requirements for Electrical Installations; IET On-Site Guide (Appendix A – diversity); IET Guidance Note 1'],
  ['HTM 06-01 – Electrical Services Supply and Distribution (healthcare)'],
  ['ENA Engineering Recommendation G81 – new connections; relevant DNO ADMD / design policy for residential'],
  [''],
  ['ASSUMPTIONS', true],
  ['1. All VA/m² densities are after diversity and apply to Gross Internal Area (GIA) unless noted.'],
  ['2. VA = W ÷ power factor. Overall power factor is set on Project Info (default 0.95).'],
  ['3. Spare capacity allowance is applied to the total of area, residential and specific loads.'],
  ['4. Residential ADMD values already include development-wide diversity.'],
  ['5. Office Breakdown sheet is a cross-check only and is not included in the Summary total.'],
  ['6. Life-safety loads on dedicated supplies (e.g. sprinkler pumps) may be excluded from maximum demand – confirm with the fire strategy.'],
  ['7. Switchboard, transformer and supply arrangement outputs are indicative; the DNO determines the point of connection and supply arrangement.'],
  ['8. Example values on each input sheet are illustrative and must be replaced with project data.'],
];
refs.forEach(([t, sec], i) => {
  const c = rf.getCell(`A${3 + i}`); c.value = t; c.font = { name: 'Arial', size: 10, bold: !!sec };
  if (sec) c.fill = secFill;
});
rf.getColumn(1).width = 120;

// order: Summary second
wb.worksheets.forEach(ws => { ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 }; });

wb.xlsx.writeFile(OUT).then(() => console.log('written', OUT));
