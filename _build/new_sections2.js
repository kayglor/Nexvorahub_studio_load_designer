// ---------------- 10. Diversity Factors (reference) ----------------
const df = wb.addWorksheet('Diversity Factors');
title(df, 'Standard Diversity, Utilisation and Simultaneity Factors', 'Default factors used by the Client Load Schedule. Edit the blue values to suit the project or the client brief.');
legend(df, 3);
function sub(ws, cell, text) { const c = ws.getCell(cell); c.value = text; c.font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY } }; }
sub(df, 'A5', 'A. Load category factors (applied per item)');
const SETS = ['UK / IEC (BS 7671 – IET OSG)', 'Middle East (DEWA / ADDC / SBC / KAHRAMAA)', 'North America (NEC Art. 220 / CEC Sec. 8)', 'Australia / NZ (AS/NZS 3000 App. C)', 'Africa (SANS 10142 / NRS 034 / BS 7671)', 'South & SE Asia (NBC India / SS 638 / MS IEC)'];
const SET_OVR = {
  1: { 'Air conditioning – split / VRF': [1.0, 1.0], 'Central cooling – chillers & cooling towers': [1.0, 1.0], 'HVAC – chillers / heat pumps': [1.0, 1.0], 'Chilled / condenser water pumps': [0.9, 1.0], 'HVAC – AHUs / fans / pumps': [0.9, 1.0], 'Water heating – instantaneous': [1.0, 0.4], 'Electric heating – direct / panel': [1.0, 0.5], 'Small power – general sockets': [1.0, 0.4] },
  2: { 'General lighting': [1.0, 1.0], 'Small power – general sockets': [1.0, 0.5], 'Small power – dedicated equipment': [1.0, 1.0], 'Commercial kitchen equipment': [1.0, 0.65], 'Motors – general': [1.0, 1.0], 'Lifts': [1.0, 0.9], 'Air conditioning – split / VRF': [1.0, 1.0], 'Central cooling – chillers & cooling towers': [1.0, 1.0], 'HVAC – chillers / heat pumps': [1.0, 1.0], 'Water heating – instantaneous': [1.0, 1.0] },
  3: { 'Small power – general sockets': [1.0, 0.4], 'Commercial kitchen equipment': [1.0, 0.5], 'Air conditioning – split / VRF': [1.0, 0.75] },
  4: { 'Air conditioning – split / VRF': [0.9, 0.9], 'Central cooling – chillers & cooling towers': [0.9, 1.0] },
  5: { 'Air conditioning – split / VRF': [1.0, 0.9], 'Central cooling – chillers & cooling towers': [1.0, 1.0], 'HVAC – chillers / heat pumps': [1.0, 1.0], 'Small power – general sockets': [1.0, 0.4] },
};
const a4 = df.getCell('A4'); a4.value = 'Active factor set (from Project Info)'; a4.font = { name: 'Arial', size: 10, bold: true };
calc(df.getCell('B4'), "'Project Info'!B34", null, GREEN); df.mergeCells('B4:F4');
df.getCell('G4').value = 'Set no.'; df.getCell('G4').font = { name: 'Arial', size: 10, bold: true };
const dfh = (c, r, v) => { const x = df.getCell(r, c); x.value = v; x.fill = hdrFill; x.border = box; x.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFFFF' } }; x.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' }; };
dfh(1, 6, 'Load category'); df.mergeCells(6, 1, 7, 1);
SETS.forEach((n, i) => { dfh(2 + i * 2, 6, n); df.mergeCells(6, 2 + i * 2, 6, 3 + i * 2); dfh(2 + i * 2, 7, 'Demand factor'); dfh(3 + i * 2, 7, 'Diversity factor'); });
dfh(14, 6, 'Active set'); df.mergeCells(6, 14, 6, 15); dfh(14, 7, 'Demand factor'); dfh(15, 7, 'Diversity factor');
dfh(16, 6, 'Project override (input)'); df.mergeCells(6, 16, 6, 17); dfh(16, 7, 'Demand factor'); dfh(17, 7, 'Diversity factor');
dfh(18, 6, 'APPLIED (used by all sheets)'); df.mergeCells(6, 18, 6, 19); dfh(18, 7, 'Demand factor'); dfh(19, 7, 'Diversity factor');
dfh(20, 6, 'Typical range (UK)'); df.mergeCells(6, 20, 7, 20); dfh(21, 6, 'Basis / notes'); df.mergeCells(6, 21, 7, 21);
df.getRow(6).height = 48; df.getRow(7).height = 28;
dfh(23, 7, 'Factor sets');
SETS.forEach((n, i) => { const x = df.getCell(8 + i, 23); x.value = n; x.border = box; x.font = FONT; });
const SET_LIST = `'Diversity Factors'!$W$8:$W$13`;
calc(df.getCell('H4'), `MATCH(B4,$W$8:$W$13,0)`, '0');
const CATS = [
  ['General lighting', 1.0, 0.9, '0.8 – 1.0', 'IET OSG: 66–90% depending on premises; large buildings ~0.9'],
  ['Emergency / life-safety lighting', 1.0, 1.0, '1.0', 'No diversity'],
  ['Small power – general sockets', 1.0, 0.5, '0.2 – 0.75', 'IET OSG: largest + 40–75% of remainder; measured office use often lower'],
  ['Small power – dedicated equipment', 1.0, 0.75, '0.6 – 1.0', ''],
  ['IT / server / comms', 1.0, 1.0, '0.9 – 1.0', 'Continuous load'],
  ['HVAC – chillers / heat pumps', 0.9, 0.9, '0.8 – 1.0', 'Use design-day duty'],
  ['HVAC – AHUs / fans / pumps', 0.8, 0.9, '0.7 – 1.0', 'Ku reflects motor loading / VSD turndown'],
  ['Electric heating – direct / panel', 1.0, 1.0, '0.8 – 1.0', 'IET OSG: thermostatic heating – no diversity'],
  ['Water heating – instantaneous', 1.0, 0.5, '0.25 – 1.0', 'IET OSG: 100% largest + 100% 2nd + 25% remainder'],
  ['Water heating – storage / thermostatic', 1.0, 1.0, '1.0', 'IET OSG: no diversity'],
  ['Commercial kitchen equipment', 1.0, 0.6, '0.5 – 0.7', 'IET OSG: 100% largest + 80% 2nd + 60% remainder'],
  ['Motors – general', 0.75, 0.8, '0.6 – 1.0', 'IET OSG: 100% largest + 80% 2nd + 60% remainder'],
  ['Lifts', 1.0, 0.7, '0.5 – 1.0', 'Group dependent; special consideration per OSG'],
  ['Escalators / travelators', 0.8, 1.0, '0.8 – 1.0', 'Continuous operation in public buildings'],
  ['EV charging – load managed', 1.0, 0.5, '0.3 – 1.0', 'Set by dynamic load management limit'],
  ['EV charging – unmanaged', 1.0, 1.0, '0.8 – 1.0', 'IET Code of Practice for EV Charging'],
  ['Refrigeration / cold rooms', 0.8, 0.9, '0.8 – 1.0', 'Compressor duty cycling'],
  ['Laboratory / process equipment', 0.75, 0.6, '0.4 – 0.8', 'Confirm with user group'],
  ['Medical imaging equipment', 1.0, 0.5, '0.3 – 1.0', 'Use manufacturer momentary / long-term demand'],
  ['Fire-fighting / life-safety pumps', 1.0, 1.0, '0 or 1.0', 'Often excluded from MD where on dedicated supply'],
  ['Air conditioning – split / VRF', 0.9, 0.85, '0.7 – 1.0', 'Outdoor units at design-day duty'],
  ['Central cooling – chillers & cooling towers', 0.9, 0.9, '0.8 – 1.0', 'Standby (N+1) chillers excluded'],
  ['Chilled / condenser water pumps', 0.8, 0.9, '0.7 – 1.0', 'Duty pumps only'],
  ['Domestic water booster pumps', 0.8, 0.5, '0.5 – 1.0', 'Duty / standby sets: 0.5'],
  ['Sewage / sump pumps', 0.8, 0.5, '0.5 – 1.0', 'Duty / standby'],
  ['Fire & security alarm systems / BMS', 1.0, 1.0, '1.0', 'Continuous standby load'],
  ['Smoke extract fans (life safety)', 1.0, 1.0, '0 or 1.0', 'Life-safety; secondary supply'],
  ['External / common-area lighting', 1.0, 1.0, '1.0', 'Switched together – no diversity'],
  ['Other / unclassified', 1.0, 1.0, '–', 'Default'],
];
const C0 = 8, C1 = C0 + CATS.length - 1;
CATS.forEach(([n, ku, dfv, rg, note], i) => {
  const r = C0 + i;
  const a = df.getCell(`A${r}`); a.value = n; a.border = box; a.font = FONT;
  SETS.forEach((_, k) => {
    const o = (SET_OVR[k] || {})[n] || [ku, dfv];
    [o[0], o[1]].forEach((v, j) => { const x = df.getCell(r, 2 + k * 2 + j); input(x, v, '0.00'); x.alignment = { horizontal: 'center' }; });
  });
  calc(df.getCell(r, 14), `INDEX($B${r}:$M${r},2*$H$4-1)`, '0.00');
  calc(df.getCell(r, 15), `INDEX($B${r}:$M${r},2*$H$4)`, '0.00');
  input(df.getCell(r, 16), null, '0.00'); input(df.getCell(r, 17), null, '0.00');
  calc(df.getCell(r, 18), `IF(P${r}<>"",P${r},N${r})`, '0.00');
  calc(df.getCell(r, 19), `IF(Q${r}<>"",Q${r},O${r})`, '0.00');
  [18, 19].forEach(c => { df.getCell(r, c).fill = totFill; df.getCell(r, c).font = { name: 'Arial', size: 10, bold: true }; df.getCell(r, c).alignment = { horizontal: 'center' }; });
  const g = df.getCell(r, 20); g.value = rg; g.border = box; g.font = FONT; g.alignment = { horizontal: 'center' };
  const e = df.getCell(r, 21); e.value = note; e.border = box; e.font = { name: 'Arial', size: 9, italic: true };
});
df.getCell(`A${C1 + 1}`).value = 'Demand factor = maximum demand ÷ connected load of that load type (utilisation). Diversity factor = allowance for items of that type not operating together. Applied demand = connected × demand factor × diversity factor; a simultaneity factor Ks is then applied per board (table B). Regional sets are indicative defaults reflecting each code\'s approach (e.g. NEC 220.44 sockets, 220.56 kitchens, 620.14 lifts) – verify against the adopted code and enter project overrides where required.';
df.getCell(`A${C1 + 1}`).font = { name: 'Arial', size: 9, italic: true }; df.getCell(`A${C1 + 1}`).alignment = { wrapText: true };
df.mergeCells(C1 + 1, 1, C1 + 3, 21);
const DFQ = "'Diversity Factors'";
const CAT_LIST = `${DFQ}!$A$${C0}:$A$${C1}`;
const CAT_KU = `${DFQ}!$R$${C0}:$R$${C1}`;
const CAT_DF = `${DFQ}!$S$${C0}:$S$${C1}`;
for (let c = 2; c <= 19; c++) df.getColumn(c).width = 10;
df.getColumn(20).width = 12; df.getColumn(21).width = 50; df.getColumn(22).width = 3; df.getColumn(23).width = 44;
const K0r = C1 + 3;
sub(df, `A${K0r}`, 'B. Simultaneity factor Ks – distribution boards by number of circuits (IEC 61439 / IEC 60439 assumed loading)');
header(df, K0r + 1, ['Number of circuits (from)', 'Simultaneity factor Ks', 'Applies to', '', '']);
const KS = [[1, 1.0, '1 circuit'], [2, 0.9, '2 – 3 circuits'], [4, 0.8, '4 – 5 circuits'], [6, 0.7, '6 – 9 circuits'], [10, 0.6, '10 circuits and more']];
const KSa = K0r + 2, KSb = KSa + KS.length - 1;
KS.forEach(([n, k, t], i) => {
  const r = KSa + i;
  const a = df.getCell(`A${r}`); a.value = n; a.border = box; a.font = { name: 'Arial', size: 10, color: { argb: BLUE } }; a.alignment = { horizontal: 'center' };
  input(df.getCell(`B${r}`), k, '0.00'); df.getCell(`B${r}`).alignment = { horizontal: 'center' };
  const c = df.getCell(`C${r}`); c.value = t; c.border = box; c.font = FONT;
});
const KS_N = `${DFQ}!$A$${KSa}:$A$${KSb}`, KS_V = `${DFQ}!$B$${KSa}:$B$${KSb}`;

const R0r = KSb + 3;
sub(df, `A${R0r}`, 'C. Simultaneity factor – residential (by number of dwellings, reference only)');
header(df, R0r + 1, ['Number of dwellings (from)', 'Simultaneity factor', 'Applies to', '', '']);
const KR = [[2, 1.0, '2 – 4'], [5, 0.78, '5 – 9'], [10, 0.63, '10 – 14'], [15, 0.53, '15 – 19'], [20, 0.49, '20 – 24'], [25, 0.46, '25 – 29'], [30, 0.44, '30 – 34'], [35, 0.42, '35 – 39'], [40, 0.41, '40 – 49'], [50, 0.40, '50 and more']];
KR.forEach(([n, k, t], i) => {
  const r = R0r + 2 + i;
  const a = df.getCell(`A${r}`); a.value = n; a.border = box; a.font = { name: 'Arial', size: 10, color: { argb: BLUE } }; a.alignment = { horizontal: 'center' };
  input(df.getCell(`B${r}`), k, '0.00'); df.getCell(`B${r}`).alignment = { horizontal: 'center' };
  const c = df.getCell(`C${r}`); c.value = t; c.border = box; c.font = FONT;
});
const KRe = R0r + 2 + KR.length - 1;
df.getCell(`A${KRe + 1}`).value = 'Source: European practice (NF C 14-100, as published in the Schneider Electric Electrical Installation Guide). UK DNOs normally use ADMD instead – see Residential ADMD sheet.';
df.getCell(`A${KRe + 1}`).font = { name: 'Arial', size: 9, italic: true };

const F0r = KRe + 4;
sub(df, `A${F0r}`, 'D. Simultaneity factor by function (reference)');
header(df, F0r + 1, ['Function', 'Simultaneity factor', 'Notes', '', '']);
const KF = [
  ['Lighting', '1.0', ''], ['Heating and air conditioning', '1.0', ''],
  ['Socket-outlets', '0.1 – 0.2', 'Higher in industrial premises'],
  ['Lifts and catering hoists – most powerful motor', '1.0', ''], ['Lifts – second most powerful motor', '0.75', ''], ['Lifts – all other motors', '0.6', ''],
];
KF.forEach(([n, k, t], i) => {
  const r = F0r + 2 + i;
  [n, k, t].forEach((v, j) => { const c = df.getCell(`${'ABC'[j]}${r}`); c.value = v; c.border = box; c.font = FONT; if (j === 1) c.alignment = { horizontal: 'center' }; });
});
df.getCell(`A${F0r + 2 + KF.length}`).value = 'Source: IEC 60364 practice as published in the Schneider Electric Electrical Installation Guide; IET On-Site Guide Appendix A.';
df.getCell(`A${F0r + 2 + KF.length}`).font = { name: 'Arial', size: 9, italic: true };
df.getColumn(1).width = 44;

// ---------------- 11. Client Load Schedule (detailed) ----------------
const cl = wb.addWorksheet('Client Load Schedule');
title(cl, 'Client Load Schedule – Detailed (actual client requirements)', 'Enter the client\'s equipment / circuit data. Demand = connected × Ku × diversity; then a simultaneity factor Ks is applied per distribution board and at the main switchboard.');
legend(cl, 3);
cl.getCell('A5').value = 'Main switchboard simultaneity factor (between DBs)'; cl.getCell('A5').font = { name: 'Arial', size: 10, bold: true };
input(cl.getCell('E5'), 0.9, '0.00');
cl.getCell('F5').value = 'Typical 0.8 – 1.0'; cl.getCell('F5').font = { name: 'Arial', size: 9, italic: true };
cl.getCell('A6').value = 'Leave PF blank to use the project power factor; leave Phases blank for 3-phase. Ku / diversity override replaces the default from the Diversity Factors sheet.';
cl.getCell('A6').font = { name: 'Arial', size: 9, italic: true };
header(cl, 8, ['Ref', 'DB / panel', 'Floor / location', 'Description', 'Load category', 'Qty', 'Rating each (kW)', 'PF', 'Phases (1/3)', 'Connected (kW)', 'Connected (kVA)', 'Demand factor (regional)', 'Demand factor override', 'Diversity factor (regional)', 'Diversity factor override', 'Demand (kW)', 'Demand (kVA)', 'Current per phase (A)', 'Notes']);
cl.getRow(8).height = 42;
const CL0 = 9, CL1 = 108;
const clEx = [
  ['L01', 'DB-GF', 'Ground', 'LED lighting – lobby & retail', 'General lighting', 1, 6.5, null, 3],
  ['L02', 'DB-GF', 'Ground', 'Small power – retail sockets', 'Small power – general sockets', 1, 18, null, 3],
  ['L03', 'DB-GF', 'Ground', 'Restaurant kitchen equipment', 'Commercial kitchen equipment', 1, 85, null, 3],
  ['L04', 'DB-GF', 'Ground', 'Emergency lighting', 'Emergency / life-safety lighting', 1, 1.2, null, 1],
  ['M01', 'MCC-ROOF', 'Roof', 'Air-source heat pumps', 'HVAC – chillers / heat pumps', 3, 45, 0.92, 3],
  ['M02', 'MCC-ROOF', 'Roof', 'AHUs', 'HVAC – AHUs / fans / pumps', 2, 11, 0.9, 3],
  ['M03', 'MCC-ROOF', 'Roof', 'Lift motors', 'Lifts', 3, 15, 0.9, 3],
  ['E01', 'DB-EV', 'Basement', '7.4 kW AC chargers (DLM)', 'EV charging – load managed', 20, 7.4, 0.99, 1],
];
const VLN = `'Project Info'!$B$18`, VLL = `'Project Info'!$B$17`;
for (let r = CL0; r <= CL1; r++) {
  const ex = clEx[r - CL0] || [];
  ['A', 'B', 'C', 'D', 'E'].forEach((c, j) => input(cl.getCell(`${c}${r}`), ex[j] !== undefined ? ex[j] : null));
  cl.getCell(`E${r}`).dataValidation = { type: 'list', allowBlank: true, formulae: [CAT_LIST], showErrorMessage: true, error: 'Choose a category from the Diversity Factors sheet' };
  input(cl.getCell(`F${r}`), ex[5] !== undefined ? ex[5] : null, '#,##0');
  input(cl.getCell(`G${r}`), ex[6] !== undefined ? ex[6] : null, '#,##0.00');
  input(cl.getCell(`H${r}`), ex[7] !== undefined ? ex[7] : null, '0.00');
  input(cl.getCell(`I${r}`), ex[8] !== undefined ? ex[8] : null, '0');
  cl.getCell(`I${r}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['"1,3"'] };
  const pfr = `IF(H${r}="",${PF},H${r})`;
  calc(cl.getCell(`J${r}`), `IF(OR(F${r}="",G${r}=""),0,F${r}*G${r})`, '#,##0.0;-#,##0.0;"-"');
  calc(cl.getCell(`K${r}`), `IF(${pfr}=0,0,J${r}/${pfr})`, '#,##0.0;-#,##0.0;"-"');
  calc(cl.getCell(`L${r}`), `IF(E${r}="","",INDEX(${CAT_KU},MATCH(E${r},${CAT_LIST},0)))`, '0.00', GREEN);
  input(cl.getCell(`M${r}`), null, '0.00');
  calc(cl.getCell(`N${r}`), `IF(E${r}="","",INDEX(${CAT_DF},MATCH(E${r},${CAT_LIST},0)))`, '0.00', GREEN);
  input(cl.getCell(`O${r}`), null, '0.00');
  calc(cl.getCell(`P${r}`), `J${r}*IF(M${r}<>"",M${r},IF(L${r}="",1,L${r}))*IF(O${r}<>"",O${r},IF(N${r}="",1,N${r}))`, '#,##0.0;-#,##0.0;"-"');
  calc(cl.getCell(`Q${r}`), `IF(${pfr}=0,0,P${r}/${pfr})`, '#,##0.0;-#,##0.0;"-"');
  calc(cl.getCell(`R${r}`), `IF(Q${r}=0,0,IF(I${r}=1,Q${r}*1000/${VLN},Q${r}*1000/(SQRT(3)*${VLL})))`, '#,##0.0;-#,##0.0;"-"');
  input(cl.getCell(`S${r}`), null);
}
const CLT = CL1 + 1;
cl.getCell(`A${CLT}`).value = 'TOTAL (before simultaneity)'; cl.mergeCells(`A${CLT}:E${CLT}`);
['J', 'K', 'P', 'Q'].forEach(c => calc(cl.getCell(`${c}${CLT}`), `SUM(${c}${CL0}:${c}${CL1})`, '#,##0.0'));
'ABCDEFGHIJKLMNOPQRS'.split('').forEach(c => { const x = cl.getCell(`${c}${CLT}`); x.fill = totFill; x.border = box; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); });

// DB summary with simultaneity
const DBh = CLT + 3;
sub(cl, `A${DBh - 1}`, 'Distribution board summary – simultaneity factor Ks applied per board');
['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].forEach((c, j) => {
  const x = cl.getCell(`${c}${DBh}`); x.value = ['DB / panel', 'No. of circuits / items', 'Σ demand after Ku × DF (kVA)', 'Ks from table', 'Ks override', 'Applied Ks', 'DB demand (kVA)', 'DB current, 3-ph (A)'][j];
  x.fill = hdrFill; x.border = box; x.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } }; x.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
});
cl.getRow(DBh).height = 42;
const DB0 = DBh + 1, DB1 = DB0 + 19;
const dbEx = ['DB-GF', 'MCC-ROOF', 'DB-EV'];
const RNG = (c) => `$${c}$${CL0}:$${c}$${CL1}`;
for (let r = DB0; r <= DB1; r++) {
  input(cl.getCell(`A${r}`), dbEx[r - DB0] || null);
  calc(cl.getCell(`B${r}`), `IF(A${r}="",0,COUNTIFS(${RNG('B')},A${r},${RNG('J')},">0"))`, '0;-0;"-"');
  calc(cl.getCell(`C${r}`), `IF(A${r}="",0,SUMIF(${RNG('B')},A${r},${RNG('Q')}))`, '#,##0.0;-#,##0.0;"-"');
  calc(cl.getCell(`D${r}`), `IF(B${r}=0,"",INDEX(${KS_V},MATCH(B${r},${KS_N},1)))`, '0.00', GREEN);
  input(cl.getCell(`E${r}`), null, '0.00');
  calc(cl.getCell(`F${r}`), `IF(E${r}<>"",E${r},IF(D${r}="",1,D${r}))`, '0.00');
  calc(cl.getCell(`G${r}`), `C${r}*F${r}`, '#,##0.0;-#,##0.0;"-"');
  calc(cl.getCell(`H${r}`), `IF(${VLL}=0,0,G${r}*1000/(SQRT(3)*${VLL}))`, '#,##0.0;-#,##0.0;"-"');
}
const DBT = DB1 + 1, DBU = DB1 + 2, DBM = DB1 + 3;
cl.getCell(`A${DBT}`).value = 'Σ DB demands';
calc(cl.getCell(`B${DBT}`), `SUM(B${DB0}:B${DB1})`, '0');
calc(cl.getCell(`C${DBT}`), `SUM(C${DB0}:C${DB1})`, '#,##0.0');
calc(cl.getCell(`G${DBT}`), `SUM(G${DB0}:G${DB1})`, '#,##0.0');
cl.getCell(`A${DBU}`).value = 'Check – demand on DBs not listed above (should be 0)';
calc(cl.getCell(`C${DBU}`), `ROUND(Q${CLT}-C${DBT},3)`, '#,##0.0');
cl.getCell(`A${DBM}`).value = 'CLIENT MAXIMUM DEMAND (Σ DB × main SF)';
calc(cl.getCell(`G${DBM}`), `G${DBT}*$E$5`, '#,##0.0');
calc(cl.getCell(`H${DBM}`), `IF(${VLL}=0,0,G${DBM}*1000/(SQRT(3)*${VLL}))`, '#,##0.0');
[DBT, DBM].forEach(r => 'ABCDEFGH'.split('').forEach(c => { const x = cl.getCell(`${c}${r}`); x.fill = totFill; x.border = box; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); }));
cl.getColumn('A').width = 30; cl.getColumn('B').width = 13; cl.getColumn('C').width = 15; cl.getColumn('D').width = 32; cl.getColumn('E').width = 34;
'FGHIJKLMNOPQR'.split('').forEach(c => cl.getColumn(c).width = 11); cl.getColumn('S').width = 30;
cl.views = [{ state: 'frozen', xSplit: 4, ySplit: 8 }];
const CLQ = "'Client Load Schedule'";

// ---------------- 12. Lighting Data ----------------
const ld = wb.addWorksheet('Lighting Data');
title(ld, 'Lighting Requirements by Building and Space Type', 'Maintained illuminance and quality criteria per BS EN 12464-1:2021 / SLL Code for Lighting; LED benchmark power densities are typical UK design values.');
legend(ld, 3);
header(ld, 5, ['Space type', 'Building group', 'Maintained illuminance Em (lux)', 'UGR limit', 'Colour rendering Ra (min)', 'Benchmark LED W/m²', 'Notes']);
ld.getRow(5).height = 42;
const LT_ = [
  ['Office – open plan / writing, reading', 'Commercial office', 500, 19, 80, 7, 'Uniformity Uo ≥ 0.60'],
  ['Office – meeting / conference room', 'Commercial office', 500, 19, 80, 7, 'Provide dimming / scene control'],
  ['Office – reception desk', 'Commercial office', 300, 22, 80, 6, ''],
  ['School – classroom', 'School', 300, 19, 80, 6, '500 lux for evening / adult education'],
  ['School – lecture hall / auditorium', 'School', 500, 19, 80, 8, 'Dimmable for AV'],
  ['School – science laboratory', 'School', 500, 19, 80, 8, ''],
  ['School – library reading area', 'School', 500, 19, 80, 7, 'Bookshelves 200 lux vertical'],
  ['School / leisure – sports hall (training)', 'School', 300, 22, 60, 8, 'BS EN 12193 Class III; competition 500–750 lux'],
  ['Hospital – ward (general)', 'Hospital', 100, 19, 80, 5, 'Reading light 300 lux; night 5 lux'],
  ['Hospital – examination / treatment room', 'Hospital', 1000, 19, 90, 12, 'High colour rendering for clinical assessment'],
  ['Hospital – operating theatre (general)', 'Hospital', 1000, 19, 90, 20, 'Operating cavity by surgical luminaire (excluded)'],
  ['Hospital – corridors (day)', 'Hospital', 200, 22, 80, 4, 'Night setting 50 lux'],
  ['Hotel – reception / lobby', 'Hotel', 300, 22, 80, 8, 'Decorative lighting often adds 30–50%'],
  ['Hotel – guest bedroom', 'Hotel', 150, 22, 80, 6, 'Task lighting at desk / bed 300 lux'],
  ['Restaurant / hotel – dining area', 'Restaurant', 200, 22, 80, 8, 'Ambience-led; dimmable'],
  ['Commercial kitchen', 'Restaurant', 500, 22, 80, 10, 'IP-rated luminaires'],
  ['Retail – sales area', 'Shopping mall', 300, 22, 80, 12, 'Accent / display lighting extra'],
  ['Retail – checkout / till area', 'Shopping mall', 500, 19, 80, 12, ''],
  ['Shopping mall – circulation / atrium', 'Shopping mall', 200, 22, 80, 8, 'Daylight-linked control'],
  ['Stadium – concourse / circulation', 'Stadium', 200, 25, 60, 6, ''],
  ['Stadium – hospitality boxes / lounges', 'Stadium', 300, 22, 80, 8, ''],
  ['Stadium – pitch floodlighting', 'Stadium', 500, 0, 65, 0, 'Competition 500–2,000+ lux; broadcast criteria vertical lux – size on Specific Loads'],
  ['Place of worship – prayer hall / nave', 'Place of worship', 200, 22, 80, 6, '100–300 lux; reading of texts 300 lux'],
  ['Place of worship – ablution area', 'Place of worship', 200, 25, 80, 4, 'IP-rated luminaires'],
  ['Residential – living / bedroom (furnished flat)', 'Residential', 100, 0, 80, 4, 'Kitchen worktops 300 lux'],
  ['Residential – common corridors / stairs', 'Residential', 100, 25, 80, 3, 'Presence detection; emergency lighting'],
  ['Public service – counter / waiting area', 'Public services', 300, 22, 80, 6, ''],
  ['Car park – indoor ramps / entrances (day)', 'Car park', 300, 25, 40, 5, 'Night 75 lux'],
  ['Circulation – corridors', 'General', 100, 28, 40, 3, ''],
  ['Circulation – stairs / lift lobbies', 'General', 150, 25, 40, 3.5, ''],
  ['Toilets / washrooms', 'General', 200, 25, 80, 4, ''],
  ['Plant rooms / switch rooms', 'General', 200, 25, 60, 4, ''],
  ['Store rooms', 'General', 100, 25, 60, 3, ''],
  ['Server / data room', 'General', 500, 19, 80, 8, ''],
  ['Car park – parking areas (internal)', 'Car park', 75, 0, 40, 2, 'Entry / exit ramps 300 lux by day'],
  ['Warehouse – storage aisles', 'Industrial', 150, 25, 60, 4, 'Picking areas 200–300 lux'],
  ['Industrial – medium assembly', 'Industrial', 500, 22, 80, 8, 'Rough assembly 300 lux; fine 750 lux'],
];
const LD0 = 6, LD1 = LD0 + LT_.length - 1;
LT_.forEach((row, i) => {
  const r = LD0 + i;
  row.forEach((v, j) => {
    const c = ld.getCell(`${'ABCDEFG'[j]}${r}`);
    if (j >= 2 && j <= 5) { input(c, v, j === 5 ? '0.0' : '#,##0'); c.alignment = { horizontal: 'center' }; if (j === 3 && v === 0) { c.value = '–'; } }
    else { c.value = v; c.border = box; c.font = j === 6 ? { name: 'Arial', size: 9, italic: true } : FONT; }
  });
});
const EMh = LD1 + 3;
sub(ld, `A${EMh}`, 'Emergency lighting minimum requirements (BS 5266-1 / BS EN 1838)');
header(ld, EMh + 1, ['Area', 'Minimum illuminance', 'Duration', 'Notes', '', '', '']);
const EM = [
  ['Escape route (≤ 2 m wide)', '1 lux on centre line', '3 h typical', '1 h only where premises evacuated immediately and not re-occupied'],
  ['Open area (anti-panic) > 60 m²', '0.5 lux', '3 h typical', 'Excluding 0.5 m border'],
  ['High-risk task area', '10% of normal, min 15 lux', 'As risk assessment', 'Response within 0.5 s'],
  ['Points of emphasis', 'Within 2 m: exits, stairs, changes of level, fire equipment, first aid', '–', ''],
  ['Premises with sleeping accommodation / public assembly', '–', '3 h', 'Hotels, hospitals, care homes, stadia, theatres'],
];
EM.forEach((row, i) => { const r = EMh + 2 + i; row.forEach((v, j) => { const c = ld.getCell(`${'ABCD'[j]}${r}`); c.value = v; c.border = box; c.font = FONT; c.alignment = { wrapText: true, vertical: 'top' }; }); ld.getRow(r).height = 28; });
ld.getColumn('A').width = 44; ld.getColumn('B').width = 22; ld.getColumn('C').width = 16; ld.getColumn('D').width = 34; ld.getColumn('E').width = 14; ld.getColumn('F').width = 14; ld.getColumn('G').width = 60;
const LDQ = "'Lighting Data'";
const L_LIST = `${LDQ}!$A$${LD0}:$A$${LD1}`, L_LUX = `${LDQ}!$C$${LD0}:$C$${LD1}`, L_LPD = `${LDQ}!$F$${LD0}:$F$${LD1}`;

// ---------------- 13. Lighting Design (calculator) ----------------
const lg = wb.addWorksheet('Lighting Design');
title(lg, 'Lighting Design Calculator – Lumen Method', 'N = (E × A) ÷ (Φ × UF × MF). Select a space type to look up the required lux and benchmark W/m²; enter the client\'s chosen luminaire.');
legend(lg, 3);
const lgIn = [[5, 'Maintenance factor MF', 0.8, 'LED, clean environment: 0.8; dirty: 0.6–0.7'], [6, 'Utilisation factor UF (default)', 0.6, 'Typical 0.5–0.7; from manufacturer / DIALux at detailed design'], [7, 'Benchmark tolerance', 0.1, 'Flag if installed W/m² exceeds benchmark by more than this']];
lgIn.forEach(([r, l, v, n]) => {
  const a = lg.getCell(`A${r}`); a.value = l; a.border = box; a.font = FONT; lg.mergeCells(`A${r}:B${r}`);
  input(lg.getCell(`C${r}`), v, r === 7 ? '0%' : '0.00');
  const x = lg.getCell(`D${r}`); x.value = n; x.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } };
});
header(lg, 9, ['Ref', 'Building / zone', 'Space type', 'Area per room (m²)', 'No. of identical rooms / floors', 'Total area (m²)', 'Required lux', 'Lux override', 'Design lux', 'Luminaire output (lm)', 'Luminaire power (W)', 'Luminaires per room', 'Total luminaires', 'Installed load (kW)', 'Installed W/m²', 'Benchmark W/m²', 'Achieved lux', 'Check']);
lg.getRow(9).height = 42;
const LG0 = 10, LG1 = 49;
const lgEx = [
  ['LT01', 'Offices L1–L4', 'Office – open plan / writing, reading', 850, 4, null, 4000, 32],
  ['LT02', 'Offices L1–L4', 'Office – meeting / conference room', 30, 12, null, 3600, 30],
  ['LT03', 'Core', 'Circulation – corridors', 120, 5, null, 1800, 15],
  ['LT04', 'Hotel', 'Hotel – guest bedroom', 22, 120, null, 1500, 14],
  ['LT05', 'Basement', 'Car park – parking areas (internal)', 1200, 1, null, 3500, 25],
];
for (let r = LG0; r <= LG1; r++) {
  const ex = lgEx[r - LG0] || [];
  input(lg.getCell(`A${r}`), ex[0] || null); input(lg.getCell(`B${r}`), ex[1] || null); input(lg.getCell(`C${r}`), ex[2] || null);
  lg.getCell(`C${r}`).dataValidation = { type: 'list', allowBlank: true, formulae: [L_LIST], showErrorMessage: true, error: 'Choose a space type from Lighting Data' };
  input(lg.getCell(`D${r}`), ex[3] !== undefined ? ex[3] : null, '#,##0.0');
  input(lg.getCell(`E${r}`), ex[4] !== undefined ? ex[4] : null, '#,##0');
  calc(lg.getCell(`F${r}`), `IF(OR(D${r}="",E${r}=""),0,D${r}*E${r})`, '#,##0;-#,##0;"-"');
  calc(lg.getCell(`G${r}`), `IF(C${r}="","",INDEX(${L_LUX},MATCH(C${r},${L_LIST},0)))`, '#,##0', GREEN);
  input(lg.getCell(`H${r}`), null, '#,##0');
  calc(lg.getCell(`I${r}`), `IF(H${r}<>"",H${r},G${r})`, '#,##0');
  input(lg.getCell(`J${r}`), ex[6] !== undefined ? ex[6] : null, '#,##0');
  input(lg.getCell(`K${r}`), ex[7] !== undefined ? ex[7] : null, '#,##0.0');
  calc(lg.getCell(`L${r}`), `IF(OR(D${r}="",J${r}="",J${r}=0,I${r}=""),0,ROUNDUP(I${r}*D${r}/(J${r}*$C$6*$C$5),0))`, '#,##0;-#,##0;"-"');
  calc(lg.getCell(`M${r}`), `IF(E${r}="",0,L${r}*E${r})`, '#,##0;-#,##0;"-"');
  calc(lg.getCell(`N${r}`), `IF(K${r}="",0,M${r}*K${r}/1000)`, '#,##0.00;-#,##0.00;"-"');
  calc(lg.getCell(`O${r}`), `IF(F${r}=0,0,N${r}*1000/F${r})`, '#,##0.0;-#,##0.0;"-"');
  calc(lg.getCell(`P${r}`), `IF(C${r}="","",INDEX(${L_LPD},MATCH(C${r},${L_LIST},0)))`, '#,##0.0', GREEN);
  calc(lg.getCell(`Q${r}`), `IF(OR(D${r}="",D${r}=0),0,L${r}*J${r}*$C$6*$C$5/D${r})`, '#,##0;-#,##0;"-"');
  calc(lg.getCell(`R${r}`), `IF(OR(C${r}="",F${r}=0),"",IF(OR(P${r}="",P${r}=0),"No benchmark",IF(O${r}<=P${r}*(1+$C$7),"OK","Above benchmark – review")))`);
}
const LGT = LG1 + 1;
lg.getCell(`A${LGT}`).value = 'TOTAL – lighting'; lg.mergeCells(`A${LGT}:E${LGT}`);
calc(lg.getCell(`F${LGT}`), `SUM(F${LG0}:F${LG1})`, '#,##0');
calc(lg.getCell(`M${LGT}`), `SUM(M${LG0}:M${LG1})`, '#,##0');
calc(lg.getCell(`N${LGT}`), `SUM(N${LG0}:N${LG1})`, '#,##0.00');
calc(lg.getCell(`O${LGT}`), `IF(F${LGT}=0,0,N${LGT}*1000/F${LGT})`, '#,##0.0');
'ABCDEFGHIJKLMNOPQR'.split('').forEach(c => { const x = lg.getCell(`${c}${LGT}`); x.fill = totFill; x.border = box; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); });
lg.getCell(`A${LGT + 2}`).value = 'Lighting demand after diversity (kVA)'; lg.mergeCells(`A${LGT + 2}:E${LGT + 2}`);
calc(lg.getCell(`F${LGT + 2}`), `IF(${PF}=0,0,N${LGT}*INDEX(${CAT_DF},MATCH("General lighting",${CAT_LIST},0))/${PF})`, '#,##0.0');
lg.getCell(`G${LGT + 2}`).value = 'Uses the General lighting diversity factor. Cross-check only – enter in the Client Load Schedule if the lighting is to be counted.';
lg.getCell(`G${LGT + 2}`).font = { name: 'Arial', size: 9, italic: true };
lg.getColumn('A').width = 8; lg.getColumn('B').width = 20; lg.getColumn('C').width = 38;
'DEFGHIJKLMNOPQ'.split('').forEach(c => lg.getColumn(c).width = 11); lg.getColumn('R').width = 24;
lg.views = [{ state: 'frozen', xSplit: 3, ySplit: 9 }];

