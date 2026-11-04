// ---------------- 14. Cable Data (BS 7671 Appendix 4) ----------------
const cd = wb.addWorksheet('Cable Data');
title(cd, 'Cable and Protective Device Data – BS 7671:2018 (18th Edition) Appendix 4', 'Copper conductors. Current-carrying capacity It (A) and voltage drop (mV/A/m) per Tables 4D5, 4D4A and 4E4A; correction factors per Tables 4B1 and 4C1. Verify against your copy of BS 7671 before issue.');
const CSA = [1, 1.5, 2.5, 4, 6, 10, 16, 25, 35, 50, 70, 95, 120, 150, 185, 240, 300];
const NA = null;
const COMBOS = [
  { name: 'Twin & earth 70°C PVC – Method C (clipped direct) [4D5]', temp: 70,
    it1: [16, 20, 27, 37, 47, 64, 85], it3: [], mv1: [44, 29, 18, 11, 7.3, 4.4, 2.8], mv3: [] },
  { name: 'SWA 70°C PVC – Method C (clipped direct) [4D4A]', temp: 70,
    it1: [0, 21, 28, 38, 49, 67, 89, 118, 145, 175, 222, 269, 310, 356, 405, 476, 547],
    it3: [0, 18, 25, 33, 42, 58, 77, 102, 125, 151, 192, 231, 267, 306, 348, 409, 469],
    mv1: [9999, 29, 18, 11, 7.3, 4.4, 2.8, 1.75, 1.25, 0.93, 0.63, 0.47, 0.38, 0.31, 0.26, 0.21, 0.19],
    mv3: [9999, 25, 15, 9.5, 6.4, 3.8, 2.4, 1.5, 1.1, 0.8, 0.55, 0.41, 0.33, 0.29, 0.24, 0.21, 0.185] },
  { name: 'SWA 70°C PVC – Method E (tray / free air) [4D4A]', temp: 70,
    it1: [0, 22, 31, 41, 53, 72, 97, 128, 157, 190, 241, 291, 336, 386, 439, 516, 592],
    it3: [0, 19, 26, 35, 45, 62, 83, 110, 135, 163, 207, 251, 290, 332, 378, 445, 510],
    mv1: [9999, 29, 18, 11, 7.3, 4.4, 2.8, 1.75, 1.25, 0.93, 0.63, 0.47, 0.38, 0.31, 0.26, 0.21, 0.19],
    mv3: [9999, 25, 15, 9.5, 6.4, 3.8, 2.4, 1.5, 1.1, 0.8, 0.55, 0.41, 0.33, 0.29, 0.24, 0.21, 0.185] },
  { name: 'SWA 90°C XLPE – Method C (clipped direct) [4E4A]', temp: 90,
    it1: [0, 27, 36, 49, 62, 85, 110, 146, 180, 219, 279, 338, 392, 451, 515, 607, 698],
    it3: [0, 23, 31, 42, 53, 73, 94, 124, 154, 187, 238, 289, 335, 386, 441, 520, 599],
    mv1: [9999, 31, 19, 12, 7.9, 4.7, 2.9, 1.85, 1.35, 1.0, 0.69, 0.52, 0.42, 0.35, 0.29, 0.24, 0.21],
    mv3: [9999, 27, 16, 10, 6.8, 4.0, 2.5, 1.65, 1.15, 0.87, 0.6, 0.45, 0.37, 0.3, 0.26, 0.21, 0.185] },
  { name: 'SWA 90°C XLPE – Method E (tray / free air) [4E4A]', temp: 90,
    it1: [0, 29, 39, 52, 66, 90, 115, 152, 188, 228, 291, 354, 410, 472, 539, 636, 732],
    it3: [0, 25, 33, 44, 56, 78, 99, 131, 162, 197, 251, 304, 353, 406, 463, 546, 628],
    mv1: [9999, 31, 19, 12, 7.9, 4.7, 2.9, 1.85, 1.35, 1.0, 0.69, 0.52, 0.42, 0.35, 0.29, 0.24, 0.21],
    mv3: [9999, 27, 16, 10, 6.8, 4.0, 2.5, 1.65, 1.15, 0.87, 0.6, 0.45, 0.37, 0.3, 0.26, 0.21, 0.185] },
];
const RCU = [18.1, 12.1, 7.41, 4.61, 3.08, 1.83, 1.15, 0.727, 0.524, 0.387, 0.268, 0.193, 0.153, 0.124, 0.0991, 0.0754, 0.0601];
const R12TE = [36.2, 30.2, 19.51, 16.71, 10.49, 6.44, 4.23];
const CPCTE = [1, 1, 1.5, 1.5, 2.5, 4, 6];
const colL = (n) => { let s = ''; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }; // 0-based → letter
// Header rows 5 (group) and 6 (sub)
const CD0 = 7, CD1 = CD0 + CSA.length - 1; // 7..23
const hdr2 = (c, r, v, fill) => { const x = cd.getCell(r, c); x.value = v; x.fill = fill || hdrFill; x.border = box; x.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFFFF' } }; x.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' }; };
hdr2(1, 5, 'Conductor csa (mm²)'); cd.mergeCells(5, 1, 6, 1);
COMBOS.forEach((cmb, i) => {
  hdr2(2 + i * 2, 5, `It (A): ${cmb.name}`); cd.mergeCells(5, 2 + i * 2, 5, 3 + i * 2);
  hdr2(2 + i * 2, 6, '1-ph (2-core)'); hdr2(3 + i * 2, 6, '3-ph (3/4-core)');
  hdr2(12 + i * 2, 5, `mV/A/m: ${cmb.name}`); cd.mergeCells(5, 12 + i * 2, 5, 13 + i * 2);
  hdr2(12 + i * 2, 6, '1-ph'); hdr2(13 + i * 2, 6, '3-ph');
});
hdr2(22, 5, 'Line conductor r (mΩ/m, 20°C)'); cd.mergeCells(5, 22, 6, 22);
hdr2(23, 5, 'R1+R2 twin & earth (mΩ/m)'); cd.mergeCells(5, 23, 6, 23);
hdr2(24, 5, 'Twin & earth cpc (mm²)'); cd.mergeCells(5, 24, 6, 24);
cd.getRow(5).height = 70; cd.getRow(6).height = 28;
CSA.forEach((s, k) => {
  const r = CD0 + k;
  const a = cd.getCell(r, 1); a.value = s; a.border = box; a.font = { name: 'Arial', size: 10, bold: true }; a.alignment = { horizontal: 'center' };
  COMBOS.forEach((cmb, i) => {
    [[cmb.it1, 2 + i * 2], [cmb.it3, 3 + i * 2], [cmb.mv1, 12 + i * 2], [cmb.mv3, 13 + i * 2]].forEach(([arr, c]) => {
      const x = cd.getCell(r, c); const v = arr[k];
      x.border = box; x.alignment = { horizontal: 'center' };
      if (v !== undefined && v !== null) { input(x, v); x.numFmt = c >= 12 ? (v >= 9999 ? '"–"' : '0.00#') : (v === 0 ? '"–"' : '0'); }
    });
  });
  input(cd.getCell(r, 22), RCU[k], '0.00##');
  if (R12TE[k] !== undefined) { input(cd.getCell(r, 23), R12TE[k], '0.00'); input(cd.getCell(r, 24), CPCTE[k], '0.0'); }
});
cd.getCell(`A${CD1 + 1}`).value = 'Notes: "–" = size not available in that cable type (stored as 0 A / 9999 mV so it is never selected). Twin & earth ratings for Method C only; 3-phase not applicable. For SWA, R1+R2 assumes a cpc of equal csa to the line conductor (armour as cpc must be checked separately).';
cd.getCell(`A${CD1 + 1}`).font = { name: 'Arial', size: 9, italic: true };
const CDQ = "'Cable Data'";
const IT_TAB = `${CDQ}!$B$${CD0}:$K$${CD1}`, MV_TAB = `${CDQ}!$L$${CD0}:$U$${CD1}`;
const CSA_COL = `${CDQ}!$A$${CD0}:$A$${CD1}`, R_COL = `${CDQ}!$V$${CD0}:$V$${CD1}`, R12_COL = `${CDQ}!$W$${CD0}:$W$${CD1}`, CPC_COL = `${CDQ}!$X$${CD0}:$X$${CD1}`;

// Combo list, correction factors and device tables (below)
let rr = CD1 + 4;
const subh = (r, t) => { const c = cd.getCell(`A${r}`); c.value = t; c.font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY } }; };
subh(rr, 'Cable types / installation methods'); rr++;
hdr2(1, rr, 'No.'); hdr2(2, rr, 'Cable type & reference method'); cd.mergeCells(rr, 2, rr, 7); hdr2(8, rr, 'Max conductor temp (°C)');
const CB0 = rr + 1;
COMBOS.forEach((cmb, i) => {
  const r = CB0 + i;
  const a = cd.getCell(r, 1); a.value = i + 1; a.border = box; a.font = FONT; a.alignment = { horizontal: 'center' };
  const b = cd.getCell(r, 2); b.value = cmb.name; b.border = box; b.font = FONT; cd.mergeCells(r, 2, r, 7);
  const t = cd.getCell(r, 8); t.value = cmb.temp; t.border = box; t.font = FONT; t.alignment = { horizontal: 'center' };
});
const CB1 = CB0 + COMBOS.length - 1;
const COMBO_LIST = `${CDQ}!$B$${CB0}:$B$${CB1}`, COMBO_T = `${CDQ}!$H$${CB0}:$H$${CB1}`;

rr = CB1 + 3;
subh(rr, 'Ambient temperature factor Ca (Table 4B1)'); rr++;
hdr2(1, rr, 'Ambient (°C) up to'); hdr2(2, rr, '70°C thermoplastic'); hdr2(3, rr, '90°C thermosetting');
const CA0 = rr + 1;
const CAA = [25, 30, 35, 40, 45, 50, 55, 60], CA70 = [1.03, 1.0, 0.94, 0.87, 0.79, 0.71, 0.61, 0.5], CA90 = [1.02, 1.0, 0.96, 0.91, 0.87, 0.82, 0.76, 0.71];
CAA.forEach((v, i) => { const r = CA0 + i; input(cd.getCell(r, 1), v, '0'); input(cd.getCell(r, 2), CA70[i], '0.00'); input(cd.getCell(r, 3), CA90[i], '0.00'); });
const CA1 = CA0 + CAA.length - 1;
const CA_AMB = `${CDQ}!$A$${CA0}:$A$${CA1}`, CA_TAB = `${CDQ}!$B$${CA0}:$C$${CA1}`;

rr = CA1 + 3;
subh(rr, 'Grouping factor Cg (Table 4C1)'); rr++;
hdr2(1, rr, 'No. of circuits up to'); hdr2(2, rr, 'Bunched / enclosed'); hdr2(3, rr, 'Single layer on wall / tray');
const CG0 = rr + 1;
const CGN = [1, 2, 3, 4, 5, 6, 7, 8, 9, 12, 16, 20], CGB = [1, 0.8, 0.7, 0.65, 0.6, 0.57, 0.54, 0.52, 0.5, 0.45, 0.41, 0.38], CGS = [1, 0.85, 0.79, 0.75, 0.73, 0.72, 0.72, 0.71, 0.7, 0.7, 0.7, 0.7];
CGN.forEach((v, i) => { const r = CG0 + i; input(cd.getCell(r, 1), v, '0'); input(cd.getCell(r, 2), CGB[i], '0.00'); input(cd.getCell(r, 3), CGS[i], '0.00'); });
const CG1 = CG0 + CGN.length - 1;
const CG_N = `${CDQ}!$A$${CG0}:$A$${CG1}`, CG_TAB = `${CDQ}!$B$${CG0}:$C$${CG1}`;

rr = CG1 + 3;
subh(rr, 'Protective device standard ratings (A) and Zs multiplier'); rr++;
const DEV = [
  { n: 'MCB Type B', m: 5, std: 'BS EN 60898-1', r: [6, 10, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125] },
  { n: 'MCB Type C', m: 10, std: 'BS EN 60898-1', r: [6, 10, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125] },
  { n: 'MCB Type D', m: 20, std: 'BS EN 60898-1', r: [6, 10, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125] },
  { n: 'RCBO Type B', m: 5, std: 'BS EN 61009-1', r: [6, 10, 16, 20, 25, 32, 40, 50, 63] },
  { n: 'MCCB', m: 10, std: 'BS EN 60947-2 (Im assumed 10 × In)', r: [16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 200, 250, 320, 400, 500, 630, 800, 1000, 1250, 1600] },
  { n: 'ACB', m: 0, std: 'BS EN 60947-2 (manufacturer data)', r: [630, 800, 1000, 1250, 1600, 2000, 2500, 3200, 4000] },
];
const DVH = rr;
DEV.forEach((d, i) => hdr2(1 + i, DVH, d.n));
const DV0 = DVH + 1, DVN = 21, DV1 = DV0 + DVN - 1;
DEV.forEach((d, i) => d.r.forEach((v, k) => { const x = cd.getCell(DV0 + k, 1 + i); input(x, v, '0'); x.alignment = { horizontal: 'center' }; }));
const DVM = DV1 + 1, DVS = DV1 + 2;
DEV.forEach((d, i) => {
  const m = cd.getCell(DVM, 1 + i); input(m, d.m, '0'); m.alignment = { horizontal: 'center' };
  const s = cd.getCell(DVS, 1 + i); s.value = d.std; s.font = { name: 'Arial', size: 8, italic: true }; s.alignment = { wrapText: true }; s.border = box;
});
cd.getCell(DVM, 8).value = '← magnetic trip multiple used for Zs max (Table 41.3: Zs = 0.95 × U0 ÷ (m × In)); 0 = use manufacturer data';
cd.getCell(DVM, 8).font = { name: 'Arial', size: 9, italic: true };
cd.getRow(DVS).height = 36;
const DEV_NAMES = `${CDQ}!$A$${DVH}:$F$${DVH}`, DEV_TAB = `${CDQ}!$A$${DV0}:$F$${DV1}`, DEV_M = `${CDQ}!$A$${DVM}:$F$${DVM}`;

rr = DVS + 3;
subh(rr, 'Standard final circuits (IET On-Site Guide Appendix H) and BS 7671 additional protection'); rr++;
['Circuit', 'Protective device', 'Cable (Cu, 70°C T&E unless noted)', 'Max floor area / notes'].forEach((t, i) => hdr2(1 + i * 2, rr, t));
[[1, 2], [3, 4], [5, 6], [7, 10]].forEach(([a, b]) => cd.mergeCells(rr, a, rr, b));
const FC = [
  ['Ring final – socket-outlets', '30/32 A MCB / RCBO', '2.5 / 1.5 mm²', '≤ 100 m² floor area; 30 mA RCD (411.3.3)'],
  ['Radial – socket-outlets', '20 A MCB / RCBO', '2.5 / 1.5 mm²', '≤ 50 m² floor area'],
  ['Radial – socket-outlets', '32 A MCB / RCBO', '4.0 / 1.5 mm²', '≤ 75 m² floor area'],
  ['Lighting', '6 A or 10 A Type B', '1.0 or 1.5 mm²', 'Domestic: 30 mA RCD for luminaire circuits (411.3.4)'],
  ['Cooker (≤ 15 kW)', '32 / 40 A', '6 or 10 mm²', 'Apply OSG cooking diversity'],
  ['Electric shower (9.5 – 10.8 kW)', '40 / 45 A', '10 mm²', 'Dedicated circuit; 30 mA RCD in bathroom'],
  ['Immersion heater (3 kW)', '16 A', '2.5 mm²', 'No diversity'],
  ['EV charge point (7.4 kW)', '32 A RCBO Type A + 6 mA DC detection / Type B', '6 mm² (length dependent)', 'Section 722; PME considerations (722.411.4.1)'],
  ['AFDDs', '–', '–', '421.1.7: required in HRRBs, HMOs, PBSA and care homes; recommended elsewhere'],
  ['Surge protection (SPD)', 'Type 2 (Type 1+2 where LPS)', '–', 'Section 443 / 534 risk assessment'],
];
FC.forEach((row, i) => {
  const r = rr + 1 + i;
  row.forEach((v, j) => { const c = cd.getCell(r, 1 + j * 2); c.value = v; c.border = box; c.font = { name: 'Arial', size: 9 }; c.alignment = { wrapText: true, vertical: 'top' }; });
  [[1, 2], [3, 4], [5, 6], [7, 10]].forEach(([a, b]) => cd.mergeCells(r, a, r, b));
  cd.getRow(r).height = 26;
});
cd.getColumn(1).width = 14; for (let c = 2; c <= 21; c++) cd.getColumn(c).width = 11; cd.getColumn(22).width = 12; cd.getColumn(23).width = 12; cd.getColumn(24).width = 12;
cd.views = [{ state: 'frozen', xSplit: 1, ySplit: 6 }];

// ---------------- 15. Cable & Protection (circuit design) ----------------
const cp = wb.addWorksheet('Cable & Protection');
title(cp, 'Circuit Design – Cable Sizing and Protective Device Selection (BS 7671:2018)', 'Ib ≤ In ≤ Iz (433.1.1); It ≥ In ÷ (Ca × Cg × Ci); voltage drop per Appendix 4 / 525; Zs ≤ Table 41.3 limit. Selects the smallest cable meeting current and voltage-drop limits.');
legend(cp, 3);
cp.getCell('A5').value = 'External earth fault loop impedance Ze (Ω)'; cp.getCell('A5').font = { name: 'Arial', size: 10, bold: true };
input(cp.getCell('D5'), 0.35, '0.00');
cp.getCell('E5').value = 'Typical max: TN-C-S 0.35 Ω, TN-S 0.8 Ω, TT ~21 Ω+ (RCD required). Use Zdb column for circuits fed from sub-boards.';
cp.getCell('E5').font = { name: 'Arial', size: 9, italic: true };
calc(cp.getCell('A6'), `IF('Project Info'!$B$36<>"IEC","WARNING: cable tables are BS 7671 / IEC values – use the local code tables (NEC 310.16 / CEC / AS/NZS 3008) for this region","Cable data per BS 7671 / IEC 60364-5-52 – valid for the selected region ("&'Project Info'!$B$22&")")`);
cp.getCell('A6').font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFC00000' } };
const CPH = ['Ref', 'From (board)', 'To / description', 'Circuit type', 'Load (kW)', 'PF', 'Phases (1/3)', 'Ib calc (A)', 'Ib override (A)', 'Design Ib (A)', 'Device type', 'In selected (A)', 'In override (A)', 'Applied In (A)', 'Cable type & method', 'Length (m)', 'Ambient (°C)', 'Circuits grouped', 'Grouping arrangement', 'Thermal insulation Ci', 'Zdb at source (Ω)', 'Ca', 'Cg', 'Required It (A)', 'VD limit (%)', 'Max mV/A/m', 'Selected csa (mm²)', 'Tabulated It (A)', 'Iz (A)', 'Voltage drop (%)', 'Zs calc (Ω)', 'Zs max (Ω)', 'PFC at board (kA)', 'Device Icn/Icu (kA)', 'Poles', 'RCD type', 'RCD IΔn (mA)', 'Earth fault current (A)', 'cpc size – Table 54.7 (mm²)', 'Min cpc – adiabatic (mm²)', 'Selected cpc (mm²)', 'Check', 'Cable & device specification', 'h: combo', 'h: col', 'h: idx', 'h: n'];
header(cp, 8, CPH);
cp.getRow(8).height = 45;
const K = {}; CPH.forEach((h, i) => K[h] = colL(i));
const CP0 = 9, CP1 = 68;
const cpEx = [
  ['C01', 'DB-GF', 'Lighting – lobby & retail', 'Lighting', 1.2, null, 1, 'MCB Type B', 'Twin & earth 70°C PVC – Method C (clipped direct) [4D5]', 25, 30, 2, 'Bunched / enclosed', 1, 0.45, null, 10],
  ['C02', 'DB-GF', 'Radial socket circuit – retail', 'Other', 4, null, 1, 'RCBO Type B', 'Twin & earth 70°C PVC – Method C (clipped direct) [4D5]', 30, 30, 2, 'Bunched / enclosed', 1, 0.45, null, 10],
  ['S01', 'MSB', 'Sub-main to DB-GF', 'Other', 120, null, 3, 'MCCB', 'SWA 90°C XLPE – Method E (tray / free air) [4E4A]', 40, 30, 3, 'Single layer', 1, 0.02, 25, 36],
  ['S02', 'MSB', 'Sub-main to MCC-ROOF', 'Other', 160, null, 3, 'MCCB', 'SWA 90°C XLPE – Method E (tray / free air) [4E4A]', 60, 35, 3, 'Single layer', 1, 0.02, 25, 36],
  ['M01', 'MCC-ROOF', 'ASHP-01', 'Other', 45, 0.92, 3, 'MCCB', 'SWA 90°C XLPE – Method C (clipped direct) [4E4A]', 15, 35, 1, 'Single layer', 1, 0.2, 10, 25],
  ['E01', 'DB-EV', 'EV charger 7.4 kW', 'Other', 7.4, 0.99, 1, 'RCBO Type B', 'SWA 90°C XLPE – Method C (clipped direct) [4E4A]', 30, 30, 4, 'Single layer', 1, 0.3, 6, 10],
];
for (let r = CP0; r <= CP1; r++) {
  const ex = cpEx[r - CP0] || [];
  const c = (k) => cp.getCell(`${K[k]}${r}`);
  const R = (k) => `${K[k]}${r}`;
  input(c('Ref'), ex[0] || null); input(c('From (board)'), ex[1] || null); input(c('To / description'), ex[2] || null);
  input(c('Circuit type'), ex[3] || null); c('Circuit type').dataValidation = { type: 'list', allowBlank: true, formulae: ['"Lighting,Other"'] };
  input(c('Load (kW)'), ex[4] !== undefined ? ex[4] : null, '#,##0.00');
  input(c('PF'), ex[5] !== undefined ? ex[5] : null, '0.00');
  input(c('Phases (1/3)'), ex[6] !== undefined ? ex[6] : null, '0'); c('Phases (1/3)').dataValidation = { type: 'list', allowBlank: true, formulae: ['"1,3"'] };
  const pfr = `IF(${R('PF')}="",${PF},${R('PF')})`;
  const vsys = `IF(${R('Phases (1/3)')}=1,${VLN},${VLL})`;
  calc(c('Ib calc (A)'), `IF(OR(${R('Load (kW)')}="",${pfr}=0),0,IF(${R('Phases (1/3)')}=1,${R('Load (kW)')}*1000/(${VLN}*${pfr}),${R('Load (kW)')}*1000/(SQRT(3)*${VLL}*${pfr})))`, '#,##0.0;-#,##0.0;"-"');
  input(c('Ib override (A)'), null, '#,##0.0');
  calc(c('Design Ib (A)'), `IF(${R('Ib override (A)')}<>"",${R('Ib override (A)')},${R('Ib calc (A)')})`, '#,##0.0;-#,##0.0;"-"');
  input(c('Device type'), ex[7] || null); c('Device type').dataValidation = { type: 'list', allowBlank: true, formulae: [DEV_NAMES] };
  const dcol = `INDEX(${DEV_TAB},0,MATCH(${R('Device type')},${DEV_NAMES},0))`;
  calc(c('In selected (A)'), `IF(OR(${R('Device type')}="",${R('Design Ib (A)')}=0),"",IF(COUNTIF(${dcol},"<"&${R('Design Ib (A)')})+1>COUNT(${dcol}),"Exceeds range",INDEX(${dcol},COUNTIF(${dcol},"<"&${R('Design Ib (A)')})+1)))`, '0');
  input(c('In override (A)'), null, '0');
  calc(c('Applied In (A)'), `IF(${R('In override (A)')}<>"",${R('In override (A)')},${R('In selected (A)')})`, '0');
  input(c('Cable type & method'), ex[8] || null); c('Cable type & method').dataValidation = { type: 'list', allowBlank: true, formulae: [COMBO_LIST] };
  input(c('Length (m)'), ex[9] !== undefined ? ex[9] : null, '#,##0');
  input(c('Ambient (°C)'), ex[10] !== undefined ? ex[10] : null, '0');
  input(c('Circuits grouped'), ex[11] !== undefined ? ex[11] : null, '0');
  input(c('Grouping arrangement'), ex[12] || null); c('Grouping arrangement').dataValidation = { type: 'list', allowBlank: true, formulae: ['"Bunched / enclosed,Single layer"'] };
  input(c('Thermal insulation Ci'), ex[13] !== undefined ? ex[13] : null, '0.00');
  input(c('Zdb at source (Ω)'), ex[14] !== undefined ? ex[14] : null, '0.00');
  // helpers
  calc(c('h: combo'), `IF(${R('Cable type & method')}="","",MATCH(${R('Cable type & method')},${COMBO_LIST},0))`, '0');
  calc(c('h: col'), `IF(${R('h: combo')}="","",(${R('h: combo')}-1)*2+IF(${R('Phases (1/3)')}=1,1,2))`, '0');
  const temp = `INDEX(${COMBO_T},${R('h: combo')})`;
  calc(c('Ca'), `IF(${R('h: combo')}="","",INDEX(${CA_TAB},MIN(COUNTIF(${CA_AMB},"<"&IF(${R('Ambient (°C)')}="",30,${R('Ambient (°C)')}))+1,ROWS(${CA_AMB})),IF(${temp}=70,1,2)))`, '0.00');
  calc(c('Cg'), `IF(${R('h: combo')}="","",INDEX(${CG_TAB},MIN(COUNTIF(${CG_N},"<"&MAX(1,N(${R('Circuits grouped')})))+1,ROWS(${CG_N})),IF(${R('Grouping arrangement')}="Single layer",2,1)))`, '0.00');
  const ci = `IF(${R('Thermal insulation Ci')}="",1,${R('Thermal insulation Ci')})`;
  calc(c('Required It (A)'), `IF(OR(${R('h: combo')}="",NOT(ISNUMBER(${R('Applied In (A)')}))),"",${R('Applied In (A)')}/(${R('Ca')}*${R('Cg')}*${ci}))`, '#,##0.0');
  calc(c('VD limit (%)'), `IF(${R('Circuit type')}="Lighting",${VDL},${VDO})`, '0.0');
  calc(c('Max mV/A/m'), `IF(OR(${R('Design Ib (A)')}=0,N(${R('Length (m)')})=0),9999,${R('VD limit (%)')}/100*${vsys}*1000/(${R('Design Ib (A)')}*${R('Length (m)')}))`, '#,##0.00');
  const itc = `INDEX(${IT_TAB},0,${R('h: col')})`, mvc = `INDEX(${MV_TAB},0,${R('h: col')})`;
  calc(c('h: n'), `IF(${R('h: col')}="","",COUNT(${itc}))`, '0');
  calc(c('h: idx'), `IF(OR(${R('h: col')}="",${R('Required It (A)')}=""),"",MAX(COUNTIF(${itc},"<"&${R('Required It (A)')})+1,COUNTIF(${mvc},">"&${R('Max mV/A/m')})+1))`, '0');
  calc(c('Selected csa (mm²)'), `IF(${R('h: idx')}="","",IF(${R('h: n')}=0,"N/A – 3-ph",IF(${R('h: idx')}>${R('h: n')},"Exceeds – parallel cables",INDEX(${CSA_COL},${R('h: idx')}))))`, '0.0#');
  const ok = `ISNUMBER(${R('Selected csa (mm²)')})`;
  calc(c('Tabulated It (A)'), `IF(${ok},INDEX(${IT_TAB},${R('h: idx')},${R('h: col')}),"")`, '0');
  calc(c('Iz (A)'), `IF(${ok},${R('Tabulated It (A)')}*${R('Ca')}*${R('Cg')}*${ci},"")`, '#,##0.0');
  calc(c('Voltage drop (%)'), `IF(${ok},INDEX(${MV_TAB},${R('h: idx')},${R('h: col')})*${R('Design Ib (A)')}*N(${R('Length (m)')})/1000/${vsys}*100,"")`, '0.00');
  const r12 = `IF(${R('h: combo')}=1,INDEX(${R12_COL},${R('h: idx')}),INDEX(${R_COL},${R('h: idx')})+INDEX(${R_COL},MATCH(${R('cpc size – Table 54.7 (mm²)')},${CSA_COL},0)))`;
  calc(c('Zs calc (Ω)'), `IF(${ok},IF(${R('Zdb at source (Ω)')}="",$D$5,${R('Zdb at source (Ω)')})+${r12}*N(${R('Length (m)')})*IF(${temp}=70,1.2,1.28)/1000,"")`, '0.000');
  const mdev = `INDEX(${DEV_M},1,MATCH(${R('Device type')},${DEV_NAMES},0))`;
  calc(c('Zs max (Ω)'), `IF(OR(${R('Device type')}="",NOT(ISNUMBER(${R('Applied In (A)')}))),"",IF(${mdev}=0,"Manufacturer",0.95*${VLN}/(${mdev}*${R('Applied In (A)')})))`, '0.000');
  input(c('PFC at board (kA)'), ex[15] !== undefined ? ex[15] : null, '0.0');
  input(c('Device Icn/Icu (kA)'), ex[16] !== undefined ? ex[16] : null, '0.0');
  const xx = [['SP', 'None', null], ['SP&N', 'Type A', 30], ['TP&N', 'None', null], ['TP&N', 'None', null], ['TP', 'None', null], ['SP&N', 'Type A + RDC-DD (6 mA DC)', 30]][r - CP0] || [];
  input(c('Poles'), xx[0] || null); c('Poles').dataValidation = { type: 'list', allowBlank: true, formulae: ['"SP,SP&N,DP,TP,TP&N,4P"'] };
  input(c('RCD type'), xx[1] || null); c('RCD type').dataValidation = { type: 'list', allowBlank: true, formulae: ['"None,Type AC,Type A,Type F,Type B,Type A + RDC-DD (6 mA DC)"'] };
  input(c('RCD IΔn (mA)'), xx[2] || null, '0'); c('RCD IΔn (mA)').dataValidation = { type: 'list', allowBlank: true, formulae: ['"10,30,100,300,500"'] };
  calc(c('Earth fault current (A)'), `IF(${ok},0.95*${VLN}/${R('Zs calc (Ω)')},"")`, '#,##0');
  calc(c('cpc size – Table 54.7 (mm²)'), `IF(${ok},INDEX(${CSA_COL},COUNTIF(${CSA_COL},"<"&IF(${R('Selected csa (mm²)')}<=16,${R('Selected csa (mm²)')},IF(${R('Selected csa (mm²)')}<=35,16,${R('Selected csa (mm²)')}/2)))+1),"")`, '0.0#');
  const tdis = `IF(LEFT(${R('Device type')},3)="ACB",1,IF(LEFT(${R('Device type')},4)="MCCB",0.4,0.1))`;
  calc(c('Min cpc – adiabatic (mm²)'), `IF(${ok},SQRT(${R('Earth fault current (A)')}^2*${tdis})/IF(${temp}=70,115,143),"")`, '0.00');
  calc(c('Selected cpc (mm²)'), `IF(${ok},IF(${R('h: combo')}=1,INDEX(${CPC_COL},${R('h: idx')}),INDEX(${CSA_COL},MIN(ROWS(${CSA_COL}),COUNTIF(${CSA_COL},"<"&MAX(${R('cpc size – Table 54.7 (mm²)')},${R('Min cpc – adiabatic (mm²)')}))+1))),"")`, '0.0#');
  calc(c('Check'), `IF(${R('Selected csa (mm²)')}="","",IF(NOT(${ok}),"REVIEW – cable size",IF(${R('Iz (A)')}<${R('Applied In (A)')},"FAIL – Iz < In",IF(${R('Voltage drop (%)')}>${R('VD limit (%)')},"FAIL – voltage drop",IF(AND(ISNUMBER(${R('Zs max (Ω)')}),${R('Zs calc (Ω)')}>${R('Zs max (Ω)')},N(${R('RCD IΔn (mA)')})=0),"FAIL – Zs (add RCD / upsize cpc)",IF(AND(${R('h: combo')}=1,${R('Selected cpc (mm²)')}<${R('Min cpc – adiabatic (mm²)')}),"FAIL – cpc adiabatic (543.1.3)",IF(AND(${R('PFC at board (kA)')}<>"",${R('Device Icn/Icu (kA)')}<>"",N(${R('Device Icn/Icu (kA)')})<N(${R('PFC at board (kA)')})),"FAIL – breaking capacity",IF(AND(ISNUMBER(${R('Zs max (Ω)')}),${R('Zs calc (Ω)')}>${R('Zs max (Ω)')}),"PASS – ADS by RCD (Zs ≤ 50/IΔn)","PASS"))))))))`);
  calc(c('Cable & device specification'), `IF(${ok},IF(${R('h: combo')}=1,"Twin & earth "&${R('Selected csa (mm²)')}&"/"&${R('Selected cpc (mm²)')}&" mm² Cu PVC/PVC (6242Y)",IF(${R('Phases (1/3)')}=1,"3","4")&"-core "&${R('Selected csa (mm²)')}&" mm² Cu "&IF(${temp}=70,"PVC/SWA/PVC (BS 6346)","XLPE/SWA/LSZH (BS 6724)")&" + "&${R('Selected cpc (mm²)')}&" mm² cpc")&" | "&${R('Device type')}&" "&${R('Applied In (A)')}&" A "&IF(${R('Poles')}<>"",${R('Poles')},IF(${R('Phases (1/3)')}=1,"SP","TP"))&IF(N(${R('Device Icn/Icu (kA)')})>0,", "&${R('Device Icn/Icu (kA)')}&" kA","")&IF(AND(${R('RCD type')}<>"",${R('RCD type')}<>"None")," | RCD "&${R('RCD type')}&" "&${R('RCD IΔn (mA)')}&" mA",""),"")`);
}
const CPT = CP1 + 1;
cp.getCell(`A${CPT + 2}`).value = 'Earthing: cpc per Table 54.7 (S ≤ 16: S; 16 < S ≤ 35: 16; S > 35: S/2) and adiabatic S = √(I²t) ÷ k (543.1.3), k = 115 (70°C PVC) / 143 (90°C XLPE), t = 0.1 s MCB/RCBO, 0.4 s MCCB, 1 s ACB. Zs uses line + Table 54.7 cpc resistance. Where Zs exceeds the device limit, an RCD may provide ADS (411.4.204, Zs ≤ 50 V ÷ IΔn).';
cp.getCell(`A${CPT + 2}`).font = { name: 'Arial', size: 9, italic: true };
cp.getCell(`A${CPT + 1}`).value = 'Notes: Design current from kW uses the project PF where PF is blank. Ca/Cg rounded to the next tabulated value (conservative). Zs uses R1+R2 at 20°C × 1.20 (70°C) or × 1.28 (90°C). Selected sizes must be confirmed for fault-current withstand (adiabatic, 543.1.3), cpc sizing and manufacturer data.';
cp.getCell(`A${CPT + 1}`).font = { name: 'Arial', size: 9, italic: true };
const widths = { 'RCD type': 16, 'cpc size – Table 54.7 (mm²)': 12, 'Min cpc – adiabatic (mm²)': 12, 'Selected cpc (mm²)': 11, 'Ref': 7, 'From (board)': 11, 'To / description': 28, 'Circuit type': 9, 'Cable type & method': 40, 'Grouping arrangement': 16, 'Check': 26, 'Cable & device specification': 62, 'Device type': 13 };
CPH.forEach((h, i) => cp.getColumn(i + 1).width = widths[h] || 10);
['h: combo', 'h: col', 'h: idx', 'h: n'].forEach(h => { cp.getColumn(K[h]).hidden = true; });
cp.views = [{ state: 'frozen', xSplit: 3, ySplit: 8 }];

// ---------------- 16. Transformer & Panels ----------------
const tp = wb.addWorksheet('Transformer & Panels');
title(tp, 'Equipment Sizing – Transformer, Main LV Switchboard, Distribution Boards, Standby Generation', 'Driven from the Summary maximum demand and the Client Load Schedule DB summary. Edit blue inputs to suit the client specification.');
legend(tp, 3);
const tsec = (r, t) => { const c = tp.getCell(`A${r}`); c.value = t; c.font = { name: 'Arial', size: 10, bold: true }; ['A', 'B', 'C'].forEach(x => tp.getCell(`${x}${r}`).fill = secFill); };
const trow = (r, label, kind, v, fmt, note, color) => {
  const a = tp.getCell(`A${r}`); a.value = label; a.border = box; a.font = FONT;
  if (kind === 'in') input(tp.getCell(`B${r}`), v, fmt); else calc(tp.getCell(`B${r}`), v, fmt, color);
  tp.getCell(`B${r}`).alignment = { wrapText: true };
  if (note) { const n = tp.getCell(`C${r}`); n.value = note; n.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } }; }
};
const KA = [6, 10, 16, 25, 36, 50, 65, 80, 100];
KA.forEach((v, i) => { const x = tp.getCell(`H${6 + i}`); input(x, v, '0'); });
tp.getCell('H5').value = 'Std kA ratings'; tp.getCell('H5').font = { name: 'Arial', size: 9, bold: true };
const KAR = `'Transformer & Panels'!$H$6:$H$${5 + KA.length}`;
const GEN = [20, 30, 45, 60, 80, 100, 150, 200, 250, 300, 400, 500, 630, 800, 1000, 1250, 1500, 2000, 2500];
GEN.forEach((v, i) => input(tp.getCell(`I${6 + i}`), v, '0'));
tp.getCell('I5').value = 'Std genset kVA'; tp.getCell('I5').font = { name: 'Arial', size: 9, bold: true };
const GENR = `'Transformer & Panels'!$I$6:$I$${5 + GEN.length}`;
const WAYS = [4, 6, 8, 12, 16, 18, 24, 30, 36, 48];
WAYS.forEach((v, i) => input(tp.getCell(`J${6 + i}`), v, '0'));
tp.getCell('J5').value = 'Std DB ways'; tp.getCell('J5').font = { name: 'Arial', size: 9, bold: true };
const WAYR = `'Transformer & Panels'!$J$6:$J$${5 + WAYS.length}`;
const MCCBR = `${CDQ}!$E$${DV0}:$E$${DV1}`, ACBR = `${CDQ}!$F$${DV0}:$F$${DV1}`;
const nextStd = (rng, v) => `IF(${v}>MAX(${rng}),"> "&MAX(${rng}),INDEX(${rng},COUNTIF(${rng},"<"&${v})+1))`;

tsec(5, 'A. TRANSFORMER');
trow(6, 'Design maximum demand incl. spare (kVA)', 'f', `Summary!C18`, '#,##0.0', 'From Summary', GREEN);
trow(7, 'Target transformer loading', 'in', 0.8, '0%', 'Typical 75–85% for efficiency and future growth');
trow(8, 'Number of transformers', 'in', 2, '0', 'Increase if the selected rating shows "> 2000"');
trow(9, 'Resilience', 'in', 'N', null, 'N = share load; N+1 = full load with one unit out of service');
tp.getCell('B9').dataValidation = { type: 'list', allowBlank: false, formulae: ['"N,N+1"'] };
trow(10, 'Required rating per transformer (kVA)', 'f', `IF(B9="N+1",B6/B7/MAX(B8-1,1),B6/B7/MAX(B8,1))`, '#,##0', '');
trow(11, 'Selected standard rating (kVA)', 'f', nextStd(TX, 'B10'), '#,##0', 'Next standard size; utility to confirm');
trow(12, 'Impedance voltage Uk (%)', 'in', 6, '0.00', 'Typical 4–4.75% (≤ 630 kVA), 6% (800–2,500 kVA)');
trow(13, 'Secondary voltage (V)', 'f', `'Project Info'!B17`, '0', 'From Project Info', GREEN);
trow(14, 'Full-load secondary current (A)', 'f', `IF(ISNUMBER(B11),B11*1000/(SQRT(3)*B13),"")`, '#,##0', 'I = S ÷ (√3 × V)');
trow(15, 'LV fault level at terminals – single unit (kA)', 'f', `IF(ISNUMBER(B14),B14*100/B12/1000,"")`, '#,##0.0', 'Infinite-source estimate: Isc = IFL × 100 ÷ Uk%');
trow(16, 'LV fault level with units in parallel (kA)', 'f', `IF(ISNUMBER(B15),B15*IF(B9="N+1",MAX(B8-1,1),B8),"")`, '#,##0.0', 'Use for switchboard rating if transformers operate in parallel');
trow(17, 'Transformer type', 'in', 'Cast resin (dry type)', null, 'Cast resin indoors; oil-immersed (ONAN) for outdoor / utility substations');
tp.getCell('B17').dataValidation = { type: 'list', allowBlank: false, formulae: ['"Cast resin (dry type),Oil-immersed (ONAN),Utility-owned (DNO)"'] };
trow(18, 'Vector group / tapping', 'in', 'Dyn11, ±2 × 2.5% off-load taps', null, '');
trow(19, 'Standards', 'in', 'BS EN / IEC 60076; EU Ecodesign Tier 2 losses', null, '');

tsec(21, 'B. MAIN LV SWITCHBOARD');
trow(22, 'Incomer / busbar rating (A)', 'f', `IF(ISNUMBER(B14),${nextStd(ACBR, 'B14')},"")`, '#,##0', 'Sized to transformer full-load current');
trow(23, 'Incomer device', 'f', `IF(NOT(ISNUMBER(B22)),"",IF(B22>=630,"ACB, 4-pole, draw-out, electronic LSIG trip","MCCB, 4-pole, electronic trip"))`, null, 'BS EN 60947-2');
trow(24, 'Required short-circuit rating Icw 1 s (kA)', 'f', `IF(ISNUMBER(B16),${nextStd(KAR, 'B16')},"")`, '0', 'Next standard rating ≥ prospective fault level');
trow(25, 'Form of separation', 'in', 'Form 4 Type 2', null, 'Form 2/3 for smaller boards; Form 4 for critical buildings');
tp.getCell('B25').dataValidation = { type: 'list', allowBlank: false, formulae: ['"Form 1,Form 2b,Form 3b,Form 4 Type 2,Form 4 Type 6,Form 4 Type 7"'] };
trow(26, 'IP rating', 'in', 'IP31 (indoor)', null, 'IP54+ for plant rooms with sprinklers / outdoors');
trow(27, 'Assembly standard', 'f', `'Project Info'!B30`, null, 'From region', GREEN);
trow(28, 'Spare ways / capacity allowance', 'in', 0.25, '0%', 'Typical 20–25% spare equipped + space');
trow(29, 'Metering', 'in', 'Utility tariff meter + MID class 1 sub-meters on outgoing ways', null, 'Building Regulations Part L sub-metering');
trow(30, 'Power factor correction', 'in', 'Not required (PF ≥ 0.95) – review at detailed design', null, '');
trow(31, 'Surge protection', 'in', 'Type 1+2 SPD at origin', null, 'BS 7671 Section 443 / 534');

tsec(33, 'C. DISTRIBUTION BOARDS / PANELS (from Client Load Schedule)');
const TPH = ['DB / panel', 'Circuits', 'Demand (kVA)', 'Design current (A)', 'Incomer rating (A)', 'Board type', 'Ways required', 'Standard ways', 'PFC at board (kA)', 'Min device breaking capacity (kA)'];
TPH.forEach((h, i) => { const x = tp.getCell(34, 1 + i); x.value = h; x.fill = hdrFill; x.border = box; x.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } }; x.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' }; });
tp.getRow(34).height = 42;
const TD0 = 35;
for (let i = 0; i < 20; i++) {
  const r = TD0 + i, s = DB0 + i;
  calc(tp.getCell(`A${r}`), `IF(${CLQ}!A${s}="","",${CLQ}!A${s})`, null, GREEN);
  calc(tp.getCell(`B${r}`), `IF(A${r}="","",${CLQ}!B${s})`, '0', GREEN);
  calc(tp.getCell(`C${r}`), `IF(A${r}="","",${CLQ}!G${s})`, '#,##0.0', GREEN);
  calc(tp.getCell(`D${r}`), `IF(A${r}="","",${CLQ}!H${s})`, '#,##0.0', GREEN);
  calc(tp.getCell(`E${r}`), `IF(A${r}="","",${nextStd(MCCBR, `D${r}*(1+$B$28)`)})`, '0');
  calc(tp.getCell(`F${r}`), `IF(A${r}="","",IF(N(E${r})<=125,"TP&N MCB distribution board (Type B)",IF(N(E${r})<=630,"MCCB panelboard","Switchboard section")))`);
  calc(tp.getCell(`G${r}`), `IF(A${r}="","",ROUNDUP(B${r}*(1+$B$28),0))`, '0');
  calc(tp.getCell(`H${r}`), `IF(A${r}="","",${nextStd(WAYR, `G${r}`)})`, '0');
  input(tp.getCell(`I${r}`), i < 3 ? [10, 25, 6][i] : null, '0.0');
  calc(tp.getCell(`J${r}`), `IF(OR(A${r}="",I${r}=""),"",${nextStd(KAR, `I${r}`)})`, '0');
}
const TDE = TD0 + 20;
tp.getCell(`A${TDE}`).value = 'Board standard: BS EN 61439-2 (power switchgear) / BS EN 61439-3 (DBO). Ways = circuits × (1 + spare), rounded up to a standard board; 3-phase ways count as one TP way.';
tp.getCell(`A${TDE}`).font = { name: 'Arial', size: 9, italic: true };

const GS = TDE + 2;
tsec(GS, 'D. STANDBY GENERATOR / UPS');
trow(GS + 1, 'Essential / life-safety load as % of maximum demand', 'in', 0.4, '0%', 'Hospitals typically 60–100%; offices 10–30%');
trow(GS + 2, 'Essential load (kVA)', 'f', `B6*B${GS + 1}`, '#,##0.0', '');
trow(GS + 3, 'Sizing margin (motor starting / step load)', 'in', 1.25, '0.00', 'Typical 1.2–1.5; check largest motor start');
trow(GS + 4, 'Selected generator rating (kVA, prime)', 'f', nextStd(GENR, `B${GS + 2}*B${GS + 3}`), '#,##0', 'ISO 8528; BS 7698');
trow(GS + 5, 'Fuel autonomy required (hours)', 'in', 24, '0', 'HTM 06-01 hospitals typically 24 h+; life safety per BS 9999 / BS 7346');
trow(GS + 6, 'Indicative fuel tank (litres, 75% load)', 'f', `IF(ISNUMBER(B${GS + 4}),B${GS + 4}*${PF}*0.75*0.27*B${GS + 5},"")`, '#,##0', '≈ 0.27 l/kWh diesel; confirm with manufacturer');
trow(GS + 7, 'Critical IT / UPS load (kVA)', 'in', 0, '#,##0.0', 'Server rooms, BMS, security');
trow(GS + 8, 'UPS rating (kVA)', 'f', `B${GS + 7}*1.25`, '#,##0.0', '25% margin; BS EN 62040');

const ES = GS + 10;
tsec(ES, 'E. EARTHING AND BONDING (BS 7671 Chapter 54)');
trow(ES + 1, 'Earthing system', 'f', `'Project Info'!B25`, null, 'From region', GREEN);
trow(ES + 2, 'Main incoming line conductor csa (mm², Cu)', 'in', 240, '0', 'Main incomer cable, or equivalent csa of busbar / transformer tails');
const se = `IF(B${ES + 2}<=16,B${ES + 2},IF(B${ES + 2}<=35,16,B${ES + 2}/2))`;
trow(ES + 3, 'Main earthing conductor (mm², Table 54.7)', 'f', `INDEX(${CSA_COL},MIN(ROWS(${CSA_COL}),COUNTIF(${CSA_COL},"<"&${se})+1))`, '0', 'Copper; same material as line conductor');
trow(ES + 4, 'Main protective bonding conductor (mm²)', 'f', `IF(OR(ISNUMBER(SEARCH("PME",B${ES + 1})),ISNUMBER(SEARCH("TN-C-S",B${ES + 1})),ISNUMBER(SEARCH("MEN",B${ES + 1}))),IF(B${ES + 2}<=35,10,IF(B${ES + 2}<=50,16,IF(B${ES + 2}<=95,25,IF(B${ES + 2}<=150,35,50)))),MIN(25,MAX(6,INDEX(${CSA_COL},COUNTIF(${CSA_COL},"<"&B${ES + 3}/2)+1))))`, '0', 'PME: Table 54.8 (by supply neutral); otherwise ≥ ½ earthing conductor, min 6, max 25 mm² (544.1.1)');
trow(ES + 5, 'Supplementary bonding (where required)', 'in', '4 mm² Cu (2.5 mm² if mechanically protected)', null, '415.2; may be omitted in bathrooms where 701.415.2 conditions are met');
trow(ES + 6, 'Transformer neutral–earth conductor (mm², Cu)', 'f', `IF(ISNUMBER(B16),INDEX(${CSA_COL},MIN(ROWS(${CSA_COL}),COUNTIF(${CSA_COL},"<"&SQRT((B16*1000)^2*1)/143)+1)),"")`, '0', 'Adiabatic for 1 s LV fault, k = 143; parallel conductors where > 300 mm²');
trow(ES + 7, 'Earth electrode resistance target (Ω)', 'in', 1, '0.0', 'Combined HV/LV substation earth typically ≤ 1 Ω (ENA TS 41-24); TT installation ≤ 200 Ω');
trow(ES + 8, 'Main earthing terminal', 'in', 'Copper MET bar with disconnecting test link, labelled (542.4, 514.13)', null, '');

const TS = ES + 10;
tsec(TS, 'F. TRANSFORMER / SUBSTATION AND LV SWITCHROOM SIZING');
const TXD = [[200, 1.6, 1.0, 1.7, 1.1, 3.0, 300], [315, 1.7, 1.0, 1.8, 1.4, 4.2, 400], [500, 1.9, 1.1, 1.9, 1.9, 6.0, 500], [800, 2.1, 1.2, 2.1, 2.6, 8.5, 700], [1000, 2.2, 1.2, 2.2, 3.0, 10.0, 800], [1250, 2.3, 1.3, 2.3, 3.6, 12.0, 950], [1500, 2.4, 1.3, 2.4, 4.2, 14.0, 1100], [1600, 2.4, 1.3, 2.4, 4.4, 15.0, 1150], [2000, 2.6, 1.4, 2.5, 5.3, 18.0, 1400]];
['kVA', 'L (m)', 'W (m)', 'H (m)', 'Weight (t)', 'Total losses (kW)', 'Oil (litres, ONAN)'].forEach((h, i) => { const x = tp.getCell(4, 12 + i); x.value = h; x.fill = hdrFill; x.border = box; x.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFFFF' } }; x.alignment = { wrapText: true, horizontal: 'center' }; });
tp.getCell('L3').value = 'Typical transformer envelope incl. IP31 enclosure – confirm with manufacturer'; tp.getCell('L3').font = { name: 'Arial', size: 9, italic: true };
TXD.forEach((row, k) => row.forEach((v, i) => input(tp.getCell(5 + k, 12 + i), v, i === 0 || i === 6 ? '#,##0' : '0.0')));
const TXr = (col) => `${col}$5:${col}${4 + TXD.length}`;
const tm = `MATCH(B${TS + 1},${TXr('L')},0)`;
const tl = (col) => `IFERROR(INDEX(${TXr(col)},${tm}),0)`;
trow(TS + 1, 'Transformer rating (kVA)', 'f', 'B11', '#,##0', 'From section A', GREEN);
trow(TS + 2, 'Number of transformers', 'f', 'B8', '0', '', GREEN);
trow(TS + 3, 'Unit envelope L × W × H (m)', 'f', `IF(${tl('M')}=0,"Enter dimensions – size not in table",TEXT(${tl('M')},"0.0")&" × "&TEXT(${tl('N')},"0.0")&" × "&TEXT(${tl('O')},"0.0"))`, null, 'Typical cast-resin with enclosure');
trow(TS + 4, 'Unit weight (t) – floor loading / lifting route', 'f', tl('P'), '0.0', '');
trow(TS + 5, 'Clearance around each transformer (m)', 'in', 1.0, '0.0', 'Maintenance access 0.8–1.0 m (BS EN 61936-1 / ENA guidance)');
trow(TS + 6, 'Transformer room length (m)', 'f', `B${TS + 2}*${tl('M')}+(B${TS + 2}+1)*B${TS + 5}`, '0.0', 'Units side by side');
trow(TS + 7, 'Transformer room width (m)', 'f', `${tl('N')}+2*B${TS + 5}`, '0.0', '');
trow(TS + 8, 'Clear height (m)', 'f', `MAX(3,${tl('O')}+0.8)`, '0.0', 'Min 3.0 m; allow for cable trays / lifting');
trow(TS + 9, 'Transformer room area (m²)', 'f', `B${TS + 6}*B${TS + 7}`, '0.0', '');
trow(TS + 10, 'HV switchgear / RMU room allowance (m²)', 'in', 6, '0.0', 'RMU ≈ 1.6 × 1.0 m + 1.0 m front clearance; utility to confirm');
trow(TS + 11, 'LV switchboard outgoing ways', 'in', 24, '0', '');
trow(TS + 12, 'LV switchboard length (m)', 'f', `(ROUNDUP(B${TS + 11}/4,0)+B${TS + 2}+IF(B${TS + 2}>1,1,0))*0.8`, '0.0', '≈ 0.8 m per section (4 ways/section, incomers, bus-coupler)');
trow(TS + 13, 'LV switchboard depth (m)', 'in', 1.0, '0.0', '');
trow(TS + 14, 'Front clearance / ACB withdrawal (m)', 'in', 1.5, '0.0', 'BS 7671 Section 729: ≥ 0.7 m passage; ACB withdrawal typically 1.2–1.5 m');
trow(TS + 15, 'Rear access (m)', 'in', 0.8, '0.0', '0 for front-access-only boards');
trow(TS + 16, 'LV switchroom size L × W (m)', 'f', `TEXT(B${TS + 12}+1.6,"0.0")&" × "&TEXT(B${TS + 13}+B${TS + 14}+B${TS + 15},"0.0")`, null, '0.8 m end clearance each side; two exits if board > 20 m or > 3,000 A (729.3)');
trow(TS + 17, 'LV switchroom area (m²)', 'f', `(B${TS + 12}+1.6)*(B${TS + 13}+B${TS + 14}+B${TS + 15})`, '0.0', '');
trow(TS + 18, 'TOTAL substation + switchroom area (m²)', 'f', `B${TS + 9}+B${TS + 10}+B${TS + 17}`, '0.0', '');
trow(TS + 19, 'Transformer heat loss (kW)', 'f', `B${TS + 2}*${tl('Q')}`, '0.0', 'Full-load losses');
trow(TS + 20, 'Permitted air temperature rise ΔT (K)', 'in', 10, '0', '');
trow(TS + 21, 'Ventilation airflow required (m³/s)', 'f', `IF(B${TS + 20}=0,0,B${TS + 19}/(1.2*1.005*B${TS + 20}))`, '0.00', 'Q = P ÷ (ρ·cp·ΔT)');
trow(TS + 22, 'Louvre free area – inlet and outlet each (m²)', 'f', `B${TS + 21}/1.5`, '0.00', 'At 1.5 m/s face velocity (mechanical); double for natural ventilation');
trow(TS + 23, 'Oil containment bund (litres)', 'f', `IF(B17="Oil-immersed (ONAN)",B${TS + 2}*${tl('R')}*1.1,"n/a – dry type")`, '#,##0', '110% of oil volume');
trow(TS + 24, 'Fire separation', 'in', 'Oil-filled: 2 h fire-rated enclosure; cast resin: 1 h (confirm with fire strategy)', null, '');
trow(TS + 25, 'Access', 'in', 'Direct external access for utility 24/7; door ≥ unit width + 0.3 m; lifting route and floor loading checked', null, '');

const GSP = TS + 27;
tsec(GSP, 'G. SWITCHGEAR AND PROTECTION SPECIFICATION (select / edit)');
const spec = [
  ['Final circuit devices – general', 'MCB Type B (BS EN 60898-1)', '"MCB Type B (BS EN 60898-1),MCB Type C (BS EN 60898-1),MCB Type D (BS EN 60898-1),RCBO Type A (BS EN 61009-1),RCBO Type B (BS EN 61009-1)"', 'Type C/D for motor / inductive loads'],
  ['MCB / RCBO breaking capacity (kA)', '10', '"6,10,15,25"', 'Must exceed PFC at the board (434.5.1)'],
  ['RCD protection – socket circuits ≤ 32 A', 'Type A 30 mA RCBO per circuit', '"Type A 30 mA RCBO per circuit,Type A 30 mA RCCB per group,Type F 30 mA,Type B 30 mA"', '411.3.3; Type AC restricted (531.3.3)'],
  ['RCD protection – EV charge points', 'Type A + RDC-DD 6 mA DC', '"Type A + RDC-DD 6 mA DC,Type B 30 mA"', 'Section 722'],
  ['Arc fault detection devices (AFDD)', 'Recommended', '"Required (HRRB / HMO / PBSA / care home),Recommended,Not required"', '421.1.7'],
  ['Surge protection devices (SPD)', 'Type 1+2 at origin; Type 2 at DBs', '"Type 1+2 at origin; Type 2 at DBs,Type 2 at origin,Type 2+3 (sensitive equipment)"', 'Section 443 / 534'],
  ['Distribution boards', 'TP&N Type B MCB board (BS EN 61439-3)', '"TP&N Type B MCB board (BS EN 61439-3),Consumer unit – non-combustible (BS EN 61439-3),MCCB panelboard (BS EN 61439-2)"', ''],
  ['MCCB trip units', 'Electronic LSI', '"Thermal-magnetic,Electronic LSI,Electronic LSIG"', 'G = earth-fault protection'],
  ['ACB', 'Draw-out, motorised, electronic LSIG, Modbus comms', null, 'BS EN 60947-2'],
  ['Main switchboard construction', 'Form 4 Type 2, IP31, tin-plated copper busbars, 100% neutral, full-length earth bar', null, 'BS EN 61439-2'],
  ['Discrimination / selectivity', 'Full selectivity ACB → MCCB → MCB per manufacturer tables', null, '536.4'],
  ['Isolation and switching', 'Lockable isolators at plant; emergency switching for kitchens / workshops', null, 'Chapter 46 / 537'],
  ['Sub-main cables', 'XLPE/SWA/LSZH (BS 6724) on tray / ladder', null, ''],
  ['Final circuit cables', 'LSZH singles in conduit / trunking (BS 7211) or LSZH multicore', null, ''],
  ['Life-safety circuits', 'Enhanced fire-resistant cables (BS 8519 / BS 7629-1, BS EN 50200 PH120)', null, 'Fire pumps, smoke fans, fire-fighting lifts'],
  ['Metering', 'MID class 1 multifunction meters on main and sub-main ways; BMS integration', null, 'Building Regulations Part L'],
];
spec.forEach(([l, v, list, n], i) => {
  const r = GSP + 1 + i;
  trow(r, l, 'in', v, null, n);
  if (list) tp.getCell(`B${r}`).dataValidation = { type: 'list', allowBlank: true, formulae: [list] };
});
tp.getColumn('A').width = 46; tp.getColumn('B').width = 34; tp.getColumn('C').width = 52; tp.getColumn('D').width = 14; tp.getColumn('E').width = 14;
tp.getColumn('F').width = 34; tp.getColumn('G').width = 11; tp.getColumn('H').width = 12; tp.getColumn('I').width = 12; tp.getColumn('J').width = 14; for (let c = 12; c <= 18; c++) tp.getColumn(c).width = 11;

