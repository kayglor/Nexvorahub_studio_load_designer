let EMQ, EM_EUI, EM_MDA, EM_BM;
{
// ---------------- 18c. Energy Management ----------------
const em = wb.addWorksheet('Energy Management');
EMQ = "'Energy Management'";
title(em, 'Energy Management – Annual Energy, Savings Measures, Renewables, Demand Management and Power Factor', 'System energy uses demand × equivalent full-load hours (EFLH) from the Building Services Systems sheet; measures reduce the relevant end uses. Regional carbon factors and PV yields are indicative defaults – replace with project / utility data.');
legend(em, 3);
// regional defaults table (columns J..M)
const EREG = [['United Kingdom', 0.21, 900], ['Europe (IEC)', 0.25, 1100], ['USA (NEC)', 0.37, 1400], ['Canada (CEC)', 0.12, 1150], ['Australia / NZ', 0.6, 1500], ['Middle East (GCC)', 0.5, 1750], ['Africa', 0.55, 1650], ['South & SE Asia', 0.7, 1400], ['Latin America & Caribbean', 0.2, 1500], ['East Asia', 0.55, 1200]];
['Density region', 'Grid carbon (kgCO₂e/kWh)', 'PV yield (kWh/kWp/yr)'].forEach((h, i) => { const x = em.getCell(5, 10 + i); x.value = h; x.fill = hdrFill; x.border = box; x.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFFFF' } }; x.alignment = { wrapText: true, horizontal: 'center' }; });
em.getRow(5).height = 30;
EREG.forEach(([n, c, y], i) => { const a = em.getCell(6 + i, 10); a.value = n; a.border = box; a.font = FONT; input(em.getCell(6 + i, 11), c, '0.00'); input(em.getCell(6 + i, 12), y, '#,##0'); });
em.getCell('J17').value = 'Indicative averages; national / state factors vary widely (e.g. NZ, Canadian provinces, South Africa).'; em.getCell('J17').font = { name: 'Arial', size: 8, italic: true };
const ER = `$J$6:$J$${5 + EREG.length}`;
const erow = (r, label, kind, v, fmt, note, color) => {
  const a = em.getCell(`A${r}`); a.value = label; a.border = box; a.font = FONT;
  if (kind === 'in') input(em.getCell(`B${r}`), v, fmt); else calc(em.getCell(`B${r}`), v, fmt, color);
  if (note) { const n = em.getCell(`D${r}`); n.value = note; n.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } }; }
};
const esec = (r, t) => { const c = em.getCell(`A${r}`); c.value = t; c.font = { name: 'Arial', size: 10, bold: true }; ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].forEach(x => em.getCell(`${x}${r}`).fill = secFill); };
esec(5, 'A. PROJECT ENERGY INPUTS');
erow(6, 'Density region', 'f', `'Project Info'!B35`, null, 'From Project Info', GREEN);
erow(7, 'Floor area for EUI (m²)', 'f', `IF(C7<>"",C7,IF(Summary!C24>0,Summary!C24,${SYQ}!B5))`, '#,##0', 'Override in C7', GREEN); input(em.getCell('C7'), null, '#,##0');
erow(8, 'Currency', 'in', 'GBP', null, '');
erow(9, 'Energy tariff (per kWh)', 'in', 0.25, '0.000', 'Blended unit rate');
erow(10, 'Capacity / demand charge (per kVA per month)', 'in', 3, '0.00', 'Utility availability / demand charge');
erow(11, 'Grid carbon factor (kgCO₂e/kWh)', 'f', `IF(C11<>"",C11,IFERROR(INDEX($K$6:$K$15,MATCH(B6,${ER},0)),0.3))`, '0.000', 'Regional default; override in C11', GREEN); input(em.getCell('C11'), null, '0.000');
erow(12, 'PV specific yield (kWh/kWp/yr)', 'f', `IF(C12<>"",C12,IFERROR(INDEX($L$6:$L$15,MATCH(B6,${ER},0)),1200))`, '#,##0', 'Regional default; override in C12', GREEN); input(em.getCell('C12'), null, '#,##0');
erow(13, 'Annual load factor (MD cross-check method)', 'in', 0.45, '0.00', 'Offices 0.35–0.45; hospitals / data centres 0.6–0.9');
erow(14, 'Electricity EUI benchmark (kWh/m²/yr)', 'in', 95, '#,##0', 'e.g. CIBSE TM46 general office 95; set per building type / local energy code');

esec(16, 'B. ENERGY MANAGEMENT MEASURES');
header(em, 17, ['Measure', 'Include?', 'Savings %', 'Applies to', 'Standard / reference']);
const MEAS = [
  ['LED lighting with daylight-linked dimming', 'Yes', 0.3, 'Lighting', 'EN 15193-1 LENI / ASHRAE 90.1 §9.4'],
  ['Presence / absence detection', 'Yes', 0.15, 'Lighting', 'EN 15193-1 / ASHRAE 90.1 §9.4.1'],
  ['BMS optimisation (scheduling, setpoints, optimum start)', 'Yes', 0.15, 'HVAC', 'ISO 52120-1 BACS class A/B'],
  ['Heat recovery ventilation', 'No', 0.1, 'HVAC', 'Ecodesign (EU) 1253/2014'],
  ['Variable speed drives on pumps and fans', 'Yes', 0.25, 'Pumps & fans', 'IEC 61800-9; IE3/IE4 motors (IEC 60034-30-1)'],
  ['Energy monitoring & targeting (sub-metering, M&T)', 'Yes', 0.05, 'All', 'ISO 50001; Part L; ASHRAE 90.1 §8.4.3'],
  ['EV smart charging (OCPP, off-peak scheduling)', 'Yes', 0, 'EV', 'IEC 61851; OCPP 2.0.1 – shifts demand, no kWh saving'],
];
const M0 = 18, M1 = M0 + MEAS.length - 1;
MEAS.forEach(([n, inc, sv, grp, ref], i) => {
  const r = M0 + i;
  const a = em.getCell(`A${r}`); a.value = n; a.border = box; a.font = FONT;
  input(em.getCell(`B${r}`), inc); em.getCell(`B${r}`).dataValidation = { type: 'list', allowBlank: false, formulae: ['"Yes,No"'] };
  input(em.getCell(`C${r}`), sv, '0%');
  input(em.getCell(`D${r}`), grp); em.getCell(`D${r}`).dataValidation = { type: 'list', allowBlank: false, formulae: ['"Lighting,HVAC,Pumps & fans,EV,Other,All"'] };
  const e = em.getCell(`E${r}`); e.value = ref; e.border = box; e.font = { name: 'Arial', size: 9 };
});

const S0h = M1 + 2;
esec(S0h, 'C. ANNUAL ENERGY BY SYSTEM (from Building Services Systems)');
header(em, S0h + 1, ['System', 'Included?', 'Demand (kW)', 'EFLH (h/yr)', 'Baseline (kWh/yr)', 'End-use group', 'Savings %', 'Net (kWh/yr)']);
em.getRow(S0h + 1).height = 30;
const eflh = (n) => /Emergency|Fire alarm|Security|BMS/i.test(n) ? 8760 : /Street|external/i.test(n) ? 4100 : /Common area lighting/i.test(n) ? 4000 : /General lighting/i.test(n) ? 2500 : /Socket/i.test(n) ? 2000 : /AHU|Chilled|pumps – all|Pumps – all/i.test(n) ? 3000 : /Escalator/i.test(n) ? 4000 : /Lift/i.test(n) ? 1000 : /booster/i.test(n) ? 2000 : /Sewage/i.test(n) ? 500 : /Fire pump/i.test(n) ? 50 : /Smoke/i.test(n) ? 20 : /Heat pumps – space|heater/i.test(n) ? 1500 : /hot water/i.test(n) ? 1500 : /chillers|Air conditioning|Cooling towers/i.test(n) ? 1200 : /EV/i.test(n) ? 1500 : 1000;
const grpOf = (n) => /lighting/i.test(n) ? 'Lighting' : /pumps|booster|Sewage|fans/i.test(n) ? 'Pumps & fans' : /heater|Air conditioning|chillers|Cooling towers|Heat pumps|AHU|hot water/i.test(n) ? 'HVAC' : /EV/i.test(n) ? 'EV' : 'Other';
const E0 = S0h + 2;
SYS.forEach((s, i) => {
  const r = E0 + i, sr = SY0 + i;
  calc(em.getCell(`A${r}`), `${SYQ}!A${sr}`, null, GREEN);
  calc(em.getCell(`B${r}`), `${SYQ}!B${sr}`, null, GREEN);
  calc(em.getCell(`C${r}`), `${SYQ}!O${sr}`, '#,##0.0;-#,##0.0;"-"', GREEN);
  input(em.getCell(`D${r}`), eflh(s[0]), '#,##0');
  calc(em.getCell(`E${r}`), `IF(B${r}="Yes",C${r}*D${r},0)`, '#,##0;-#,##0;"-"');
  input(em.getCell(`F${r}`), grpOf(s[0])); em.getCell(`F${r}`).dataValidation = { type: 'list', allowBlank: false, formulae: ['"Lighting,HVAC,Pumps & fans,EV,Other"'] };
  const prod = MEAS.map((_, k) => `(1-IF(AND($B$${M0 + k}="Yes",OR($D$${M0 + k}=F${r},$D$${M0 + k}="All")),$C$${M0 + k},0))`).join('*');
  calc(em.getCell(`G${r}`), `1-${prod}`, '0%');
  calc(em.getCell(`H${r}`), `E${r}*(1-G${r})`, '#,##0;-#,##0;"-"');
});
const E1 = E0 + SYS.length - 1, ET2 = E1 + 1;
em.getCell(`A${ET2}`).value = 'TOTAL';
['E', 'H'].forEach(c => calc(em.getCell(`${c}${ET2}`), `SUM(${c}${E0}:${c}${E1})`, '#,##0'));
calc(em.getCell(`G${ET2}`), `IF(E${ET2}=0,0,1-H${ET2}/E${ET2})`, '0%');
'ABCDEFGH'.split('').forEach(c => { const x = em.getCell(`${c}${ET2}`); x.fill = totFill; x.border = box; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); });
em.getCell(`A${ET2 + 1}`).value = 'Systems must be ticked "Yes" on Building Services Systems to be counted. EFLH = equivalent full-load hours; edit to suit occupancy pattern and climate.';
em.getCell(`A${ET2 + 1}`).font = { name: 'Arial', size: 9, italic: true };

const RN = ET2 + 3;
esec(RN, 'D. RENEWABLES, STORAGE AND DEMAND MANAGEMENT');
erow(RN + 1, 'Solar PV array (kWp)', 'in', 100, '#,##0.0', 'Roof / canopy area ≈ 5–7 m² per kWp');
erow(RN + 2, 'PV annual generation (kWh/yr)', 'f', `B${RN + 1}*B12`, '#,##0', '');
erow(RN + 3, 'PV self-consumption (%)', 'in', 0.9, '0%', 'Remainder exported (check export limits / G99, NEC 705, AS/NZS 4777)');
erow(RN + 4, 'Battery storage – power (kW)', 'in', 50, '#,##0', '');
erow(RN + 5, 'Battery storage – energy (kWh)', 'in', 100, '#,##0', 'Peak-shaving duration = kWh ÷ kW');
erow(RN + 6, 'Demand response / load shedding (% of MD)', 'in', 0.05, '0%', 'Non-essential loads shed by EMS at peak');
erow(RN + 7, 'Maximum demand before management (kVA)', 'f', `Summary!C18`, '#,##0.0', 'From Summary', GREEN);
erow(RN + 8, 'Maximum demand after management (kVA)', 'f', `MAX(0,B${RN + 7}*(1-B${RN + 6})-B${RN + 4}/${PF})`, '#,##0.0', 'Battery peak shaving + load shedding');
erow(RN + 9, 'MD reduction (kVA)', 'f', `B${RN + 7}-B${RN + 8}`, '#,##0.0', '');
erow(RN + 10, 'Transformer size after management (kVA, per unit)', 'f', `IFERROR(INDEX(${TX},COUNTIF(${TX},"<"&B${RN + 8}/'Transformer & Panels'!B7/MAX(1,'Transformer & Panels'!B8))+1),"> max")`, '#,##0', 'Compare with Transformer & Panels selection');

const PFC = RN + 12;
esec(PFC, 'E. POWER FACTOR CORRECTION');
erow(PFC + 1, 'Active power demand (kW)', 'f', `Summary!C19`, '#,##0.0', 'From Summary', GREEN);
erow(PFC + 2, 'Uncorrected power factor', 'in', 0.85, '0.00', 'Measured / estimated before correction');
erow(PFC + 3, 'Target power factor', 'in', 0.95, '0.00', 'Utility minimum typically 0.9–0.95');
erow(PFC + 4, 'Reactive compensation required (kVAr)', 'f', `IF(B${PFC + 2}>=B${PFC + 3},0,B${PFC + 1}*(TAN(ACOS(B${PFC + 2}))-TAN(ACOS(B${PFC + 3}))))`, '#,##0', 'Q = P × (tan φ1 − tan φ2)');
erow(PFC + 5, 'Capacitor bank step size (kVAr)', 'in', 25, '0', '');
erow(PFC + 6, 'Automatic capacitor bank rating (kVAr)', 'f', `CEILING(B${PFC + 4},MAX(1,B${PFC + 5}))`, '#,##0', '');
erow(PFC + 7, 'Harmonic-producing load (VSD / UPS / LED) – % of total', 'in', 0.3, '0%', '');
erow(PFC + 8, 'Recommended bank type', 'f', `IF(B${PFC + 4}=0,"Not required",IF(B${PFC + 7}>0.2,"Detuned (7% / 189 Hz) automatic bank or active harmonic filter – IEEE 519 / ENA G5/5 / IEC 61000-3","Standard automatic capacitor bank (IEC 60831)"))`, null, '');

const RS2 = PFC + 10;
esec(RS2, 'F. RESULTS');
erow(RS2 + 1, 'Baseline electricity use (kWh/yr)', 'f', `E${ET2}`, '#,##0', 'System method');
erow(RS2 + 2, 'Savings from measures (kWh/yr)', 'f', `E${ET2}-H${ET2}`, '#,##0', '');
erow(RS2 + 3, 'PV self-consumed (kWh/yr)', 'f', `MIN(H${ET2},B${RN + 2}*B${RN + 3})`, '#,##0', '');
erow(RS2 + 4, 'Net grid import (kWh/yr)', 'f', `MAX(0,H${ET2}-B${RS2 + 3})`, '#,##0', '');
erow(RS2 + 5, 'Net electricity EUI (kWh/m²/yr)', 'f', `IF(B7=0,0,B${RS2 + 4}/B7)`, '#,##0', '');
erow(RS2 + 6, 'Benchmark check', 'f', `IF(B${RS2 + 5}<=B14,"Within benchmark ("&B14&" kWh/m²/yr)","Above benchmark – review measures")`, null, '');
erow(RS2 + 7, 'Annual energy cost', 'f', `B8&" "&TEXT(B${RS2 + 4}*B9+B${RN + 8}*B10*12,"#,##0")`, null, 'Energy + capacity charge (after demand management)');
erow(RS2 + 8, 'Annual cost saving vs baseline', 'f', `B8&" "&TEXT((B${RS2 + 1}-B${RS2 + 4})*B9+(B${RN + 7}-B${RN + 8})*B10*12,"#,##0")`, null, '');
erow(RS2 + 9, 'Annual carbon emissions (tCO₂e/yr)', 'f', `B${RS2 + 4}*B11/1000`, '#,##0.0', 'Location-based');
erow(RS2 + 10, 'Carbon saving (tCO₂e/yr)', 'f', `(B${RS2 + 1}-B${RS2 + 4})*B11/1000`, '#,##0.0', '');
erow(RS2 + 11, 'Cross-check: MD × load factor × 8,760 h (kWh/yr)', 'f', `Summary!C19*B13*8760`, '#,##0', 'Top-down check; large differences indicate EFLH or load-factor assumptions to review');
[RS2 + 5, RS2 + 7].forEach(r => { em.getCell(`B${r}`).font = { name: 'Arial', size: 10, bold: true }; });

const ES2 = RS2 + 13;
esec(ES2, 'G. ENERGY MANAGEMENT SYSTEM (EMS) SPECIFICATION (select / edit)');
const EMS = [
  ['Sub-metering scheme', 'Main incomer + every sub-main + major plant (lighting, small power, HVAC, lifts, EV, PV)', 'Part L / ASHRAE 90.1 §8.4.3 / NCC J'],
  ['Meter accuracy and protocol', 'MID class 1 (IEC 62053-21/22), Modbus TCP to BMS / EMS', ''],
  ['BACS efficiency class', 'Class B', 'ISO 52120-1'],
  ['Energy management system standard', 'ISO 50001 aligned monitoring & targeting', 'ISO 50001'],
  ['Demand response / load shedding', 'Priority shedding of non-essential loads (EV, AC set-back, decorative lighting) on MD threshold', ''],
  ['PV / battery EMS', 'Self-consumption optimisation, export limitation, peak shaving', 'G99 / NEC 705 / AS/NZS 4777'],
  ['EV charging management', 'Dynamic load management, OCPP 2.0.1, off-peak scheduling', 'IEC 61851'],
  ['Power quality monitoring', 'Class A analyser at incomer (harmonics, flicker, sags)', 'IEC 61000-4-30 / EN 50160 / IEEE 519'],
  ['Dashboards and reporting', 'Monthly energy & carbon report, EUI tracking, alarms on abnormal consumption', ''],
];
EMS.forEach(([l, v, ref], i) => {
  const r = ES2 + 1 + i;
  erow(r, l, 'in', v, null, ref);
  em.getCell(`B${r}`).alignment = { wrapText: true };
  em.mergeCells(`B${r}:C${r}`);
});
em.getCell('B' + (ES2 + 3)).dataValidation = { type: 'list', allowBlank: true, formulae: ['"Class A,Class B,Class C,Class D"'] };
em.getColumn('A').width = 50; em.getColumn('B').width = 16; em.getColumn('C').width = 14; em.getColumn('D').width = 22; em.getColumn('E').width = 18;
em.getColumn('F').width = 14; em.getColumn('G').width = 11; em.getColumn('H').width = 14; em.getColumn('I').width = 3; em.getColumn('J').width = 24; em.getColumn('K').width = 14; em.getColumn('L').width = 14;
EM_EUI = RS2 + 5; EM_MDA = RN + 8; EM_BM = 14;


}
