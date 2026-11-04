// ---------------- 18a. Building Services Systems ----------------
const sy = wb.addWorksheet('Building Services Systems');
title(sy, 'Building Services Systems – System-by-System Load Estimate', 'Tick the systems the client requires. Each is sized from floor area, cooling/heating load or quantity, then diversity is applied from the Diversity Factors sheet.');
legend(sy, 3);
const syIn = [
  [5, 'Building GIA (m²)', `IF(C5<>"",C5,${BCQ}!E${FB})`, 'Default from Building Calculator; enter override in C5'],
  [6, 'Common area (m²) – corridors, lobbies, stairs, car park circulation', `IF(C6<>"",C6,B5*0.15)`, 'Default 15% of GIA; enter override in C6'],
];
syIn.forEach(([r, l, f, n]) => {
  const a = sy.getCell(`A${r}`); a.value = l; a.border = box; a.font = { name: 'Arial', size: 10, bold: true };
  calc(sy.getCell(`B${r}`), f, '#,##0', GREEN);
  input(sy.getCell(`C${r}`), null, '#,##0');
  const x = sy.getCell(`D${r}`); x.value = n; x.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } };
});
header(sy, 8, ['System', 'Include?', 'Sizing basis', 'Parameter', 'Unit', 'Default driver (m² / no. / kW)', 'Driver override', 'Applied driver', 'EER / COP / SFP', 'Thermal load (kW)', 'Electrical connected (kW)', 'Load category', 'Ku', 'Diversity', 'Demand (kW)', 'Demand (kVA)', 'Notes']);
sy.getRow(8).height = 45;
const SY0 = 9;
// [name, include, basis, param, unit, type, driver formula/const, eff, category, note]
const GIA = '$B$5', COM = '$B$6';
const SYS = [
  ['General lighting (LED)', 'Yes', 'W/m² × GIA', 8, 'W/m²', 'wpm2', GIA, null, 'General lighting', 'See Lighting Design for lux-based check'],
  ['Common area lighting', 'Yes', 'W/m² × common area', 5, 'W/m²', 'wpm2', COM, null, 'General lighting', 'Presence detection recommended'],
  ['Emergency lighting', 'Yes', 'W/m² × GIA', 1, 'W/m²', 'wpm2', GIA, null, 'Emergency / life-safety lighting', 'Self-contained or central battery'],
  ['Socket outlets / small power', 'Yes', 'W/m² × GIA', 15, 'W/m²', 'wpm2', GIA, null, 'Small power – general sockets', ''],
  ['Electric heaters (panel / convector / storage)', 'No', 'Heat loss W/m² ÷ efficiency', 50, 'W/m²', 'heat', GIA, 1, 'Electric heating – direct / panel', 'Efficiency 1.0 for direct electric'],
  ['Air conditioning – split / VRF systems', 'No', 'Cooling W/m² ÷ EER', 100, 'W/m²', 'cool', GIA, 3.5, 'Air conditioning – split / VRF', 'Typical EER 3.0–4.0'],
  ['Central cooling – chillers', 'Yes', 'Cooling W/m² ÷ EER', 100, 'W/m²', 'cool', GIA, 3.0, 'Central cooling – chillers & cooling towers', 'Air-cooled EER ~3.0; water-cooled ~5.0'],
  ['Chilled & condenser water pumps', 'Yes', 'kWe per kW cooling', 0.04, 'kWe/kWth', 'pump', 'CHILLER', null, 'Chilled / condenser water pumps', 'Typical 0.03–0.05'],
  ['Cooling towers / dry-air coolers (fans)', 'No', 'kWe per kW cooling', 0.02, 'kWe/kWth', 'pump', 'CHILLER', null, 'Central cooling – chillers & cooling towers', 'Water-cooled systems only'],
  ['Heat pumps – space heating', 'Yes', 'Heat load W/m² ÷ COP', 50, 'W/m²', 'heat', GIA, 2.5, 'HVAC – chillers / heat pumps', 'Design COP at winter design temperature'],
  ['Heat pumps / water heaters – domestic hot water', 'Yes', 'W/m² × GIA', 3, 'W/m²', 'wpm2', GIA, null, 'Water heating – storage / thermostatic', 'Hotels / hospitals 10–15 W/m²'],
  ['HVAC – AHUs and ventilation fans', 'Yes', 'Air l/s per m² × SFP', 1.5, 'l/s/m²', 'vent', GIA, 1.6, 'HVAC – AHUs / fans / pumps', 'SFP W/(l/s): 1.6 AHU (Part L limit 1.8–2.0)'],
  ['Lifts / elevators', 'Yes', 'Factor × connected lift kW (Pumps & Lifts)', 1, 'factor', 'qty', `SUMPRODUCT(${PLQ}!P${L0r}:P${L1r},${PLQ}!D${L0r}:D${L1r})*${PF}`, null, 'Lifts', 'Persons-based lift sizing on Pumps & Lifts; lift diversity applied here'],
  ['Pumps – all normal services (Pumps & Lifts)', 'No', 'Factor × duty pump kVA→kW', 1, 'factor', 'qty', `${PLQ}!U${PT}*${PF}`, null, 'Motors – general', 'Use instead of the individual pump rows below when pumps are scheduled on Pumps & Lifts'],
  ['Escalators / travelators', 'No', 'kW each × number', 7.5, 'kW each', 'qty', 0, null, 'Escalators / travelators', ''],
  ['Domestic water booster pumps', 'Yes', 'kW each × number', 5.5, 'kW each', 'qty', 2, null, 'Domestic water booster pumps', 'Duty / standby set'],
  ['Sewage / sump / drainage pumps', 'Yes', 'kW each × number', 3, 'kW each', 'qty', 2, null, 'Sewage / sump pumps', 'Duty / standby'],
  ['Fire pumps (sprinkler / wet riser)', 'No', 'kW each × number', 45, 'kW each', 'qty', 1, null, 'Fire-fighting / life-safety pumps', 'Normally dedicated life-safety supply (BS EN 12845) – excluded from MD'],
  ['Smoke extract / stair pressurisation fans', 'No', 'kW each × number', 11, 'kW each', 'qty', 2, null, 'Smoke extract fans (life safety)', 'Secondary supply / generator'],
  ['Fire alarm & detection system', 'Yes', 'W/m² × GIA', 0.1, 'W/m²', 'wpm2', GIA, null, 'Fire & security alarm systems / BMS', 'BS 5839-1'],
  ['Security, access control, CCTV, intruder & power-failure alarms', 'Yes', 'W/m² × GIA', 0.5, 'W/m²', 'wpm2', GIA, null, 'Fire & security alarm systems / BMS', ''],
  ['BMS, controls and comms rooms', 'Yes', 'W/m² × GIA', 2, 'W/m²', 'wpm2', GIA, null, 'IT / server / comms', ''],
  ['EV charge points', 'No', 'kW each × number', 7.4, 'kW each', 'qty', 0, null, 'EV charging – load managed', 'Diversity set by load management'],
  ['Street, external & common-area external lighting', 'Yes', 'Factor × External Lighting kW', 1, 'factor', 'qty', `${EXQ}!N${ET}`, null, 'External / common-area lighting', 'From External Lighting sheet'],
  ['Other client-specified system', 'No', 'kW each × number', 0, 'kW each', 'qty', 0, null, 'Other / unclassified', 'Describe in notes'],
];
const SY1 = SY0 + SYS.length - 1;
const CHR = SY0 + SYS.findIndex(s => s[0] === 'Central cooling – chillers');
SYS.forEach(([name, inc, basis, param, unit, type, drv, eff, cat, note], i) => {
  const r = SY0 + i;
  const a = sy.getCell(`A${r}`); a.value = name; a.border = box; a.font = FONT;
  input(sy.getCell(`B${r}`), inc); sy.getCell(`B${r}`).alignment = { horizontal: 'center' };
  sy.getCell(`B${r}`).dataValidation = { type: 'list', allowBlank: false, formulae: ['"Yes,No"'] };
  const c = sy.getCell(`C${r}`); c.value = basis; c.border = box; c.font = { name: 'Arial', size: 9 };
  input(sy.getCell(`D${r}`), param, '#,##0.0##');
  const u = sy.getCell(`E${r}`); u.value = unit; u.border = box; u.font = { name: 'Arial', size: 9 };
  const d = drv === 'CHILLER' ? `J${CHR}` : drv;
  if (typeof d === 'number') input(sy.getCell(`F${r}`), d, '#,##0.0'); else calc(sy.getCell(`F${r}`), d, '#,##0.0', GREEN);
  input(sy.getCell(`G${r}`), null, '#,##0.0');
  calc(sy.getCell(`H${r}`), `IF(G${r}<>"",G${r},N(F${r}))`, '#,##0.0');
  if (eff !== null) input(sy.getCell(`I${r}`), eff, '0.00'); else { sy.getCell(`I${r}`).border = box; }
  const th = (type === 'cool' || type === 'heat') ? `D${r}*H${r}/1000` : null;
  if (th) calc(sy.getCell(`J${r}`), th, '#,##0.0'); else sy.getCell(`J${r}`).border = box;
  const el = { wpm2: `D${r}*H${r}/1000`, cool: `IF(N(I${r})=0,0,J${r}/I${r})`, heat: `IF(N(I${r})=0,0,J${r}/I${r})`, vent: `D${r}*H${r}*N(I${r})/1000`, qty: `D${r}*H${r}`, pump: `D${r}*H${r}` }[type];
  calc(sy.getCell(`K${r}`), el, '#,##0.0;-#,##0.0;"-"');
  input(sy.getCell(`L${r}`), cat); sy.getCell(`L${r}`).dataValidation = { type: 'list', allowBlank: false, formulae: [CAT_LIST] };
  calc(sy.getCell(`M${r}`), `IFERROR(INDEX(${CAT_KU},MATCH(L${r},${CAT_LIST},0)),1)`, '0.00', GREEN);
  calc(sy.getCell(`N${r}`), `IFERROR(INDEX(${CAT_DF},MATCH(L${r},${CAT_LIST},0)),1)`, '0.00', GREEN);
  calc(sy.getCell(`O${r}`), `IF(B${r}="Yes",K${r}*M${r}*N${r},0)`, '#,##0.0;-#,##0.0;"-"');
  calc(sy.getCell(`P${r}`), `IF(${PF}=0,0,O${r}/${PF})`, '#,##0.0;-#,##0.0;"-"');
  const n = sy.getCell(`Q${r}`); n.value = note; n.border = box; n.font = { name: 'Arial', size: 9, italic: true };
});
const SYT = SY1 + 1;
sy.getCell(`A${SYT}`).value = 'TOTAL – included systems'; sy.mergeCells(`A${SYT}:I${SYT}`);
calc(sy.getCell(`K${SYT}`), `SUMPRODUCT((B${SY0}:B${SY1}="Yes")*K${SY0}:K${SY1})`, '#,##0.0');
calc(sy.getCell(`O${SYT}`), `SUM(O${SY0}:O${SY1})`, '#,##0.0');
calc(sy.getCell(`P${SYT}`), `SUM(P${SY0}:P${SY1})`, '#,##0.0');
'ABCDEFGHIJKLMNOPQ'.split('').forEach(c => { const x = sy.getCell(`${c}${SYT}`); x.fill = totFill; x.border = box; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); });
const SYS2 = SYT + 2;
const cpl = [
  ['Total cooling load – included systems (kW thermal)', SYS.map((s, i) => s[5] === 'cool' ? `IF(B${SY0 + i}="Yes",J${SY0 + i},0)` : null).filter(Boolean).join('+')],
  ['Cooling load (refrigeration tons, TR)', `B${SYS2}/3.517`],
  ['Electrical load density of included systems (VA/m²)', `IF($B$5=0,0,P${SYT}*1000/$B$5)`],
];
cpl.forEach(([l, f], i) => {
  const r = SYS2 + i; const a = sy.getCell(`A${r}`); a.value = l; a.border = box; a.font = { name: 'Arial', size: 10, bold: true };
  calc(sy.getCell(`B${r}`), f, '#,##0.0');
});
sy.getCell(`A${SYS2 + 4}`).value = 'Life-safety systems (fire pumps, smoke fans) are normally on a dedicated / secondary supply and excluded from the normal maximum demand – set Include? to Yes only if they share the normal supply. This sheet overlaps the Building Calculator – include only one of them on the Summary.';
sy.getCell(`A${SYS2 + 4}`).font = { name: 'Arial', size: 9, italic: true };
sy.getColumn('A').width = 52; sy.getColumn('B').width = 9; sy.getColumn('C').width = 26; sy.getColumn('D').width = 10; sy.getColumn('E').width = 10;
'FGHIJK'.split('').forEach(c => sy.getColumn(c).width = 12); sy.getColumn('L').width = 38; 'MNOP'.split('').forEach(c => sy.getColumn(c).width = 10); sy.getColumn('Q').width = 46;
sy.views = [{ state: 'frozen', xSplit: 1, ySplit: 8 }];
const SYQ = "'Building Services Systems'";

