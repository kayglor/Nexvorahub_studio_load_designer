// ---------------- 17. External & Street Lighting ----------------
const ex = wb.addWorksheet('External Lighting');
title(ex, 'External, Street and Public Realm Lighting – Load Estimate', 'Street lighting (BS 5489-1 / BS EN 13201), outdoor car parks and open areas (BS EN 12464-2), parks, gardens and sports areas (BS EN 12193). Linear routes use column spacing; areas use W/m².');
legend(ex, 3);
header(ex, 5, ['Lighting application', 'Basis', 'Lighting class', 'Design criterion', 'Typical mounting height (m)', 'Typical column spacing (m)', 'Typical LED luminaire (W)', 'Area density (W/m²)', 'Notes']);
ex.getRow(5).height = 42;
const EXD = [
  ['Street – main traffic route', 'Linear', 'M3', '1.0 cd/m² avg luminance', 10, 36, 110, 0, 'BS EN 13201-2; staggered or opposite on wide roads'],
  ['Street – distributor road', 'Linear', 'M4 / M5', '0.75 – 0.5 cd/m²', 8, 32, 65, 0, ''],
  ['Street – residential road', 'Linear', 'P3 / P4', '7.5 – 5 lux avg', 6, 28, 30, 0, 'BS 5489-1 subsidiary roads'],
  ['Footpath / cycleway', 'Linear', 'P4 – P6', '5 – 2 lux avg', 5, 22, 20, 0, ''],
  ['Parks & gardens – paths (columns / bollards)', 'Linear', 'P5 / P6', '3 – 2 lux avg', 4, 15, 15, 0, 'Low-level bollards at 6–10 m spacing'],
  ['Car park – outdoor', 'Area', 'EN 12464-2', '10 – 20 lux', 8, 28, 60, 1.0, 'Light traffic 5 lux; medium 10 lux; heavy 20 lux'],
  ['Open space / public plaza', 'Area', 'EN 12464-2', '10 – 20 lux', 6, 20, 45, 1.5, 'Feature lighting extra'],
  ['Parks & gardens – amenity areas', 'Area', 'EN 12464-2', '5 – 10 lux', 4, 15, 20, 0.4, 'Add irrigation pumps / water features on Specific Loads'],
  ['Sports pitch / MUGA – recreational', 'Area', 'EN 12193 Class III', '75 – 200 lux', 12, 0, 1000, 5, 'Competition pitches on Specific Loads'],
  ['Façade / feature / signage lighting', 'Area', '–', 'Design-led', 0, 0, 30, 2, 'Per m² of lit façade'],
  ['Security / perimeter lighting', 'Linear', 'EN 12464-2', '5 – 20 lux', 6, 25, 40, 0, ''],
];
const ED0 = 6, ED1 = ED0 + EXD.length - 1;
EXD.forEach((row, i) => {
  const r = ED0 + i;
  row.forEach((v, j) => {
    const c = ex.getCell(r, j + 1);
    if (j >= 4 && j <= 7) { input(c, v, j === 7 ? '0.0' : '0'); c.alignment = { horizontal: 'center' }; }
    else { c.value = v; c.border = box; c.font = { name: 'Arial', size: j === 8 ? 9 : 10, italic: j === 8 }; }
  });
});
const EXQ = "'External Lighting'";
const EX_LIST = `${EXQ}!$A$${ED0}:$A$${ED1}`;
const EXL = (col) => `${EXQ}!$${col}$${ED0}:$${col}$${ED1}`;

const EIr = ED1 + 2;
ex.getCell(`A${EIr}`).value = 'Annual operating hours (dusk to dawn)'; ex.getCell(`A${EIr}`).border = box; ex.getCell(`A${EIr}`).font = FONT;
input(ex.getCell(`C${EIr}`), 4100, '#,##0');
ex.getCell(`D${EIr}`).value = 'UK ≈ 4,100 h/yr all-night; part-night dimming reduces energy'; ex.getCell(`D${EIr}`).font = { name: 'Arial', size: 9, italic: true };
ex.getCell(`A${EIr + 1}`).value = 'Lighting diversity (simultaneous operation)'; ex.getCell(`A${EIr + 1}`).border = box; ex.getCell(`A${EIr + 1}`).font = FONT;
input(ex.getCell(`C${EIr + 1}`), 1.0, '0.00');
ex.getCell(`D${EIr + 1}`).value = 'External lighting normally switched together – no diversity'; ex.getCell(`D${EIr + 1}`).font = { name: 'Arial', size: 9, italic: true };
const HRS = `$C$${EIr}`, EDIV = `$C$${EIr + 1}`;

const EH = EIr + 3;
header(ex, EH, ['Ref', 'Location / road name', 'Lighting application', 'Basis', 'Route length (m)', 'Sides lit (1/2)', 'Spacing override (m)', 'Area (m²)', 'Luminaire W override', 'Qty override', 'Design spacing (m)', 'Luminaire (W)', 'Quantity', 'Connected (kW)', 'Demand (kVA)', 'Annual energy (kWh)', 'Notes']);
ex.getRow(EH).height = 42;
const E0 = EH + 1, E1 = E0 + 29;
const exEx = [
  ['EX01', 'Spine road (600 m)', 'Street – distributor road', 600, 1],
  ['EX02', 'Residential streets', 'Street – residential road', 1400, 1],
  ['EX03', 'Footpaths / cycleway', 'Footpath / cycleway', 900, 1],
  ['EX04', 'Visitor car park', 'Car park – outdoor', null, null, 3500],
  ['EX05', 'Central park', 'Parks & gardens – amenity areas', null, null, 12000],
  ['EX06', 'Park paths', 'Parks & gardens – paths (columns / bollards)', 1100, 1],
  ['EX07', 'Town square', 'Open space / public plaza', null, null, 2500],
];
for (let r = E0; r <= E1; r++) {
  const e = exEx[r - E0] || [];
  input(ex.getCell(`A${r}`), e[0] || null); input(ex.getCell(`B${r}`), e[1] || null); input(ex.getCell(`C${r}`), e[2] || null);
  ex.getCell(`C${r}`).dataValidation = { type: 'list', allowBlank: true, formulae: [EX_LIST], showErrorMessage: true, error: 'Choose an application from the table above' };
  const idx = `MATCH(C${r},${EX_LIST},0)`;
  calc(ex.getCell(`D${r}`), `IF(C${r}="","",INDEX(${EXL('B')},${idx}))`, null, GREEN);
  input(ex.getCell(`E${r}`), e[3] !== undefined ? e[3] : null, '#,##0');
  input(ex.getCell(`F${r}`), e[4] !== undefined ? e[4] : null, '0');
  ex.getCell(`F${r}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['"1,2"'] };
  input(ex.getCell(`G${r}`), null, '0');
  input(ex.getCell(`H${r}`), e[5] !== undefined ? e[5] : null, '#,##0');
  input(ex.getCell(`I${r}`), null, '0');
  input(ex.getCell(`J${r}`), null, '0');
  calc(ex.getCell(`K${r}`), `IF(C${r}="","",IF(G${r}<>"",G${r},INDEX(${EXL('F')},${idx})))`, '0');
  calc(ex.getCell(`L${r}`), `IF(C${r}="","",IF(I${r}<>"",I${r},INDEX(${EXL('G')},${idx})))`, '0');
  calc(ex.getCell(`M${r}`), `IF(C${r}="",0,IF(J${r}<>"",J${r},IF(D${r}="Linear",IF(OR(N(E${r})=0,N(K${r})=0),0,(ROUNDUP(E${r}/K${r},0)+1)*IF(F${r}="",1,F${r})),IF(N(L${r})=0,0,ROUNDUP(N(H${r})*INDEX(${EXL('H')},${idx})/L${r},0)))))`, '#,##0;-#,##0;"-"');
  calc(ex.getCell(`N${r}`), `M${r}*N(L${r})/1000`, '#,##0.00;-#,##0.00;"-"');
  calc(ex.getCell(`O${r}`), `IF(${PF}=0,0,N${r}*${EDIV}/${PF})`, '#,##0.00;-#,##0.00;"-"');
  calc(ex.getCell(`P${r}`), `N${r}*${HRS}`, '#,##0;-#,##0;"-"');
  input(ex.getCell(`Q${r}`), null);
}
const ET = E1 + 1;
ex.getCell(`A${ET}`).value = 'TOTAL – external lighting'; ex.mergeCells(`A${ET}:L${ET}`);
['M', 'N', 'O', 'P'].forEach(c => calc(ex.getCell(`${c}${ET}`), `SUM(${c}${E0}:${c}${E1})`, c === 'M' || c === 'P' ? '#,##0' : '#,##0.00'));
'ABCDEFGHIJKLMNOPQ'.split('').forEach(c => { const x = ex.getCell(`${c}${ET}`); x.fill = totFill; x.border = box; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); });
ex.getCell(`A${ET + 1}`).value = 'Linear quantity = (length ÷ spacing, rounded up, + 1) × sides lit. Area quantity = area × W/m² ÷ luminaire W. Street lighting adopted by the highway authority is normally on an unmetered supply – confirm whether it is included in the site maximum demand.';
ex.getCell(`A${ET + 1}`).font = { name: 'Arial', size: 9, italic: true };
ex.getColumn('A').width = 40; ex.getColumn('B').width = 24; ex.getColumn('C').width = 40; ex.getColumn('D').width = 18;
'EFGHIJKLMNOP'.split('').forEach(c => ex.getColumn(c).width = 12); ex.getColumn('I').width = 44; ex.getColumn('Q').width = 30;
const EXT_TOTAL = `${EXQ}!O${ET}`;

