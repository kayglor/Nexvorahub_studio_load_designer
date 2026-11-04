// ---------------- 6. Service Breakdown (data) ----------------
const sb = wb.addWorksheet('Service Breakdown');
title(sb, 'Load Breakdown by Service and Building Type – W/m² GIA (after diversity)', 'Indicative design allowances per engineering service. Edit any blue value. Totals feed the Building Calculator.');
legend(sb, 3);
const SB_TYPES = ['Hospital – acute', 'Hotel – 4/5 star', 'School – secondary', 'Commercial office – AC', 'Shopping mall (incl. tenants)', 'Stadium – stands / concourse', 'Restaurant (dining + electric kitchen)', 'Car park – enclosed', 'Residential – apartment block', 'Place of worship (mosque / church)'];
header(sb, 5, ['Service', ...SB_TYPES, 'Notes']);
sb.getCell('L5').value = 'Notes'; sb.getCell('L5').fill = hdrFill; sb.getCell('L5').font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
sb.getRow(5).height = 45;
const SB_SERV = [
  ['General lighting (LED)', [12, 10, 8, 8, 15, 10, 15, 3, 4, 8], 'Retail/display and clinical areas higher'],
  ['Small power (sockets / equipment)', [25, 12, 12, 15, 15, 8, 10, 1, 10, 3], 'BCO office ~15 W/m² diversified'],
  ['Cooling (chillers / VRF / DX)', [35, 25, 3, 30, 35, 10, 30, 0, 0, 25], 'Schools usually cooling to ICT/server rooms only'],
  ['Heating (heat pumps / electric)', [25, 20, 15, 25, 15, 10, 15, 0, 15, 5], 'Set to 0 where gas / district heating'],
  ['Ventilation (AHUs, extract fans)', [20, 10, 8, 12, 12, 8, 25, 8, 3, 5], 'Restaurant includes kitchen extract; car park jet fans'],
  ['Domestic hot water (electric / HP)', [5, 15, 3, 3, 2, 6, 15, 0, 8, 3], 'Hotels and restaurants high DHW demand'],
  ['Lifts and escalators', [5, 4, 1, 5, 6, 3, 0, 1, 2, 0], 'Can be replaced by a lift schedule on the Building Calculator'],
  ['Catering / kitchen equipment', [5, 8, 6, 3, 8, 15, 150, 0, 5, 1], 'All-electric kitchens; use equipment schedule at detailed design'],
  ['Specialist / process', [20, 5, 5, 0, 5, 0, 5, 0, 0, 0], 'Hospital: imaging, theatres, sterile services; mall: tenant refrigeration'],
  ['IT / comms / AV', [5, 3, 4, 7, 2, 5, 2, 0.5, 1, 3], 'Stadium: broadcast and PA (excl. big screens)'],
  ['Life safety (fire alarm, emergency lighting, smoke control)', [3, 2, 1, 2, 2, 3, 2, 1, 1, 1], 'Exclude fire pumps on dedicated supplies'],
  ['External lighting / miscellaneous', [1, 1, 1, 1, 2, 3, 2, 0.5, 1, 2], ''],
];
const SB0 = 6, SB1 = SB0 + SB_SERV.length - 1; // 6..17
const SBC = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'];
SB_SERV.forEach(([name, vals, note], i) => {
  const r = SB0 + i;
  const a = sb.getCell(`A${r}`); a.value = name; a.border = box; a.font = FONT;
  vals.forEach((v, j) => { const c = sb.getCell(`${SBC[j]}${r}`); input(c, v, '#,##0.0'); c.alignment = { horizontal: 'center' }; });
  const n = sb.getCell(`L${r}`); n.value = note; n.font = { name: 'Arial', size: 9, italic: true }; n.border = box;
});
const SBW = SB1 + 1, SBPF = SB1 + 2, SBVA = SB1 + 3, SBXL = SB1 + 4, SBRG = SB1 + 5; // 18..22
const SB_LIFT = SB0 + 6; // lifts row (17th? -> row 12)
sb.getCell(`A${SBW}`).value = 'Total (W/m²)';
sb.getCell(`A${SBPF}`).value = 'Power factor';
sb.getCell(`A${SBVA}`).value = 'TOTAL DESIGN DENSITY (VA/m²)';
sb.getCell(`A${SBXL}`).value = 'VA/m² excluding lifts (used when lift schedule selected)';
sb.getCell(`A${SBRG}`).value = 'Typical range – VA-m2 Data sheet';
const ranges = ['100–180', '90–140', '40–85', '80–140', '60–130', '60–110', '200–400', '10–20', '35–60', '20–80'];
SBC.forEach((col, j) => {
  calc(sb.getCell(`${col}${SBW}`), `SUM(${col}${SB0}:${col}${SB1})`, '#,##0.0');
  calc(sb.getCell(`${col}${SBPF}`), PF, '0.00', GREEN);
  calc(sb.getCell(`${col}${SBVA}`), `IF(${col}${SBPF}=0,0,${col}${SBW}/${col}${SBPF})`, '#,##0');
  calc(sb.getCell(`${col}${SBXL}`), `IF(${col}${SBPF}=0,0,(${col}${SBW}-${col}${SB_LIFT})/${col}${SBPF})`, '#,##0');
  const g = sb.getCell(`${col}${SBRG}`); g.value = ranges[j]; g.border = box; g.font = { name: 'Arial', size: 9, italic: true };
  [SBW, SBPF, SBVA, SBXL, SBRG].forEach(r => sb.getCell(`${col}${r}`).alignment = { horizontal: 'center' });
});
[SBW, SBVA].forEach(r => ['A', ...SBC, 'L'].forEach(c => { const x = sb.getCell(`${c}${r}`); x.fill = totFill; x.border = box; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); }));
[SBPF, SBXL, SBRG].forEach(r => { const a = sb.getCell(`A${r}`); a.border = box; a.font = FONT; });
sb.getCell(`A${SBRG + 2}`).value = 'Source: indicative allocations consistent with BSRIA BG 9 Rules of Thumb, CIBSE Guide F/K and BCO Guide; stadium values relate to the enclosed stands/concourse GIA only. Verify at detailed design.';
sb.getCell(`A${SBRG + 2}`).font = { name: 'Arial', size: 9, italic: true };
sb.getColumn(1).width = 46; SBC.forEach(c => sb.getColumn(c).width = 14); sb.getColumn('L').width = 50;
sb.views = [{ state: 'frozen', xSplit: 1, ySplit: 5 }];
const SBQ = "'Service Breakdown'";
const SB_HDR = `${SBQ}!$B$5:$K$5`;
const SB_VA = `${SBQ}!$B$${SBVA}:$K$${SBVA}`;
const SB_XL = `${SBQ}!$B$${SBXL}:$K$${SBXL}`;

// ---------------- 7. Building Calculator (floors) ----------------
const bc = wb.addWorksheet('Building Calculator');
title(bc, 'Building Calculator – Floor-by-Floor Load Estimate', 'Enter floors, basements and typical floor area; levels are generated automatically. Override any floor area or use in the yellow columns.');
legend(bc, 3);
const sec = (ws, r, text, cols) => { const a = ws.getCell(`A${r}`); a.value = text; a.font = { name: 'Arial', size: 10, bold: true }; cols.forEach(c => ws.getCell(`${c}${r}`).fill = secFill); };
sec(bc, 5, 'BUILDING INPUTS', ['A', 'B', 'C']);
const bIn = [
  [6, 'Building / block name', 'Example: Hotel Block A', null, ''],
  [7, 'Primary use (above-ground floors)', 'Hotel – 4/5 star', null, 'Select from Service Breakdown types'],
  [8, 'Floors above ground (incl. ground floor)', 8, '0', '1 – 45'],
  [9, 'Basement levels', 1, '0', '0 – 5'],
  [10, 'Typical floor area per level (m² GIA)', 1200, '#,##0', ''],
  [11, 'Basement use', 'Car park – enclosed', null, 'Select from Service Breakdown types'],
  [12, 'Lifts basis', 'Lift schedule', null, 'Density = use W/m² in Service Breakdown; Lift schedule = rows 13–15'],
  [13, 'Number of lifts', 3, '0', ''],
  [14, 'Lift motor rating each (kW)', 15, '0.0', 'Guide: 8–11 kW (630 kg, 1 m/s) to 15–25 kW (1,600 kg, 1.6–2.5 m/s)'],
  [15, 'Lift group diversity factor', 0.7, '0.00', 'Typical 1.0 single lift; 0.6–0.8 groups'],
];
bIn.forEach(([r, label, v, fmt, note]) => {
  const a = bc.getCell(`A${r}`); a.value = label; a.border = box; a.font = FONT;
  input(bc.getCell(`B${r}`), v, fmt || undefined);
  const n = bc.getCell(`C${r}`); n.value = note; n.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } };
});
bc.getCell('A16').value = 'Power factor'; bc.getCell('A16').border = box; bc.getCell('A16').font = FONT;
calc(bc.getCell('B16'), PF, '0.00', GREEN);
bc.getCell('B7').dataValidation = { type: 'list', allowBlank: false, formulae: [SB_HDR] };
bc.getCell('B11').dataValidation = { type: 'list', allowBlank: false, formulae: [SB_HDR] };
bc.getCell('B12').dataValidation = { type: 'list', allowBlank: false, formulae: ['"Density,Lift schedule"'] };
bc.getCell('B8').dataValidation = { type: 'whole', operator: 'between', formulae: [1, 45], showErrorMessage: true, error: 'Enter 1 to 45 floors' };
bc.getCell('B9').dataValidation = { type: 'whole', operator: 'between', formulae: [0, 5], showErrorMessage: true, error: 'Enter 0 to 5 basements' };

header(bc, 19, ['#', 'Level', 'Default area (m²)', 'Area override (m²)', 'Design area (m²)', 'Default use', 'Use override', 'Design use', 'Design VA/m²', 'Demand (kVA)', 'Notes']);
const F0 = 20, F1 = 69;
for (let r = F0; r <= F1; r++) {
  const n = r - F0 + 1;
  const a = bc.getCell(`A${r}`); a.value = n; a.border = box; a.font = { name: 'Arial', size: 9, color: { argb: 'FF808080' } }; a.alignment = { horizontal: 'center' };
  calc(bc.getCell(`B${r}`), `IF(A${r}>$B$9+$B$8,"",IF(A${r}<=$B$9,"Basement "&($B$9-A${r}+1),IF(A${r}=$B$9+1,"Ground","Level "&(A${r}-$B$9-1))))`);
  calc(bc.getCell(`C${r}`), `IF(B${r}="","",$B$10)`, '#,##0');
  input(bc.getCell(`D${r}`), null, '#,##0');
  calc(bc.getCell(`E${r}`), `IF(B${r}="",0,IF(D${r}<>"",D${r},C${r}))`, '#,##0;-#,##0;"-"');
  calc(bc.getCell(`F${r}`), `IF(B${r}="","",IF(A${r}<=$B$9,$B$11,$B$7))`);
  input(bc.getCell(`G${r}`), null);
  bc.getCell(`G${r}`).dataValidation = { type: 'list', allowBlank: true, formulae: [SB_HDR] };
  calc(bc.getCell(`H${r}`), `IF(B${r}="","",IF(G${r}<>"",G${r},F${r}))`);
  calc(bc.getCell(`I${r}`), `IF(H${r}="","",IF($B$12="Lift schedule",INDEX(${SB_XL},MATCH(H${r},${SB_HDR},0)),INDEX(${SB_VA},MATCH(H${r},${SB_HDR},0)))*INDEX($M$8:$V$8,MATCH(H${r},$M$6:$V$6,0)))`, '#,##0', GREEN);
  calc(bc.getCell(`J${r}`), `IF(H${r}="",0,E${r}*I${r}/1000)`, '#,##0.0;-#,##0.0;"-"');
  input(bc.getCell(`K${r}`), null);
}
// example overrides: ground floor is restaurant / lobby
bc.getCell('G22').value = 'Restaurant (dining + electric kitchen)';
bc.getCell('D22').value = 600; bc.getCell('K22').value = 'Example override – ground floor restaurant';
const FT = F1 + 1, FL = F1 + 2, FB = F1 + 3, FD = F1 + 4; // 70..73
bc.getCell(`A${FT}`).value = 'Floors total'; bc.mergeCells(`A${FT}:D${FT}`);
calc(bc.getCell(`E${FT}`), `SUM(E${F0}:E${F1})`, '#,##0');
calc(bc.getCell(`J${FT}`), `SUM(J${F0}:J${F1})`, '#,##0.0');
bc.getCell(`A${FL}`).value = 'Lifts (schedule basis): no. × kW × diversity ÷ PF'; bc.mergeCells(`A${FL}:D${FL}`);
calc(bc.getCell(`J${FL}`), `IF(AND($B$12="Lift schedule",$B$16>0),$B$13*$B$14*$B$15/$B$16,0)`, '#,##0.0;-#,##0.0;"-"');
bc.getCell(`A${FB}`).value = 'BUILDING TOTAL DEMAND'; bc.mergeCells(`A${FB}:D${FB}`);
calc(bc.getCell(`E${FB}`), `E${FT}`, '#,##0');
calc(bc.getCell(`J${FB}`), `J${FT}+J${FL}`, '#,##0.0');
bc.getCell(`A${FD}`).value = 'Overall building density (VA/m²)'; bc.mergeCells(`A${FD}:D${FD}`);
calc(bc.getCell(`J${FD}`), `IF(E${FB}=0,0,J${FB}*1000/E${FB})`, '#,##0');
[FT, FL, FB, FD].forEach(r => ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'].forEach(c => { const x = bc.getCell(`${c}${r}`); x.fill = totFill; x.border = box; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); }));

// Right-hand block: area by use and service breakdown for this building
const UC = ['M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V'];
const m5 = bc.getCell('M5'); m5.value = 'AREA BY USE (m²)'; m5.font = { name: 'Arial', size: 10, bold: true };
UC.forEach((c, j) => {
  calc(bc.getCell(`${c}6`), `${SBQ}!${SBC[j]}5`, null, GREEN);
  bc.getCell(`${c}6`).alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
  bc.getCell(`${c}6`).fill = secFill;
  calc(bc.getCell(`${c}7`), `SUMIF($H$${F0}:$H$${F1},${c}6,$E$${F0}:$E$${F1})`, '#,##0;-#,##0;"-"');
  bc.getCell(`${c}7`).alignment = { horizontal: 'center' };
});
bc.getRow(6).height = 42;
const SBG = ['Healthcare', 'Hospitality', 'Education', 'Office', 'Retail', 'Leisure & assembly', 'Hospitality', 'Car parks & external', 'Residential', 'Worship / community'];
bc.getCell('L8').value = 'Regional factor'; bc.getCell('L8').font = { name: 'Arial', size: 8, italic: true };
UC.forEach((c, j) => { calc(bc.getCell(`${c}8`), `INDEX(${MULT_TAB},MATCH("${SBG[j]}",${GROUP_LIST},0),MATCH('Project Info'!$B$35,${MULT_HDR},0))`, '0.00', GREEN); bc.getCell(`${c}8`).alignment = { horizontal: 'center' }; });
const m9 = bc.getCell('M9'); m9.value = 'BUILDING LOAD BREAKDOWN BY SERVICE'; m9.font = { name: 'Arial', size: 10, bold: true };
['M', 'N', 'O', 'P', 'Q'].forEach((c, j) => {
  const x = bc.getCell(`${c}10`); x.value = ['Service', 'Weighted W/m²', 'Demand (kW)', 'Demand (kVA)', '% of total'][j];
  x.fill = hdrFill; x.border = box; x.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } }; x.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
});
bc.getRow(10).height = 30;
const BS0 = 11, BS1 = BS0 + SB_SERV.length - 1, BST = BS1 + 1; // 11..22, total 23
SB_SERV.forEach((s, i) => {
  const r = BS0 + i, dr = SB0 + i;
  calc(bc.getCell(`M${r}`), `${SBQ}!A${dr}`, null, GREEN);
  const sp = `SUMPRODUCT(${SBQ}!$B${dr}:$K${dr},$M$7:$V$7,$M$8:$V$8)/1000`;
  const kw = dr === SB_LIFT ? `IF($B$12="Lift schedule",$B$13*$B$14*$B$15,${sp})` : sp;
  calc(bc.getCell(`O${r}`), kw, '#,##0.0;-#,##0.0;"-"');
  calc(bc.getCell(`N${r}`), `IF($E$${FB}=0,0,O${r}*1000/$E$${FB})`, '#,##0.0;-#,##0.0;"-"');
  calc(bc.getCell(`P${r}`), `IF($B$16=0,0,O${r}/$B$16)`, '#,##0.0;-#,##0.0;"-"');
  calc(bc.getCell(`Q${r}`), `IF($P$${BST}=0,0,P${r}/$P$${BST})`, '0.0%;-0.0%;"-"');
});
bc.getCell(`M${BST}`).value = 'TOTAL';
calc(bc.getCell(`N${BST}`), `SUM(N${BS0}:N${BS1})`, '#,##0.0');
calc(bc.getCell(`O${BST}`), `SUM(O${BS0}:O${BS1})`, '#,##0.0');
calc(bc.getCell(`P${BST}`), `SUM(P${BS0}:P${BS1})`, '#,##0.0');
calc(bc.getCell(`Q${BST}`), `SUM(Q${BS0}:Q${BS1})`, '0.0%');
['M', 'N', 'O', 'P', 'Q'].forEach(c => { const x = bc.getCell(`${c}${BST}`); x.fill = totFill; x.border = box; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); });
bc.getCell(`M${BST + 1}`).value = 'Check: difference vs floor schedule (should be 0)';
bc.getCell(`M${BST + 1}`).font = { name: 'Arial', size: 9, italic: true };
calc(bc.getCell(`P${BST + 1}`), `ROUND(P${BST}-J${FB},3)`, '#,##0.0');
bc.getColumn('A').width = 34; bc.getColumn('B').width = 30; bc.getColumn('C').width = 12; bc.getColumn('D').width = 12; bc.getColumn('E').width = 12;
bc.getColumn('F').width = 30; bc.getColumn('G').width = 30; bc.getColumn('H').width = 30; bc.getColumn('I').width = 10; bc.getColumn('J').width = 12; bc.getColumn('K').width = 30;
bc.getColumn('L').width = 3; bc.getColumn('M').width = 44; ['N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V'].forEach(c => bc.getColumn(c).width = 14);
const BCQ = "'Building Calculator'";

// ---------------- 8. IET On-Site Guide diversity ----------------
const ie = wb.addWorksheet('IET Diversity');
title(ie, 'IET (formerly IEE) On-Site Guide – Appendix A Allowances for Diversity', 'Current demand after diversity for final circuits of small installations. Based on OSG Table A2 (BS 7671:2018+A2). Verify against your current OSG edition.');
legend(ie, 3);
const ieIn = [
  [5, 'Type of premises', 'Household', 'Household / Small shop, store, office / Small hotel, guest house'],
  [7, 'Cooker control unit includes a socket-outlet? (households)', 'Yes', 'Adds 5 A to cooking demand'],
  [8, 'Proposed supply / main fuse rating (A)', 100, 'e.g. 60, 80 or 100 A single-phase service'],
  [9, 'Supply voltage (V)', 230, ''],
];
ieIn.forEach(([r, label, v, note]) => {
  const a = ie.getCell(`A${r}`); a.value = label; a.border = box; a.font = FONT;
  input(ie.getCell(`B${r}`), v);
  const n = ie.getCell(`C${r}`); n.value = note; n.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } };
});
const PREM = ['Household', 'Small shop, store, office', 'Small hotel, guest house'];
ie.getCell('B5').dataValidation = { type: 'list', allowBlank: false, formulae: [`"${PREM.join(',')}"`] };
ie.getCell('B7').dataValidation = { type: 'list', allowBlank: false, formulae: ['"Yes,No"'] };
ie.getCell('A6').value = 'Premises index (formula)'; ie.getCell('A6').border = box; ie.getCell('A6').font = FONT;
calc(ie.getCell('B6'), `MATCH(B5,{"${PREM.join('","')}"},0)`, '0');
header(ie, 11, ['No.', 'Purpose of final circuit', 'Largest item / circuit (A)', '2nd largest (A)', 'All remaining (A)', 'Total connected (A)', 'Diversity allowance applied', 'Demand after diversity (A)']);
const p = '$B$6';
const IE = [
  ['1', 'Lighting', '66% of total', '90% of total', '75% of total', `F{r}*CHOOSE(${p},0.66,0.9,0.75)`, [0, 0, 6]],
  ['2', 'Heating and power (other than 3–10)', '100% up to 10 A + 50% of remainder', '100% largest + 75% of remainder', '100% largest + 80% 2nd + 60% of remainder', `CHOOSE(${p},MIN(F{r},10)+0.5*MAX(F{r}-10,0),C{r}+0.75*(D{r}+E{r}),C{r}+0.8*D{r}+0.6*E{r})`, [0, 0, 0]],
  ['3', 'Cooking appliances', '10 A + 30% of remainder (+5 A if CCU socket)', '100% largest + 80% 2nd + 60% of remainder', '100% largest + 80% 2nd + 60% of remainder', `IF(F{r}=0,0,CHOOSE(${p},MIN(F{r},10)+0.3*MAX(F{r}-10,0)+IF($B$7="Yes",5,0),C{r}+0.8*D{r}+0.6*E{r},C{r}+0.8*D{r}+0.6*E{r}))`, [32, 0, 0]],
  ['4', 'Motors (excluding lift motors)', 'Not applicable – 100% taken', '100% largest + 80% 2nd + 60% of remainder', '100% largest + 50% of remainder', `CHOOSE(${p},F{r},C{r}+0.8*D{r}+0.6*E{r},C{r}+0.5*(D{r}+E{r}))`, [0, 0, 0]],
  ['5', 'Water heaters – instantaneous (e.g. showers)', '100% largest + 100% 2nd + 25% of remainder', '100% largest + 100% 2nd + 25% of remainder', '100% largest + 100% 2nd + 25% of remainder', `C{r}+D{r}+0.25*E{r}`, [40, 0, 0]],
  ['6', 'Water heaters – thermostatically controlled', 'No diversity', 'No diversity', 'No diversity', `F{r}`, [0, 0, 0]],
  ['7', 'Floor warming installations', 'No diversity', 'No diversity', 'No diversity', `F{r}`, [0, 0, 0]],
  ['8', 'Thermal storage space heating', 'No diversity', 'No diversity', 'No diversity', `F{r}`, [0, 0, 0]],
  ['9', 'Standard final circuits (OSG Appendix H – rings / radials)', '100% largest circuit + 40% of every other', '100% largest circuit + 50% of every other', '100% largest circuit + 50% of every other', `CHOOSE(${p},C{r}+0.4*(D{r}+E{r}),C{r}+0.5*(D{r}+E{r}),C{r}+0.5*(D{r}+E{r}))`, [32, 32, 20]],
  ['10', 'Other socket-outlets and stationary equipment', '100% largest point + 40% of every other', '100% largest point + 75% of every other', '100% largest point (main rooms) + 75% of every other', `CHOOSE(${p},C{r}+0.4*(D{r}+E{r}),C{r}+0.75*(D{r}+E{r}),C{r}+0.75*(D{r}+E{r}))`, [0, 0, 0]],
];
const I0 = 12, I1 = I0 + IE.length - 1, ITOT = I1 + 1; // 12..21, 22
IE.forEach(([no, name, t1, t2, t3, f, ex], i) => {
  const r = I0 + i;
  const a = ie.getCell(`A${r}`); a.value = no; a.border = box; a.font = FONT; a.alignment = { horizontal: 'center' };
  const b = ie.getCell(`B${r}`); b.value = name; b.border = box; b.font = FONT; b.alignment = { wrapText: true };
  input(ie.getCell(`C${r}`), ex[0], '#,##0.0'); input(ie.getCell(`D${r}`), ex[1], '#,##0.0'); input(ie.getCell(`E${r}`), ex[2], '#,##0.0');
  calc(ie.getCell(`F${r}`), `SUM(C${r}:E${r})`, '#,##0.0;-#,##0.0;"-"');
  calc(ie.getCell(`G${r}`), `CHOOSE(${p},"${t1}","${t2}","${t3}")`);
  ie.getCell(`G${r}`).alignment = { wrapText: true }; ie.getCell(`G${r}`).font = { name: 'Arial', size: 9 };
  calc(ie.getCell(`H${r}`), f.replace(/\{r\}/g, r), '#,##0.0;-#,##0.0;"-"');
});
ie.getCell(`B${ITOT}`).value = 'TOTAL';
calc(ie.getCell(`F${ITOT}`), `SUM(F${I0}:F${I1})`, '#,##0.0');
calc(ie.getCell(`H${ITOT}`), `SUM(H${I0}:H${I1})`, '#,##0.0');
['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].forEach(c => { const x = ie.getCell(`${c}${ITOT}`); x.fill = totFill; x.border = box; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); });
const ieRes = [
  [ITOT + 2, 'Overall diversity achieved', `IF(F${ITOT}=0,0,H${ITOT}/F${ITOT})`, '0%'],
  [ITOT + 3, 'Demand after diversity (kVA, single-phase)', `H${ITOT}*$B$9/1000`, '#,##0.0'],
  [ITOT + 4, 'Per phase if balanced over 3 phases (A)', `H${ITOT}/3`, '#,##0.0'],
  [ITOT + 5, 'Check against proposed supply', `IF(H${ITOT}<=$B$8,"OK – within "&$B$8&" A supply","EXCEEDS "&$B$8&" A – review supply / load")`, null],
];
ieRes.forEach(([r, label, f, fmt]) => {
  const a = ie.getCell(`B${r}`); a.value = label; a.border = box; a.font = { name: 'Arial', size: 10, bold: true };
  calc(ie.getCell(`H${r}`), f, fmt || undefined);
  ie.getCell(`H${r}`).font = { name: 'Arial', size: 10, bold: true };
});
const ieNotes = [
  'How to use: enter currents in amperes (I = P ÷ V). Where the allowance is based on the total (lighting; household heating/power and cooking; rows 6–8), the split between columns C–E does not matter.',
  'For row 9 enter the design current of the largest circuit in C, the second in D and the sum of all other circuits in E. Lift motors are subject to special consideration and are excluded.',
  'Allowances apply to small installations as described in the IET On-Site Guide. Larger or non-standard installations require assessment by the designer (BS 7671 Regulation 311.1).',
  'Example values (household): 7.4 kW cooker, 9.5 kW shower, two 32 A rings and one 20 A radial, 6 A lighting – replace with project data.',
];
ieNotes.forEach((t, i) => { const c = ie.getCell(`A${ITOT + 7 + i}`); c.value = t; c.font = { name: 'Arial', size: 9, italic: true }; });
ie.getColumn('A').width = 46; ie.getColumn('B').width = 44; ie.getColumn('C').width = 14; ie.getColumn('D').width = 12; ie.getColumn('E').width = 12;
ie.getColumn('F').width = 12; ie.getColumn('G').width = 44; ie.getColumn('H').width = 26;
for (let r = I0; r <= I1; r++) ie.getRow(r).height = 28;

