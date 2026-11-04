let CC_TOTAL, CCQ;
{
// ---------------- 11a. Code Load Calc (NEC Article 220 / CEC Section 8) ----------------
const cc = wb.addWorksheet('Code Load Calc');
title(cc, 'Code Service Load Calculation – NEC Article 220 (USA) / CEC Section 8 (Canada)', 'Standard-method calculated load for services and feeders. NEC: Table 220.12 unit loads, Table 220.42 lighting demand factors, 220.44/220.47 receptacles, 220.56 kitchens, 220.60 non-coincident loads, 430.24 motors, 620.14 elevators, 625.42 EVSE. CEC: basic load per Rules 8-204 to 8-210 (verify values).');
legend(cc, 3);
const ccr = (r, label, kind, v, fmt, note, color) => {
  const a = cc.getCell(`A${r}`); a.value = label; a.border = box; a.font = FONT; cc.mergeCells(`A${r}:B${r}`);
  if (kind === 'in') input(cc.getCell(`C${r}`), v, fmt); else calc(cc.getCell(`C${r}`), v, fmt, color);
  if (note) { const n = cc.getCell(`D${r}`); n.value = note; n.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } }; }
};
ccr(5, 'Code method (auto from project density region)', 'f', `IF('Project Info'!$B$36="CEC","CEC","NEC")`, null, 'From code family on Project Info: CEC for Canada, NEC otherwise', GREEN);
ccr(6, 'Method override', 'in', null, null, 'Leave blank, or choose NEC / CEC');
cc.getCell('C6').dataValidation = { type: 'list', allowBlank: true, formulae: ['"NEC,CEC"'] };
ccr(7, 'APPLIED METHOD', 'f', 'IF(C6<>"",C6,C5)', null, '');
cc.getCell('C7').font = { name: 'Arial', size: 11, bold: true };
ccr(8, 'Service voltage, line–line (V)', 'f', `'Project Info'!B17`, '0', 'e.g. 208Y/120 V or 480Y/277 V (override on Project Info)', GREEN);
ccr(9, 'Phases (1 = 120/240 V split-phase, 3 = three-phase)', 'in', 3, '0', '');
cc.getCell('C9').dataValidation = { type: 'list', allowBlank: false, formulae: ['"1,3"'] };
ccr(10, 'Apply 125% to continuous loads (NEC 215.2 / 230.42; CEC 8-104)', 'in', 'Yes', null, 'Lighting (non-dwelling), water heaters, EVSE, other continuous');
cc.getCell('C10').dataValidation = { type: 'list', allowBlank: false, formulae: ['"Yes,No"'] };
const CONT = `IF($C$10="Yes",1.25,1)`;
const NEC = `$C$7="NEC"`;

sub(cc, 'A12', '1. General lighting (NEC) / basic load (CEC) by occupancy');
header(cc, 13, ['Occupancy', 'Area (m²)', 'Area (ft²)', 'Unit load (NEC VA/ft² | CEC W/m²)', 'Connected (VA)', 'CEC demand factor', 'Demand (VA)', 'Continuous factor', 'Calculated (VA)', 'Demand-factor rule applied']);
cc.getRow(13).height = 42;
const G0 = 14, G1 = 25;
const gEx = [['Office', 4200], ['Retail', 850], ['Restaurant', 150], ['Parking garage', 1500], ['Manufacturing facility', 3000], ['Warehouse', 5000]];
for (let r = G0; r <= G1; r++) {
  const e = gEx[r - G0] || [];
  input(cc.getCell(`A${r}`), e[0] || null); cc.getCell(`A${r}`).dataValidation = { type: 'list', allowBlank: true, formulae: [OCC_LIST] };
  input(cc.getCell(`B${r}`), e[1] !== undefined ? e[1] : null, '#,##0');
  calc(cc.getCell(`C${r}`), `IF(B${r}="","",B${r}*10.764)`, '#,##0');
  const om = `MATCH(A${r},${OCC_LIST},0)`;
  calc(cc.getCell(`D${r}`), `IF(A${r}="","",IF(${NEC},INDEX(${OCC_NECF},${om}),INDEX(${OCC_CEC},${om})))`, '0.0', GREEN);
  calc(cc.getCell(`E${r}`), `IF(OR(A${r}="",B${r}=""),0,IF(${NEC},C${r}*D${r},B${r}*D${r}))`, '#,##0;-#,##0;"-"');
  input(cc.getCell(`F${r}`), null, '0%');
  const V = `E${r}`;
  calc(cc.getCell(`G${r}`), `IF(${V}=0,0,IF(${NEC},IF(A${r}="Dwelling unit",MIN(${V},3000)+0.35*MAX(MIN(${V},120000)-3000,0)+0.25*MAX(${V}-120000,0),IF(A${r}="Hotel / motel",0.6*MIN(${V},20000)+0.5*MAX(MIN(${V},100000)-20000,0)+0.35*MAX(${V}-100000,0),IF(A${r}="Hospital",0.4*MIN(${V},50000)+0.2*MAX(${V}-50000,0),IF(A${r}="Warehouse",MIN(${V},12500)+0.5*MAX(${V}-12500,0),${V})))),${V}*IF(F${r}="",1,F${r})))`, '#,##0;-#,##0;"-"');
  calc(cc.getCell(`H${r}`), `IF(OR(A${r}="",A${r}="Dwelling unit"),1,${CONT})`, '0.00');
  calc(cc.getCell(`I${r}`), `G${r}*H${r}`, '#,##0;-#,##0;"-"');
  calc(cc.getCell(`J${r}`), `IF(A${r}="","",IF(${NEC},INDEX(${RDQ}!$D$${OC0}:$D$${OC1},${om}),"CEC demand factor as entered (default 100%) – verify Rules 8-204 to 8-210"))`);
}
const GT = G1 + 1;
cc.getCell(`A${GT}`).value = 'Subtotal – general lighting / basic load';
['B', 'C', 'E', 'G', 'I'].forEach(c => calc(cc.getCell(`${c}${GT}`), `SUM(${c}${G0}:${c}${G1})`, '#,##0'));
'ABCDEFGHIJ'.split('').forEach(c => { const x = cc.getCell(`${c}${GT}`); x.fill = totFill; x.border = box; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); });
cc.getCell(`A${GT + 1}`).value = 'Enter each occupancy once (combine areas) – NEC Table 220.42 demand factors apply to the total of each occupancy type.';
cc.getCell(`A${GT + 1}`).font = { name: 'Arial', size: 9, italic: true };

const RS0 = GT + 3;
sub(cc, `A${RS0}`, '2. Receptacle loads (NEC 220.14(I), 220.14(K), 220.47) – included in CEC basic load');
ccr(RS0 + 1, 'Number of general-purpose receptacle outlets', 'in', 400, '#,##0', '180 VA per outlet (220.14(I))');
ccr(RS0 + 2, 'Office / bank floor area for 1 VA/ft² (ft²)', 'f', `SUMIF(A${G0}:A${G1},"Office",C${G0}:C${G1})`, '#,##0', '220.14(K): larger of 1 VA/ft² or 180 VA per outlet', GREEN);
ccr(RS0 + 3, 'Receptacle connected load (VA)', 'f', `IF(${NEC},MAX(C${RS0 + 1}*180,C${RS0 + 2}*1),0)`, '#,##0', '');
ccr(RS0 + 4, 'Receptacle demand (VA) – first 10 kVA at 100%, remainder 50%', 'f', `MIN(C${RS0 + 3},10000)+0.5*MAX(C${RS0 + 3}-10000,0)`, '#,##0', 'Table 220.47 (non-dwelling)');

const FX = RS0 + 6;
sub(cc, `A${FX}`, '3. Fixed appliances, HVAC, motors, elevators, EVSE');
header(cc, FX + 1, ['Load', 'Quantity', 'Rating each (kVA)', 'Connected (kVA)', 'Rule', 'Factor', 'Calculated (kVA)', 'Notes']);
cc.getRow(FX + 1).height = 30;
const F0 = FX + 2;
const FIX = [
  ['Space heating (electric / heat pump)', 1, 120, 'NEC 220.51 / 220.60; CEC 8-106(4)', 'nc_heat', ''],
  ['Air conditioning / cooling plant', 1, 180, 'NEC 220.60 / 440; CEC 8-106(4)', 'nc_cool', 'Larger of heating or cooling is taken (non-coincident)'],
  ['Commercial kitchen equipment (units)', 6, 8, 'NEC Table 220.56', 'kitchen', '1–2 units 100%, 3: 90%, 4: 80%, 5: 70%, 6+: 65%'],
  ['Water heaters (storage)', 2, 4.5, 'NEC 422.13 (continuous)', 'cont', ''],
  ['Motors – all (excl. HVAC above)', 1, 60, 'NEC 430.24', 'motor', 'Largest motor below adds 25%'],
  ['Largest motor (enter once)', 1, 15, 'NEC 430.24 – +25% of largest', 'largest', ''],
  ['Elevators (number × kVA)', 3, 13, 'NEC 620.14 demand factors', 'lift', '1: 1.00, 2: 0.95, 3: 0.90, 4: 0.85, 5: 0.82, 6: 0.79, 7: 0.77, 8: 0.75, 9: 0.73, 10+: 0.72'],
  ['EV supply equipment (EVSE)', 10, 7.2, 'NEC 220.57 / 625.42 (continuous)', 'cont', 'CEC Section 86 / 8-106(11) load management'],
  ['Other continuous loads', 1, 0, '125% continuous', 'cont', ''],
  ['Other non-continuous loads', 1, 0, '100%', 'one', ''],
];
FIX.forEach(([n, q, kva, rule, type, note], i) => {
  const r = F0 + i;
  const a = cc.getCell(`A${r}`); a.value = n; a.border = box; a.font = FONT;
  input(cc.getCell(`B${r}`), q, '0'); input(cc.getCell(`C${r}`), kva, '#,##0.0');
  calc(cc.getCell(`D${r}`), `N(B${r})*N(C${r})`, '#,##0.0;-#,##0.0;"-"');
  const e = cc.getCell(`E${r}`); e.value = rule; e.border = box; e.font = { name: 'Arial', size: 9 };
  const hR = F0, cR = F0 + 1;
  const f = {
    nc_heat: `IF(D${hR}>=D${cR},1,0)`, nc_cool: `IF(D${cR}>D${hR},1,0)`,
    kitchen: `IF(N(B${r})<=2,1,IF(B${r}=3,0.9,IF(B${r}=4,0.8,IF(B${r}=5,0.7,0.65))))`,
    cont: CONT, motor: '1', largest: '0.25',
    lift: `IF(N(B${r})<=1,1,INDEX({1,0.95,0.9,0.85,0.82,0.79,0.77,0.75,0.73,0.72},MIN(B${r},10)))`, one: '1',
  }[type];
  calc(cc.getCell(`F${r}`), f, '0.00');
  calc(cc.getCell(`G${r}`), `D${r}*F${r}`, '#,##0.0;-#,##0.0;"-"');
  const nn = cc.getCell(`H${r}`); nn.value = note; nn.border = box; nn.font = { name: 'Arial', size: 9, italic: true };
});
const F1 = F0 + FIX.length - 1, FT = F1 + 1;
cc.getCell(`A${FT}`).value = 'Subtotal – fixed loads (kVA)';
calc(cc.getCell(`G${FT}`), `SUM(G${F0}:G${F1})`, '#,##0.0');
'ABCDEFGH'.split('').forEach(c => { const x = cc.getCell(`${c}${FT}`); x.fill = totFill; x.border = box; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); });

const TT = FT + 2;
sub(cc, `A${TT}`, '4. Calculated service load');
const NEC_OCPD = [100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500, 600, 700, 800, 1000, 1200, 1600, 2000, 2500, 3000, 4000, 5000, 6000];
NEC_OCPD.forEach((v, i) => input(cc.getCell(`L${14 + i}`), v, '0'));
cc.getCell('L13').value = 'Std ratings (A) NEC 240.6(A)'; cc.getCell('L13').font = { name: 'Arial', size: 9, bold: true };
const OCPR = `$L$14:$L$${13 + NEC_OCPD.length}`;
ccr(TT + 1, 'General lighting / basic load (kVA)', 'f', `I${GT}/1000`, '#,##0.0');
ccr(TT + 2, 'Receptacles (kVA)', 'f', `C${RS0 + 4}/1000`, '#,##0.0');
ccr(TT + 3, 'Fixed loads (kVA)', 'f', `G${FT}`, '#,##0.0');
ccr(TT + 4, 'TOTAL CALCULATED LOAD (kVA)', 'f', `SUM(C${TT + 1}:C${TT + 3})`, '#,##0.0');
cc.getCell(`C${TT + 4}`).font = { name: 'Arial', size: 11, bold: true };
ccr(TT + 5, 'Calculated service current (A)', 'f', `IF(N(C8)=0,0,IF(C9=1,C${TT + 4}*1000/C8,C${TT + 4}*1000/(SQRT(3)*C8)))`, '#,##0', 'Split-phase uses L-L voltage (240 V)');
ccr(TT + 6, 'Minimum service / main OCPD rating (A)', 'f', `IF(C${TT + 5}>MAX(${OCPR}),"> "&MAX(${OCPR})&" A – multiple services",INDEX(${OCPR},COUNTIF(${OCPR},"<"&C${TT + 5})+1))`, '0', 'Next standard rating ≥ calculated current');
ccr(TT + 7, 'Equivalent density (VA/ft² of lighting area)', 'f', `IF(C${GT}=0,0,C${TT + 4}*1000/C${GT})`, '0.00', '');
cc.getCell(`A${TT + 9}`).value = 'This is the code-prescribed minimum calculated load for sizing services and feeders. It is not a measured maximum demand; utilities may accept a lower demand based on data. NEC references per NFPA 70 (2020/2023); confirm the edition adopted by the AHJ. CEC values must be verified against CSA C22.1 in force.';
cc.getCell(`A${TT + 9}`).font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FFC00000' } }; cc.getCell(`A${TT + 9}`).alignment = { wrapText: true };
cc.mergeCells(TT + 9, 1, TT + 10, 10);
cc.getColumn('A').width = 34; cc.getColumn('B').width = 12; cc.getColumn('C').width = 14; cc.getColumn('D').width = 16; cc.getColumn('E').width = 26;
cc.getColumn('F').width = 11; cc.getColumn('G').width = 13; cc.getColumn('H').width = 13; cc.getColumn('I').width = 13; cc.getColumn('J').width = 60; cc.getColumn('L').width = 14;
CCQ = "'Code Load Calc'";
CC_TOTAL = `${CCQ}!C${TT + 4}`;


}
