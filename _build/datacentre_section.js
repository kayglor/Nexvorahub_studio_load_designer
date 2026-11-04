// ---------------- 18d. Data Centre ----------------
let DCQ, DC_MD;
{
const dc = wb.addWorksheet('Data Centre');
DCQ = "'Data Centre'";
title(dc, 'Data Centre Electrical Load Design – IT Capacity, Resilience, UPS, Generation, Cooling and Distribution', 'Size by white-space area, rack count or IT capacity. Resilience follows the selected Tier / Rated / Availability class. Standards: Uptime Institute Tier Standard, ANSI/TIA-942-B, EN 50600-2-2, BICSI 002, ASHRAE TC 9.9, IEC 62040-3, ISO 8528-1 (DCP), ISO/IEC 30134-2 (PUE), NFPA 75, NEC Art. 645.');
legend(dc, 3);
const dsec = (r, t) => { const c = dc.getCell(`A${r}`); c.value = t; c.font = { name: 'Arial', size: 10, bold: true }; ['A', 'B', 'C', 'D'].forEach(x => dc.getCell(`${x}${r}`).fill = secFill); };
const drow = (r, label, kind, v, fmt, note, color) => {
  const a = dc.getCell(`A${r}`); a.value = label; a.border = box; a.font = FONT;
  if (kind === 'in') input(dc.getCell(`B${r}`), v, fmt); else calc(dc.getCell(`B${r}`), v, fmt, color);
  dc.getCell(`B${r}`).alignment = { wrapText: true, horizontal: 'left' };
  if (note) { const n = dc.getCell(`C${r}`); n.value = note; n.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } }; }
  return `B${r}`;
};
const RED = (red, n) => `CHOOSE(MATCH(${red},{"N","N+1","N+2","2N","2(N+1)"},0),${n},${n}+1,${n}+2,2*${n},2*(${n}+1))`;
const REDL = '"N,N+1,N+2,2N,2(N+1)"';
// Tier reference table (columns F..M)
const TIERS = [
  ['Tier I – Basic capacity', 'Rated-1', 'Class 1', '99.671%', 'N', 'N', 'N', 'N', 1],
  ['Tier II – Redundant capacity components', 'Rated-2', 'Class 2', '99.741%', 'N+1', 'N+1', 'N+1', 'N', 1],
  ['Tier III – Concurrently maintainable', 'Rated-3', 'Class 3', '99.982%', 'N+1', 'N+1', 'N+1', 'N+1', 2],
  ['Tier IV – Fault tolerant', 'Rated-4', 'Class 4', '99.995%', '2N', '2N', '2N', '2N', 2],
];
['Uptime Tier', 'TIA-942', 'EN 50600 class', 'Availability (commonly quoted)', 'UPS', 'Generators', 'Cooling', 'Transformers', 'Distribution paths'].forEach((h, i) => { const x = dc.getCell(5, 6 + i); x.value = h; x.fill = hdrFill; x.border = box; x.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFFFF' } }; x.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' }; });
dc.getRow(5).height = 32;
TIERS.forEach((t, i) => t.forEach((v, j) => { const c = dc.getCell(6 + i, 6 + j); if (j >= 4) input(c, v); else { c.value = v; c.border = box; c.font = { name: 'Arial', size: 9 }; } c.alignment = { horizontal: j === 0 ? 'left' : 'center', wrapText: true }; }));
const TL = `$F$6:$F$9`;
const tierCol = (k) => `INDEX($${'FGHIJKLMN'[k]}$6:$${'FGHIJKLMN'[k]}$9,MATCH($B$6,${TL},0))`;

dsec(5, 'A. DESIGN BASIS');
drow(6, 'Resilience level (Uptime Tier)', 'in', 'Tier III – Concurrently maintainable', null, 'Sets default redundancy; TIA-942 / EN 50600 equivalents in table →');
dc.getCell('B6').dataValidation = { type: 'list', allowBlank: false, formulae: [TL] };
drow(7, 'Sizing method', 'in', 'By white-space area', null, '');
dc.getCell('B7').dataValidation = { type: 'list', allowBlank: false, formulae: ['"By white-space area,By rack count,By IT capacity"'] };
drow(8, 'White space (data hall) area (m²)', 'in', 1000, '#,##0', 'Used for "By white-space area"');
drow(9, 'Area per rack incl. aisles (m²/rack)', 'in', 2.5, '0.0', 'Typical 2.3–3.0 m² (hot/cold aisle containment)');
drow(10, 'Number of racks (for "By rack count")', 'in', 400, '#,##0', '');
drow(11, 'IT capacity (kW) (for "By IT capacity")', 'in', 4000, '#,##0', '');
drow(12, 'Average rack density (kW/rack)', 'in', 10, '0.0', 'Enterprise 5–10; colocation 8–15; HPC / AI 30–100+ (liquid cooled)');
drow(13, 'Day-1 deployment (% of ultimate IT)', 'in', 0.5, '0%', 'Phased fit-out');
drow(14, 'Support / office / NOC area (m²)', 'in', 600, '#,##0', '');
drow(15, 'Support area load density (VA/m²)', 'in', 80, '#,##0', '');
drow(16, 'Average IT utilisation (energy calc)', 'in', 0.6, '0%', 'Annual average of design IT load');

dsec(18, 'B. IT LOAD');
drow(19, 'Racks (derived)', 'f', `IF(B7="By white-space area",ROUNDDOWN(B8/B9,0),IF(B7="By rack count",B10,ROUNDUP(B11/B12,0)))`, '#,##0');
drow(20, 'ULTIMATE IT CRITICAL LOAD (kW)', 'f', `IF(B7="By IT capacity",B11,B19*B12)`, '#,##0');
drow(21, 'Day-1 IT load (kW)', 'f', 'B20*B13', '#,##0');
drow(22, 'IT power density (W/m² white space)', 'f', `IF(B7="By white-space area",B20*1000/B8,B20*1000/(B19*B9))`, '#,##0');
drow(23, 'White space area (derived, m²)', 'f', `IF(B7="By white-space area",B8,B19*B9)`, '#,##0');
dc.getCell('B20').font = { name: 'Arial', size: 11, bold: true };

dsec(25, 'C. RESILIENCE (default from Tier – override as required)');
const redRow = (r, label, k) => { drow(r, label, 'f', `IF(C${r}<>"",C${r},${tierCol(k)})`, null, '', GREEN); input(dc.getCell(`C${r}`), null); dc.getCell(`C${r}`).dataValidation = { type: 'list', allowBlank: true, formulae: [REDL] }; dc.getCell(`D${r}`).value = '← override'; dc.getCell(`D${r}`).font = { name: 'Arial', size: 8, italic: true }; };
redRow(26, 'UPS redundancy', 4); redRow(27, 'Generator redundancy', 5); redRow(28, 'Cooling redundancy', 6); redRow(29, 'Transformer redundancy', 7);
drow(30, 'Power distribution paths to racks (A/B)', 'f', tierCol(8), '0', 'Dual-corded IT on A/B feeds for Tier III / IV', GREEN);

dsec(32, 'D. COOLING AND MECHANICAL');
const COOL = [['Air-cooled chillers with free cooling', 0.3], ['Water-cooled chillers + cooling towers', 0.22], ['DX CRAC units', 0.45], ['Indirect evaporative / adiabatic', 0.15], ['Direct-to-chip liquid cooling + dry coolers', 0.1], ['Immersion cooling', 0.06]];
COOL.forEach(([n, v], i) => { const a = dc.getCell(12 + i, 6); a.value = n; a.border = box; a.font = { name: 'Arial', size: 9 }; input(dc.getCell(12 + i, 7), v, '0.00'); });
const ch = dc.getCell('F11'); ch.value = 'Cooling technology'; ch.fill = hdrFill; ch.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFFFF' } }; ch.border = box;
const ch2 = dc.getCell('G11'); ch2.value = 'kWe per kWth (peak)'; ch2.fill = hdrFill; ch2.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFFFFFFF' } }; ch2.border = box; ch2.alignment = { wrapText: true };
const CL = `$F$12:$F$${11 + COOL.length}`, CV = `$G$12:$G$${11 + COOL.length}`;
drow(33, 'Cooling technology', 'in', 'Air-cooled chillers with free cooling', null, '');
dc.getCell('B33').dataValidation = { type: 'list', allowBlank: false, formulae: [CL] };
drow(34, 'Cooling plant power (kWe per kW heat, design day)', 'f', `IF(C34<>"",C34,INDEX(${CV},MATCH(B33,${CL},0)))`, '0.00', 'Override in C34', GREEN); input(dc.getCell('C34'), null, '0.00');
drow(35, 'Air movement – CRAH / fan walls (kWe per kW heat)', 'in', 0.05, '0.00', 'EC fans; 0.03–0.08');
drow(36, 'Design ambient / ASHRAE TC 9.9 inlet class', 'in', 'A1 recommended 18–27 °C', null, 'ASHRAE TC 9.9 thermal guidelines');

dsec(38, 'E. UPS AND BATTERIES');
drow(39, 'UPS topology', 'in', 'Double conversion (VFI), modular', null, 'IEC 62040-3');
drow(40, 'UPS efficiency at design load', 'in', 0.96, '0.0%', 'Eco-mode up to 0.99');
drow(41, 'UPS module rating (kW)', 'in', 500, '#,##0', '');
drow(42, 'Maximum UPS module loading', 'in', 0.8, '0%', '');
drow(43, 'Battery recharge allowance (% of IT load)', 'in', 0.1, '0%', 'Typically 10–15% for 1–2 h after an outage');
drow(44, 'Battery technology', 'in', 'Lithium-ion', null, '');
dc.getCell('B44').dataValidation = { type: 'list', allowBlank: false, formulae: ['"Lithium-ion,VRLA"'] };
drow(45, 'Battery autonomy at full IT load (minutes)', 'in', 5, '0', 'Li-ion 5–10 min; VRLA 10–15 min (generator start and transfer)');

dsec(47, 'F. FACILITY LOAD BUILD-UP (ultimate)');
drow(48, 'IT critical load (kW)', 'f', 'B20', '#,##0');
drow(49, 'UPS losses (kW)', 'f', 'B20/B40-B20', '#,##0');
drow(50, 'Battery recharge (kW)', 'f', 'B20*B43', '#,##0');
drow(51, 'Lighting and small power – white space at 10 W/m² (kW)', 'f', 'B23*10/1000', '#,##0.0');
drow(52, 'Support areas (kW)', 'f', `B14*B15*${PF}/1000`, '#,##0.0');
drow(53, 'Heat load to cooling (kW)', 'f', 'B48+B49+B51', '#,##0');
drow(54, 'Cooling plant (kW)', 'f', 'B53*B34', '#,##0');
drow(55, 'Air movement / CRAH fans (kW)', 'f', 'B53*B35', '#,##0');
drow(56, 'Ancillaries – BMS, security, fire, controls (% of IT)', 'in', 0.02, '0%', '');
drow(57, 'Ancillary load (kW)', 'f', 'B20*B56', '#,##0');
drow(58, 'Distribution / transformer losses (% of facility load)', 'in', 0.015, '0.0%', '');
drow(59, 'FACILITY PEAK LOAD (kW)', 'f', '(B48+B49+B50+B51+B52+B54+B55+B57)*(1+B58)', '#,##0');
drow(60, 'Design-day (peak) PUE', 'f', 'IF(B20=0,0,(B59-B50)/B20)', '0.00', 'ISO/IEC 30134-2 basis; excludes recharge');
drow(61, 'Annualised design PUE', 'in', 1.35, '0.00', 'Climate / free-cooling dependent');
drow(62, 'MAXIMUM DEMAND (kVA)', 'f', `B59/${PF}`, '#,##0');
drow(63, 'Facility load density (VA/m² of white space)', 'f', 'IF(B23=0,0,B62*1000/B23)', '#,##0');
['B59', 'B62'].forEach(a => dc.getCell(a).font = { name: 'Arial', size: 11, bold: true });

dsec(65, 'G. EQUIPMENT SIZING');
drow(66, 'UPS modules required (N)', 'f', 'ROUNDUP(B20/(B41*B42),0)', '0');
drow(67, 'UPS modules installed (with redundancy)', 'f', RED('B26', 'B66'), '0');
drow(68, 'Installed UPS capacity (kW)', 'f', 'B67*B41', '#,##0');
drow(69, 'UPS systems (independent paths)', 'f', `IF(OR(B26="2N",B26="2(N+1)"),2,1)`, '0');
drow(70, 'Battery energy per UPS system (kWh)', 'f', `B20*B45/60/(0.95*IF(B44="Lithium-ion",0.9,0.8))`, '#,##0', 'At 95% inverter efficiency; DoD 90% Li-ion / 80% VRLA');
drow(71, 'Total battery energy installed (kWh)', 'f', 'B70*B69', '#,##0');
drow(72, 'Generator unit rating (kVA)', 'in', 2500, '#,##0', 'ISO 8528-1 data-centre continuous (DCP) or prime rating');
drow(73, 'Maximum generator loading', 'in', 0.8, '0%', '');
drow(74, 'Generators required (N)', 'f', `ROUNDUP(B62/(B72*B73),0)`, '0', 'Generators carry the full facility load incl. cooling');
drow(75, 'Generators installed (with redundancy)', 'f', RED('B27', 'B74'), '0');
drow(76, 'On-site fuel autonomy (hours)', 'in', 48, '0', 'Uptime Tier: ≥ 12 h; common practice 24–72 h');
drow(77, 'Fuel storage required (litres)', 'f', 'B59*0.27*B76', '#,##0', '≈ 0.27 l/kWh diesel at site load');
drow(78, 'MV/LV transformer unit rating (kVA)', 'in', 2500, '#,##0', '');
drow(79, 'Maximum transformer loading', 'in', 0.8, '0%', '');
drow(80, 'Transformers required (N)', 'f', 'ROUNDUP(B62/(B78*B79),0)', '0');
drow(81, 'Transformers installed (with redundancy)', 'f', RED('B29', 'B80'), '0');
drow(82, 'LV switchboard rating per transformer (A)', 'f', `INDEX(${SWB},MIN(ROWS(${SWB}),COUNTIF(${SWB},"<"&B78*1000/(SQRT(3)*${VLL}))+1))`, '#,##0', 'Sized to transformer full-load current');
drow(83, 'Floor PDU / RPP rating (kVA)', 'in', 300, '#,##0', '');
drow(84, 'Floor PDUs / RPPs required (all paths)', 'f', `ROUNDUP(B20/(B83*0.8*${PF}),0)*B30`, '0', 'At 80% loading, per distribution path');
drow(85, 'Rack PDUs per rack', 'f', 'B30', '0', 'One per path (A/B)');
drow(86, 'Rack PDU current per path (A, 3-phase)', 'f', `B12*1000/(SQRT(3)*${VLL}*0.95)`, '0.0', 'Each PDU must carry the full rack load');
drow(87, 'Rack PDU / outlet specification', 'f', `IF(B86<=16*0.8,"3-ph 16 A (IEC 60309) – 11 kW",IF(B86<=32*0.8,"3-ph 32 A (IEC 60309) – 22 kW",IF(B86<=63*0.8,"3-ph 63 A (IEC 60309) – 43 kW","Busway tap-off / multiple feeds per rack")))`, null, '80% continuous loading');
drow(88, 'Heat rejection (kW thermal)', 'f', 'B53', '#,##0');
drow(89, 'CRAH / fan-wall unit capacity (kW)', 'in', 150, '#,##0', '');
drow(90, 'CRAH units installed', 'f', RED('B28', 'ROUNDUP(B88/B89,0)'), '0');
drow(91, 'Chiller / heat-rejection unit capacity (kWth)', 'in', 1500, '#,##0', '');
drow(92, 'Chillers installed', 'f', RED('B28', 'ROUNDUP(B88/B91,0)'), '0');
drow(93, 'Cooling capacity (refrigeration tons)', 'f', 'B88/3.517', '#,##0');

dsec(95, 'H. SPACE ALLOWANCES (indicative)');
drow(96, 'UPS and battery rooms (m²)', 'f', `B68/1000*IF(B44="Lithium-ion",50,80)*B69`, '#,##0', '≈ 50 m²/MW Li-ion, 80 m²/MW VRLA, per UPS system');
drow(97, 'LV switchrooms (m²)', 'f', 'B81*B78/1000*30', '#,##0', '≈ 30 m² per MVA of transformer');
drow(98, 'MV substation / transformer bays (m²)', 'f', 'B81*25', '#,##0', '≈ 25 m² per transformer bay');
drow(99, 'Generator yard / hall (m²)', 'f', 'B75*60', '#,##0', '≈ 60 m² per containerised 2–2.5 MVA set incl. access');
drow(100, 'Total electrical plant space (m²)', 'f', 'SUM(B96:B99)', '#,##0');

dsec(102, 'I. ENERGY');
drow(103, 'Annual IT energy (MWh/yr)', 'f', 'B20*B16*8760/1000', '#,##0');
drow(104, 'Annual facility energy (MWh/yr)', 'f', 'B103*B61', '#,##0');
drow(105, 'Annual carbon (tCO₂e/yr)', 'f', `B104*'Energy Management'!B11`, '#,##0', 'Grid factor from Energy Management');
drow(106, 'Annual energy cost', 'f', `'Energy Management'!B8&" "&TEXT(B104*1000*'Energy Management'!B9,"#,##0")`, null, 'Tariff from Energy Management');

dsec(108, 'J. SPECIFICATION (select / edit)');
const SPEC = [
  ['Utility supply', 'Two independent MV feeds from separate primary substations with automatic changeover', 'EN 50600-2-2 / TIA-942'],
  ['MV switchgear', 'SF6-free / vacuum MV switchgear, ring or dual-radial, protection relays with remote monitoring', 'IEC 62271'],
  ['Transformers', 'Cast-resin or ester-filled, Dyn11, Ecodesign Tier 2, K-rated for non-linear load', 'IEC 60076 / IEC 60076-11'],
  ['LV switchgear', 'Form 4b, ACB incomers with LSIG, generator and utility ATS, Modbus / IEC 61850 monitoring', 'IEC 61439-2'],
  ['Generators', 'Diesel / HVO, ISO 8528 DCP rated, synchronised paralleling switchboard, 12–72 h fuel with day tanks', 'ISO 8528-1; NFPA 110'],
  ['UPS', 'Modular double-conversion (VFI-SS-111), hot-swappable modules, ≥ 96% efficiency, Li-ion with BMS', 'IEC 62040-3; IEC 62619'],
  ['Critical distribution', 'Busway (overhead) or floor PDUs / RPPs on A and B paths, branch-circuit monitoring', 'IEC 61439-6; NEC 645'],
  ['Rack power', 'Dual intelligent rack PDUs, per-outlet metering, IEC 60309 inlets', 'IEC 60309; IEC 60320'],
  ['Earthing and bonding', 'Mesh-bonded network, supplementary bonding grid under raised floor, SPD at all levels', 'EN 50310; TIA-607-C; IEC 62305'],
  ['Monitoring', 'DCIM + EPMS: PUE, power quality, battery health, capacity planning', 'ISO/IEC 30134; EN 50600-4'],
  ['Fire protection', 'VESDA early warning; gaseous / water-mist suppression in white space and electrical rooms', 'NFPA 75 / 76; BS EN 15004'],
  ['EPO', 'Emergency power-off where required by local code (with abort / protective covers)', 'NEC 645.10; local code'],
];
SPEC.forEach(([l, v, ref], i) => { const r = 109 + i; drow(r, l, 'in', v, null, ref); dc.getRow(r).height = 28; });
dc.getColumn('A').width = 50; dc.getColumn('B').width = 40; dc.getColumn('C').width = 46; dc.getColumn('D').width = 10; dc.getColumn('E').width = 3;
dc.getColumn('F').width = 40; ['G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'].forEach(c => dc.getColumn(c).width = 12);
DC_MD = `${DCQ}!B62`;
}

