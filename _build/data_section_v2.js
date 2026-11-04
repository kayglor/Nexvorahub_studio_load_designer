// ---------------- 2a. Regional Densities (region methods, multipliers, code unit loads) ----------------
const rd = wb.addWorksheet('Regional Densities');
title(rd, 'Regional Load-Density Basis – Methods, Adjustment Factors and Code Unit Loads', 'The Load Schedule converts the base (UK rule-of-thumb) density into the selected region using the factors below, and checks against code lighting unit loads where the regional code prescribes them (NEC / CEC). All blue values are editable.');
legend(rd, 3);
const RD_REG = [
  ['United Kingdom', 'BS 7671:2018 (18th Ed.); CIBSE Guide F/K; BSRIA BG 9; DNO ADMD policies', 'VA/m² after diversity (rules of thumb); ADMD per dwelling', '–', 'm²'],
  ['Europe (IEC)', 'IEC 60364 / HD 60364; national rules (DIN 18015, NF C 14-100, CEI 64-8, UNE)', 'VA/m² after diversity; national simultaneity tables (e.g. NF C 14-100)', '–', 'm²'],
  ['USA (NEC)', 'NFPA 70 National Electrical Code – Article 220; ASHRAE 90.1 LPD', 'Calculated load: Table 220.12 lighting unit loads + receptacles + HVAC + motors, with NEC demand factors', 'NEC', 'ft²'],
  ['Canada (CEC)', 'CSA C22.1 Canadian Electrical Code – Section 8 (Rules 8-200 to 8-210); NECB', 'Basic load W/m² by occupancy + heating / AC / motors, with CEC demand factors', 'CEC', 'm² (ft² shown)'],
  ['Australia / NZ', 'AS/NZS 3000:2018 Clause 2.2.2 and Appendix C; AS/NZS 3008.1', 'Maximum demand by calculation, assessment or measurement; Appendix C load groups', '–', 'm²'],
  ['Middle East (GCC)', 'DEWA / ADDC Regulations; SBC 401; KAHRAMAA Regulations', 'Connected load × authority demand factors; cooling-dominant', '–', 'm²'],
  ['Africa', 'SANS 10142-1; NRS 034; BS 7671 practice (Nigeria, Ghana, Kenya)', 'VA/m² after diversity; NRS 034 ADMD for dwellings', '–', 'm²'],
  ['South & SE Asia', 'NBC India Part 8 / IS 732; SS 638; MS IEC 60364; EMSD CoP', 'Connected load × demand factors; cooling-dominant', '–', 'm²'],
  ['Latin America & Caribbean', 'NBR 5410; AEA 90364; NOM-001-SEDE; RETIE / NTC 2050; national NEC adoptions', 'IEC- or NEC-based national methods with demand factors', '–', 'm²'],
  ['East Asia', 'GB 50054 / GB 50052 (China); JIS / JEAC 8001 (Japan); KEC (Korea)', 'Demand-factor (Kx) method per national design codes', '–', 'm²'],
];
header(rd, 5, ['Density region', 'Codes / regulations', 'Load-estimate method', 'Code lighting table', 'Area unit']);
const RDR0 = 6, RDR1 = RDR0 + RD_REG.length - 1;
RD_REG.forEach((row, i) => row.forEach((v, j) => { const c = rd.getCell(RDR0 + i, j + 1); c.value = v; c.border = box; c.font = { name: 'Arial', size: j === 0 ? 10 : 9, bold: j === 0 }; c.alignment = { wrapText: true, vertical: 'top' }; }));
for (let r = RDR0; r <= RDR1; r++) rd.getRow(r).height = 30;
const RDQ = "'Regional Densities'";
const DREG_LIST = `${RDQ}!$A$${RDR0}:$A$${RDR1}`;
const DREG_CODE = `${RDQ}!$D$${RDR0}:$D$${RDR1}`;
const DREG_METH = `${RDQ}!$C$${RDR0}:$C$${RDR1}`;

const GROUPS = ['Office', 'Retail', 'Hospitality', 'Education', 'Healthcare', 'Residential', 'Industrial – light / logistics', 'Industrial – process / heavy', 'Data / IT / labs', 'Leisure & assembly', 'Transport & public services', 'Worship / community', 'Car parks & external'];
// multipliers: [UK, Europe, USA, Canada, AUS, ME, Africa, Asia, LatAm, East Asia]
const MULT = {
  'Office': [1, 1, 1.15, 1.1, 1.1, 1.4, 1.1, 1.3, 1.2, 1.2],
  'Retail': [1, 1, 1.1, 1.1, 1.1, 1.4, 1.1, 1.3, 1.2, 1.2],
  'Hospitality': [1, 1, 1.15, 1.1, 1.1, 1.4, 1.2, 1.3, 1.2, 1.2],
  'Education': [1, 1, 1.1, 1.2, 1.0, 1.5, 0.9, 1.2, 1, 1.1],
  'Healthcare': [1, 1, 1.1, 1.1, 1.05, 1.3, 1.1, 1.2, 1.1, 1.15],
  'Residential': [1, 1, 1.4, 1.6, 1.2, 1.8, 0.8, 1.2, 1, 1.1],
  'Industrial – light / logistics': [1, 1, 1.0, 1.05, 1.0, 1.3, 1.0, 1.1, 1, 1],
  'Industrial – process / heavy': [1, 1, 1.0, 1.0, 1.0, 1.1, 1.0, 1.0, 1, 1],
  'Data / IT / labs': [1, 1, 1.0, 1.0, 1.0, 1.1, 1.05, 1.05, 1, 1],
  'Leisure & assembly': [1, 1, 1.1, 1.05, 1.0, 1.4, 1.0, 1.2, 1.1, 1.1],
  'Transport & public services': [1, 1, 1.0, 1.05, 1.0, 1.3, 1.0, 1.2, 1.1, 1.1],
  'Worship / community': [1, 1, 1.1, 1.05, 1.0, 1.6, 1.0, 1.3, 1.1, 1.1],
  'Car parks & external': [1, 1, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1, 1],
};
const MG = RDR1 + 3;
sub(rd, `A${MG}`, 'B. Regional adjustment factor applied to the base (UK) density – by load group');
header(rd, MG + 1, ['Load group', ...RD_REG.map(r => r[0]), 'Basis']);
rd.getRow(MG + 1).height = 32;
const MG0 = MG + 2, MG1 = MG0 + GROUPS.length - 1;
const MBASIS = {
  'Office': 'Higher plug loads and cooling in North America; cooling-dominant in GCC / Asia',
  'Residential': 'Larger dwellings, AC and electric cooking in USA; electric heating common in Canada; AC in GCC',
  'Education': 'Electric heating (Canada) and full AC (GCC)',
  'Worship / community': 'Air-conditioned prayer halls in GCC / Asia',
};
GROUPS.forEach((g, i) => {
  const r = MG0 + i;
  const a = rd.getCell(`A${r}`); a.value = g; a.border = box; a.font = FONT;
  MULT[g].forEach((v, k) => { const c = rd.getCell(r, 2 + k); input(c, v, '0.00'); c.alignment = { horizontal: 'center' }; });
  const n = rd.getCell(r, 12); n.value = MBASIS[g] || ''; n.border = box; n.font = { name: 'Arial', size: 9, italic: true };
});
rd.getCell(`A${MG1 + 1}`).value = 'Indicative engineering adjustments reflecting climate, building practice and typical equipment in each region (UK = 1.00). Replace with local authority / utility data where available.';
rd.getCell(`A${MG1 + 1}`).font = { name: 'Arial', size: 9, italic: true };
const GROUP_LIST = `${RDQ}!$A$${MG0}:$A$${MG1}`;
const MULT_TAB = `${RDQ}!$B$${MG0}:$K$${MG1}`;
const MULT_HDR = `${RDQ}!$B$${MG + 1}:$K$${MG + 1}`;

const OC = MG1 + 4;
sub(rd, `A${OC}`, 'C. Code lighting / basic unit loads by occupancy – NEC Table 220.12 (USA) and CEC Section 8 (Canada)');
header(rd, OC + 1, ['Occupancy', 'NEC unit load (VA/ft²)', 'NEC unit load (VA/m²)', 'NEC lighting demand factors (Table 220.42)', 'CEC basic load (W/m²) – VERIFY', 'Notes']);
rd.getRow(OC + 1).height = 32;
const DF_STD = 'All others: 100%';
const OCC = [
  ['Automotive facility', 1.5, DF_STD, 10], ['Convention center', 1.4, DF_STD, 10], ['Courthouse', 1.4, DF_STD, 20], ['Dormitory', 1.5, DF_STD, 20],
  ['Exercise center', 1.4, DF_STD, 20], ['Fire station', 1.3, DF_STD, 20], ['Gymnasium', 1.7, DF_STD, 20], ['Health care clinic', 1.6, DF_STD, 50],
  ['Hospital', 1.6, 'First 50,000 VA at 40%; remainder 20% (not for areas used all at once, e.g. theatres)', 50], ['Hotel / motel', 1.7, 'First 20,000 VA at 60%; 20,001–100,000 at 50%; remainder 35%', 20],
  ['Library', 1.5, DF_STD, 30], ['Manufacturing facility', 2.2, DF_STD, 25], ['Motion picture theater', 1.6, DF_STD, 10], ['Museum', 1.6, DF_STD, 30],
  ['Office', 1.3, DF_STD + '; plus receptacles 1 VA/ft² where count unknown (220.14(K))', 50], ['Parking garage', 0.1, DF_STD, 5], ['Penitentiary', 1.2, DF_STD, 20],
  ['Performing arts theater', 1.5, DF_STD, 10], ['Police station', 1.3, DF_STD, 30], ['Post office', 1.6, DF_STD, 30], ['Religious facility', 2.2, DF_STD, 10],
  ['Restaurant', 1.5, DF_STD, 30], ['Retail', 1.9, DF_STD, 30], ['School / university', 1.5, DF_STD, 50], ['Sports arena', 1.5, DF_STD, 10],
  ['Town hall', 1.4, DF_STD, 30], ['Transportation', 1.2, DF_STD, 10], ['Warehouse', 1.2, 'First 12,500 VA at 100%; remainder 50%', 10], ['Workshop', 1.7, DF_STD, 25],
  ['Dwelling unit', 3.0, 'First 3,000 VA at 100%; 3,001–120,000 at 35%; remainder 25%', 0], ['(none – no code unit load)', 0, '–', 0],
];
const OC0 = OC + 2, OC1 = OC0 + OCC.length - 1;
OCC.forEach(([n, va, dfx, cec], i) => {
  const r = OC0 + i;
  const a = rd.getCell(`A${r}`); a.value = n; a.border = box; a.font = FONT;
  input(rd.getCell(`B${r}`), va, '0.0'); rd.getCell(`B${r}`).alignment = { horizontal: 'center' };
  calc(rd.getCell(`C${r}`), `B${r}*10.764`, '0.0'); rd.getCell(`C${r}`).alignment = { horizontal: 'center' };
  const d = rd.getCell(`D${r}`); d.value = dfx; d.border = box; d.font = { name: 'Arial', size: 9 }; d.alignment = { wrapText: true };
  input(rd.getCell(`E${r}`), cec, '0'); rd.getCell(`E${r}`).alignment = { horizontal: 'center' };
  const n2 = rd.getCell(`F${r}`); n2.value = n === 'Dwelling unit' ? 'CEC: use Rule 8-200 per dwelling (Calculated Load per unit)' : ''; n2.border = box; n2.font = { name: 'Arial', size: 9, italic: true };
});
rd.getCell(`A${OC1 + 1}`).value = 'NEC values per NFPA 70 (2020/2023) Table 220.12 and Table 220.42 – verify against the edition adopted by the Authority Having Jurisdiction. CEC values are indicative placeholders for CSA C22.1 Section 8 (Rules 8-204 to 8-210, Table 14) and MUST be verified against the code edition in force before use.';
rd.getCell(`A${OC1 + 1}`).font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FFC00000' } }; rd.getCell(`A${OC1 + 1}`).alignment = { wrapText: true };
rd.mergeCells(OC1 + 1, 1, OC1 + 2, 10);
const OCC_LIST = `${RDQ}!$A$${OC0}:$A$${OC1}`, OCC_NEC = `${RDQ}!$C$${OC0}:$C$${OC1}`, OCC_NECF = `${RDQ}!$B$${OC0}:$B$${OC1}`, OCC_CEC = `${RDQ}!$E$${OC0}:$E$${OC1}`;
rd.getColumn(1).width = 34; rd.getColumn(2).width = 44; rd.getColumn(3).width = 46; rd.getColumn(4).width = 40; rd.getColumn(5).width = 16;
for (let c = 6; c <= 11; c++) rd.getColumn(c).width = 14; rd.getColumn(12).width = 50;

// ---------------- 2. VA-m2 Data (base database, UK rule-of-thumb basis) ----------------
const dt = wb.addWorksheet('VA-m2 Data');
title(dt, 'Load Density Database – Base VA/m² GIA (after diversity), with Load Group and Code Occupancy', 'Base values are UK rules of thumb (BSRIA BG 9, CIBSE Guide F/K, BCO, consultant practice). The Load Schedule applies the regional factor for the selected region. Filter by use class, category or load group.');
header(dt, 4, ['Use class (UK)', 'Category', 'Building / space type', 'Min', 'Typical', 'Max', 'Load group', 'Code occupancy (NEC / CEC)', 'Notes']);
const USE_CLASSES = [
  'E(a) Retail', 'E(b) Food & drink', 'E(c) Financial / professional', 'E(d) Indoor sport & fitness', 'E(e) Medical / health', 'E(f) Crèche / nursery',
  'E(g)(i) Offices', 'E(g)(ii) Research & development', 'E(g)(iii) Light industrial', 'B2 General industrial', 'B8 Storage & distribution',
  'C1 Hotels', 'C2 Residential institutions', 'C2A Secure residential', 'C3 Dwellinghouses', 'C4 Houses in multiple occupation',
  'F1 Learning & non-residential institutions', 'F2 Local community', 'Sui generis',
];
const U = (s) => USE_CLASSES.find(u => u.startsWith(s + ' '));
const data = [
  [U('E(g)(i)'), 'Offices', 'Office – naturally ventilated, gas heated', 40, 50, 60, 'LED lighting, basic small power'],
  [U('E(g)(i)'), 'Offices', 'Office – mixed mode', 60, 72, 85, ''],
  [U('E(g)(i)'), 'Offices', 'Office – fully air-conditioned', 80, 100, 120, 'Typical Cat A speculative office (VRF / 4-pipe)'],
  [U('E(g)(i)'), 'Offices', 'Office – all-electric (heat pumps)', 100, 120, 140, 'Net-zero / no gas'],
  [U('E(g)(i)'), 'Offices', 'Office – prestige / trading floor', 150, 200, 250, 'Heavy IT and cooling'],
  [U('E(c)'), 'Offices', 'Bank / professional services branch', 60, 80, 100, ''],
  [U('E(g)(i)'), 'Large commercial', 'Office tower – high-rise (> 20 floors), air-conditioned', 110, 140, 170, 'Includes lifts, pressurisation, central plant'],
  [U('E(g)(i)'), 'Large commercial', 'Office campus / corporate headquarters', 100, 130, 160, 'Incl. restaurant, data room, auditorium'],
  [U('E(g)(i)'), 'Large commercial', 'Call centre / high-density office', 120, 150, 180, 'Occupancy ~1 person per 6 m²'],
  [U('Sui'), 'Large commercial', 'Mixed-use tower (office / retail / hotel podium)', 90, 120, 150, 'Use a zone per use where known'],
  [U('E(a)'), 'Retail', 'Commercial shop – small lock-up unit', 50, 70, 90, 'High-street / neighbourhood parade'],
  [U('E(a)'), 'Retail', 'Retail – non-food / shell unit', 60, 80, 100, 'Tenant fit-out dependent'],
  [U('E(a)'), 'Retail', 'Retail – department store', 80, 105, 130, ''],
  [U('E(a)'), 'Retail', 'Retail – supermarket / food', 150, 200, 250, 'Refrigeration dominant'],
  [U('E(a)'), 'Retail', 'Retail – convenience store', 120, 160, 200, 'Refrigeration dominant'],
  [U('E(a)'), 'Retail', 'Retail warehouse / garden centre', 30, 45, 60, ''],
  [U('E(a)'), 'Retail', 'Shopping centre – mall common areas', 40, 60, 80, ''],
  [U('E(a)'), 'Large commercial', 'Super-regional shopping mall (incl. tenants, AC)', 100, 140, 180, 'Whole-building GLA + malls'],
  [U('E(a)'), 'Large commercial', 'Hypermarket / big-box superstore', 150, 200, 250, 'Refrigeration, bakery, food hall'],
  [U('E(a)'), 'Large commercial', 'Retail park – big-box non-food', 40, 60, 80, 'Shell + typical tenant fit-out'],
  [U('Sui'), 'Retail', 'Car showroom', 50, 70, 90, 'Workshop loads extra'],
  [U('Sui'), 'Retail', 'Petrol station / forecourt shop', 80, 120, 160, 'Excl. canopy rapid EV chargers'],
  [U('C1'), 'Hospitality', 'Hotel – budget', 50, 65, 80, 'Alt: 3–5 kVA per bedroom'],
  [U('C1'), 'Hospitality', 'Hotel – 4/5 star, air-conditioned', 90, 115, 140, ''],
  [U('C1'), 'Hospitality', 'Serviced apartments / aparthotel', 40, 55, 70, ''],
  [U('C1'), 'Large commercial', 'Resort hotel / integrated resort (casino, F&B, spa)', 120, 160, 200, 'Excl. theme-park rides'],
  [U('E(b)'), 'Hospitality', 'Restaurant – dining area', 80, 115, 150, ''],
  [U('E(b)'), 'Hospitality', 'Commercial kitchen – all electric', 500, 750, 1000, 'Size from equipment schedule'],
  [U('E(b)'), 'Hospitality', 'Restaurant – whole (dining + electric kitchen)', 200, 285, 400, 'Kitchen typically 30–40% of floor area'],
  [U('E(b)'), 'Hospitality', 'Café / coffee shop', 120, 170, 220, ''],
  [U('Sui'), 'Hospitality', 'Pub / bar', 80, 120, 160, 'Cellar cooling and kitchen'],
  [U('Sui'), 'Hospitality', 'Hot food takeaway', 300, 450, 600, 'Fryers / ovens dominate'],
  [U('Sui'), 'Hospitality', 'Nightclub / music venue', 120, 170, 220, 'Sound and lighting rigs'],
  [U('Sui'), 'Hospitality', 'Casino / amusement centre', 100, 150, 200, ''],
  [U('F1'), 'Education', 'School – primary', 25, 35, 45, ''],
  [U('F1'), 'Education', 'School – secondary', 40, 52, 65, ''],
  [U('F1'), 'Education', 'School – secondary with labs/DT/kitchens', 55, 70, 85, ''],
  [U('F1'), 'Education', 'FE college / university teaching', 60, 80, 100, ''],
  [U('F1'), 'Large commercial', 'University campus (teaching, labs, residences)', 60, 90, 120, 'Campus average'],
  [U('E(f)'), 'Education', 'Nursery / crèche', 25, 35, 45, ''],
  [U('E(g)(ii)'), 'Education', 'University research laboratories', 120, 185, 250, ''],
  [U('E(e)'), 'Healthcare', 'GP surgery / health centre', 40, 55, 70, ''],
  [U('E(e)'), 'Healthcare', 'Dental surgery / veterinary clinic', 50, 70, 90, 'X-ray and sterilisation'],
  [U('C2'), 'Healthcare', 'Care home', 40, 55, 70, 'Alt: 2–3 kVA per bed'],
  [U('C2'), 'Healthcare', 'Mental health inpatient unit', 60, 80, 100, ''],
  [U('C2'), 'Healthcare', 'Hospital – acute', 100, 140, 180, 'HTM 06-01; add standby generation'],
  [U('C2'), 'Healthcare', 'Hospital – theatres / imaging / ICU', 150, 200, 250, ''],
  [U('C2'), 'Large commercial', 'Hospital campus (acute + support buildings)', 120, 160, 200, 'Energy centre / CHP excluded'],
  [U('C3'), 'Residential', 'Residential – gas heated', 15, 22, 30, 'Prefer ADMD sheet for DNO applications'],
  [U('C3'), 'Residential', 'Residential – all-electric (ASHP)', 35, 48, 60, 'Prefer ADMD sheet for DNO applications'],
  [U('C3'), 'Residential', 'Residential – furnished flat (all-electric, appliances)', 40, 55, 70, 'Furnished / serviced lettings; prefer ADMD for DNO'],
  [U('C3'), 'Residential', 'Residential landlord / common areas', 10, 15, 20, 'Common area services: lifts, lighting, ventilation, door entry'],
  [U('C3'), 'Residential', 'Common areas – fire & life-safety services', 3, 5, 8, 'Smoke ventilation, fire alarm, emergency lighting, sprinkler controls'],
  [U('C4'), 'Residential', 'House in multiple occupation (HMO)', 25, 35, 45, 'Or ADMD per letting room'],
  [U('Sui'), 'Residential', 'Student accommodation (PBSA)', 30, 40, 50, 'Alt: 1.5–2.5 kVA per bedroom'],
  [U('C2A'), 'Residential', 'Prison / secure residential', 50, 70, 90, 'Security systems and kitchens'],
  [U('B8'), 'Industrial – logistics', 'Warehouse – unheated', 10, 18, 25, ''],
  [U('B8'), 'Industrial – logistics', 'Warehouse – heated, high-bay LED', 25, 38, 50, ''],
  [U('B8'), 'Industrial – logistics', 'Distribution centre – automated', 50, 75, 100, ''],
  [U('B8'), 'Industrial – logistics', 'Logistics – mega distribution centre (> 50,000 m²)', 40, 65, 90, 'Conveyors, sortation, EV fleet charging extra'],
  [U('B8'), 'Industrial – logistics', 'Logistics – parcel sortation hub', 80, 120, 160, 'High-speed sorters'],
  [U('B8'), 'Industrial – logistics', 'Chilled / cold store', 80, 140, 200, 'Temperature regime dependent'],
  [U('B8'), 'Industrial – logistics', 'Logistics – frozen warehouse (−25 °C)', 150, 220, 300, 'Refrigeration plant + underfloor heating'],
  [U('B8'), 'Industrial – logistics', 'Self-storage', 10, 15, 20, ''],
  [U('E(g)(iii)'), 'Industrial – light', 'Light industrial / workshop', 40, 60, 80, ''],
  [U('B2'), 'Industrial – light', 'Industrial – vehicle maintenance depot / bus garage', 60, 100, 150, 'Excl. depot EV / bus charging'],
  [U('Sui'), 'Industrial – light', 'Industrial – aircraft hangar (maintenance)', 40, 70, 100, 'Ground power units extra'],
  [U('B2'), 'Industrial – process', 'Heavy industrial / manufacturing (general)', 100, 300, 500, 'Size from process load'],
  [U('B2'), 'Industrial – process', 'Industrial – food & beverage processing', 150, 250, 400, 'Refrigeration, process heat, CIP'],
  [U('B2'), 'Industrial – process', 'Industrial – brewery / bottling plant', 150, 225, 300, ''],
  [U('B2'), 'Industrial – process', 'Industrial – automotive assembly plant', 150, 250, 350, 'Paint shop and robots dominate'],
  [U('B2'), 'Industrial – process', 'Industrial – metal fabrication / machine shop', 80, 150, 250, 'Welding sets: check P28 flicker'],
  [U('B2'), 'Industrial – process', 'Industrial – foundry / electric furnaces', 300, 600, 1000, 'Furnace rating from process data'],
  [U('B2'), 'Industrial – process', 'Industrial – plastics / injection moulding', 200, 350, 500, ''],
  [U('B2'), 'Industrial – process', 'Industrial – printing / packaging', 100, 175, 250, ''],
  [U('B2'), 'Industrial – process', 'Industrial – textiles / garment manufacturing', 80, 130, 180, ''],
  [U('B2'), 'Industrial – process', 'Industrial – electronics assembly (SMT)', 150, 250, 350, 'Clean-air and reflow ovens'],
  [U('B2'), 'Industrial – process', 'Industrial – semiconductor fab / cleanroom (ISO 5)', 1000, 1500, 2500, 'Per m² cleanroom'],
  [U('B2'), 'Industrial – process', 'Clean room / pharmaceutical manufacturing', 300, 500, 800, 'HVAC-dominated'],
  [U('B2'), 'Industrial – process', 'Industrial – chemical / process plant', 200, 350, 600, 'Hazardous-area equipment'],
  [U('B2'), 'Industrial – process', 'Industrial – paper / pulp mill', 300, 500, 800, ''],
  [U('B2'), 'Industrial – process', 'Industrial – battery cell / gigafactory production', 400, 700, 1000, 'Dry rooms and formation'],
  [U('Sui'), 'Industrial – process', 'Industrial – water / wastewater treatment works', 100, 200, 350, 'Per m² of process buildings; pumps / blowers dominate'],
  [U('B2'), 'Industrial – process', 'Industrial – compressed air / process utilities plant room', 300, 500, 800, ''],
  [U('B8'), 'Specialist', 'Data centre – white space', 1000, 2000, 3000, 'IT load × PUE'],
  [U('B8'), 'Large commercial', 'Hyperscale data centre – white space', 1500, 2500, 4000, 'IT load × PUE; utility HV supply'],
  [U('E(g)(ii)'), 'Specialist', 'R&D / life-science building', 100, 150, 200, ''],
  [U('E(g)(ii)'), 'Specialist', 'Laboratory – general', 100, 150, 200, 'Fume cupboards and plant'],
  [U('Sui'), 'Specialist', 'Comms / server room', 500, 1000, 1500, ''],
  [U('Sui'), 'Specialist', 'Telecom exchange', 300, 500, 800, ''],
  [U('Sui'), 'Specialist', 'TV / broadcast / film studio', 150, 250, 350, 'Studio lighting rigs'],
  [U('Sui'), 'Specialist', 'Commercial laundry / laundrette', 150, 250, 350, ''],
  [U('E(d)'), 'Leisure / public', 'Sports hall (dry)', 30, 40, 50, ''],
  [U('E(d)'), 'Leisure / public', 'Gym / fitness centre', 60, 85, 110, 'Cooling and ventilation'],
  [U('F2'), 'Leisure / public', 'Leisure centre with pool', 80, 115, 150, 'Pool plant and ventilation'],
  [U('F2'), 'Leisure / public', 'Ice rink', 100, 150, 200, 'Refrigeration plant'],
  [U('F2'), 'Places of worship', 'Community hall', 20, 30, 40, ''],
  [U('F1'), 'Leisure / public', 'Library / museum / gallery', 50, 70, 90, 'Higher with environmental control'],
  [U('F1'), 'Places of worship', 'Place of worship – church / chapel', 20, 30, 40, 'Gas or heat-pump heated'],
  [U('F1'), 'Places of worship', 'Place of worship – mosque (naturally ventilated)', 20, 30, 40, 'Ablution water heating included'],
  [U('F1'), 'Places of worship', 'Place of worship – mosque (air-conditioned)', 40, 60, 80, 'Hot climates: prayer-hall cooling dominant'],
  [U('F1'), 'Public services', 'Civic centre / public service facility', 40, 60, 80, 'Council offices, job centre, community services'],
  [U('Sui'), 'Public services', 'Public toilets / welfare block', 15, 25, 35, 'Hand dryers and water heating'],
  [U('F1'), 'Public services', 'Law court', 60, 80, 100, ''],
  [U('F1'), 'Leisure / public', 'Conference / exhibition centre', 60, 90, 120, 'Exhibition power extra'],
  [U('Sui'), 'Leisure / public', 'Theatre / cinema', 80, 115, 150, ''],
  [U('Sui'), 'Leisure / public', 'Stadium – stands / concourse building', 60, 85, 110, 'Excl. floodlights, pitch heating, big screens – add on Specific Loads'],
  [U('Sui'), 'Transport / public', 'Airport terminal', 100, 140, 180, 'Baggage handling extra'],
  [U('Sui'), 'Transport / public', 'Railway / bus station', 40, 60, 80, ''],
  [U('Sui'), 'Transport / public', 'Cruise / ferry terminal', 60, 90, 120, 'Shore power extra'],
  [U('Sui'), 'Transport / public', 'Police / fire / ambulance station', 50, 70, 90, 'Standby generation'],
  [U('Sui'), 'Car parks', 'Car park – open / surface', 2, 3, 5, 'Lighting only; EV on Specific Loads sheet'],
  [U('Sui'), 'Car parks', 'Car park – enclosed, ventilated', 10, 15, 20, 'Indoor parking; EV on Specific Loads sheet'],
  [U('Sui'), 'External / public realm', 'Parks and gardens (per m² site area)', 1, 2, 4, 'Lighting, irrigation, features – detail on External Lighting'],
  [U('Sui'), 'External / public realm', 'Open space / public plaza (per m² site area)', 2, 3, 5, 'Lighting, events power – detail on External Lighting'],
  [U('Sui'), 'External / public realm', 'Street lighting (per m² of carriageway)', 0.5, 1, 1.5, 'Prefer column-based calc on External Lighting'],
];
const groupOf = (cat, name) => {
  if (/research laborator|R&D|Laboratory|Comms|Telecom|Data centre|data centre|broadcast/i.test(name)) return 'Data / IT / labs';
  if (/laundr/i.test(name)) return 'Industrial – process / heavy';
  if (/Law court|Civic|toilets/i.test(name)) return 'Transport & public services';
  if (/Community hall/i.test(name)) return 'Worship / community';
  if (/Office|Call centre|Mixed-use/i.test(name) || cat === 'Offices') return 'Office';
  if (/mall|Hypermarket|Retail park/i.test(name) || cat === 'Retail') return 'Retail';
  if (/Resort/i.test(name) || cat === 'Hospitality') return 'Hospitality';
  if (/University campus/i.test(name) || cat === 'Education') return 'Education';
  if (/Hospital campus/i.test(name) || cat === 'Healthcare') return 'Healthcare';
  if (cat === 'Residential') return 'Residential';
  if (cat === 'Industrial – logistics' || cat === 'Industrial – light') return 'Industrial – light / logistics';
  if (cat === 'Industrial – process') return 'Industrial – process / heavy';
  if (cat === 'Leisure / public') return 'Leisure & assembly';
  if (cat === 'Places of worship') return 'Worship / community';
  if (cat === 'Public services' || cat === 'Transport / public') return 'Transport & public services';
  return 'Car parks & external';
};
const OCC_MAP = [
  [/Car showroom|Petrol|vehicle maintenance/i, 'Automotive facility'], [/Office|Bank|Call centre|Mixed-use/i, 'Office'],
  [/Hotel|Serviced|Resort/i, 'Hotel / motel'], [/Restaurant|Café|Pub|takeaway|kitchen/i, 'Restaurant'],
  [/Nightclub|Casino|broadcast/i, 'Performing arts theater'], [/School|college|University|Nursery/i, 'School / university'],
  [/GP|Dental/i, 'Health care clinic'], [/Hospital|Mental health/i, 'Hospital'], [/Care home|Student|HMO/i, 'Dormitory'],
  [/Prison/i, 'Penitentiary'], [/landlord|Common areas/i, '(none – no code unit load)'], [/Residential/i, 'Dwelling unit'],
  [/Data centre|Comms|Telecom|toilets|Parks|Open space|Street|open \/ surface/i, '(none – no code unit load)'],
  [/Retail|shop|mall|Hypermarket|superstore/i, 'Retail'],
  [/Warehouse|storage|Distribution|Logistics|Chilled|Self-storage/i, 'Warehouse'], [/Light industrial|hangar|laundr/i, 'Workshop'],
  [/Industrial|Clean room|manufacturing/i, 'Manufacturing facility'], [/R&D|Laboratory/i, 'School / university'],
  [/Retail|shop|mall|Hypermarket|store|superstore/i, 'Retail'], [/Sports hall/i, 'Gymnasium'], [/Gym/i, 'Exercise center'],
  [/Leisure centre|Ice rink|Stadium/i, 'Sports arena'], [/Community hall|Conference/i, 'Convention center'], [/Library/i, 'Library'],
  [/worship/i, 'Religious facility'], [/Civic/i, 'Town hall'], [/Law court/i, 'Courthouse'], [/Theatre/i, 'Motion picture theater'],
  [/Airport|Railway|Cruise/i, 'Transportation'], [/Police/i, 'Police station'], [/enclosed/i, 'Parking garage'],
];
const occOf = (name) => { for (const [re, o] of OCC_MAP) if (re.test(name)) return o; return '(none – no code unit load)'; };
const D0 = 5, D1 = D0 + data.length - 1;
data.forEach((row, i) => {
  const r = D0 + i;
  const vals = [row[0], row[1], row[2], row[3], row[4], row[5], groupOf(row[1], row[2]), occOf(row[2]), row[6]];
  vals.forEach((v, j) => {
    const c = dt.getCell(r, j + 1);
    if (j >= 3 && j <= 5) { input(c, v, '#,##0.0'); c.alignment = { horizontal: 'center' }; }
    else if (j === 6 || j === 7) { input(c, v); c.font = { name: 'Arial', size: 9, color: { argb: BLUE } }; }
    else { c.value = v; c.border = box; c.font = { name: 'Arial', size: j === 8 ? 9 : 10 }; }
  });
  dt.getCell(`G${r}`).dataValidation = { type: 'list', allowBlank: false, formulae: [GROUP_LIST] };
  dt.getCell(`H${r}`).dataValidation = { type: 'list', allowBlank: false, formulae: [OCC_LIST] };
});
const srcRow = D1 + 2;
dt.getCell(`A${srcRow}`).value = 'Source: indicative values from BSRIA BG 9/2011 Rules of Thumb, CIBSE Guide F & K, BCO Guide to Specification and UK/international consultant practice. Values are VA/m² GIA after diversity (base / UK). Load group sets the regional factor; code occupancy sets the NEC / CEC unit-load check. All blue values are editable.';
dt.getCell(`A${srcRow}`).font = { name: 'Arial', size: 9, italic: true };
dt.getCell('K4').value = 'Standard main switchboard / ACB ratings (A)';
dt.getCell('M4').value = 'Standard transformer sizes (kVA)';
dt.getCell('O4').value = 'UK Use Classes (2020)';
['K4', 'M4', 'O4'].forEach(a => { const c = dt.getCell(a); c.fill = hdrFill; c.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } }; c.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' }; c.border = box; });
const swb = [100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3200, 4000];
const tx = [200, 315, 500, 800, 1000, 1250, 1500, 1600, 2000, 2500, 3150];
swb.forEach((v, i) => { const c = dt.getCell(`K${5 + i}`); input(c, v, '#,##0'); c.alignment = { horizontal: 'center' }; });
tx.forEach((v, i) => { const c = dt.getCell(`M${5 + i}`); input(c, v, '#,##0'); c.alignment = { horizontal: 'center' }; });
USE_CLASSES.forEach((v, i) => { const c = dt.getCell(`O${5 + i}`); c.value = v; c.border = box; c.font = FONT; });
const SWB = `'VA-m2 Data'!$K$5:$K$${4 + swb.length}`;
const TX = `'VA-m2 Data'!$M$5:$M$${4 + tx.length}`;
dt.getColumn(1).width = 26; dt.getColumn(2).width = 18; dt.getColumn(3).width = 50; [4, 5, 6].forEach(i => dt.getColumn(i).width = 9);
dt.getColumn(7).width = 28; dt.getColumn(8).width = 26; dt.getColumn(9).width = 42; dt.getColumn(10).width = 3; dt.getColumn(11).width = 16; dt.getColumn(12).width = 3; dt.getColumn(13).width = 16; dt.getColumn(14).width = 3; dt.getColumn(15).width = 34;
dt.views = [{ state: 'frozen', ySplit: 4 }];
dt.autoFilter = { from: { row: 4, column: 1 }, to: { row: D1, column: 9 } };

const TYPES = `'VA-m2 Data'!$C$${D0}:$C$${D1}`;
const UCOL = `'VA-m2 Data'!$A$${D0}:$A$${D1}`;
const GCOL = `'VA-m2 Data'!$G$${D0}:$G$${D1}`;
const OCOL = `'VA-m2 Data'!$H$${D0}:$H$${D1}`;
const DENS = `'VA-m2 Data'!$D$${D0}:$F$${D1}`;
const PF = `'Project Info'!$B$19`;

// ---------------- 3. Load Schedule (regional) ----------------
const ls = wb.addWorksheet('Load Schedule');
title(ls, 'Area-Based Load Schedule – Regional (UK, Europe, USA, Canada, Australia/NZ, Middle East, Africa, Asia)', 'Select the design region for this schedule, then a building / space type per zone. Base density × regional factor = regional design density; NEC / CEC regions also show the code lighting unit load as a minimum check.');
legend(ls, 3);
const r4 = (cell, v, bold) => { const c = ls.getCell(cell); c.value = v; c.font = { name: 'Arial', size: 10, bold: !!bold }; };
r4('A4', 'Design region:', true);
input(ls.getCell('C4'), 'Use project region');
ls.getCell('C4').dataValidation = { type: 'list', allowBlank: false, formulae: [`"Use project region,${RD_REG.map(r => r[0]).join(',')}"`] };
r4('D4', 'Applied:', true);
calc(ls.getCell('E4'), `IF(C4="Use project region",'Project Info'!$B$35,C4)`, null, GREEN); ls.mergeCells('E4:G4');
ls.getCell('E4').font = { name: 'Arial', size: 10, bold: true, color: { argb: GREEN } };
r4('H4', 'Method:', true);
calc(ls.getCell('I4'), `INDEX(${DREG_METH},MATCH(E4,${DREG_LIST},0))&"  |  Code check: "&INDEX(${DREG_CODE},MATCH(E4,${DREG_LIST},0))`); ls.mergeCells('I4:R4');
ls.getCell('I4').font = { name: 'Arial', size: 9, italic: true };
const LREG = '$E$4';
const LCODE = `INDEX(${DREG_CODE},MATCH(${LREG},${DREG_LIST},0))`;
header(ls, 5, ['Ref', 'Zone / description', 'Building / space type', 'GIA (m²)', 'GIA (ft²)', 'Base VA/m² (UK)', 'Regional factor', 'Regional VA/m²', 'Override VA/m²', 'Design VA/m²', 'Design VA/ft²', 'Demand (kVA)', 'Code lighting unit load (VA/m²)', 'Code lighting load (kVA)', 'Code check', 'Use class (UK)', 'Load group', 'Notes']);
ls.getRow(5).height = 45;
const L0 = 6, L1 = 55;
const examples = [
  ['Z01', 'Ground floor retail units', 'Retail – non-food / shell unit', 850],
  ['Z02', 'Ground floor café', 'Café / coffee shop', 150],
  ['Z03', 'Levels 1–4 offices', 'Office – all-electric (heat pumps)', 4200],
  ['Z04', 'Gym (level 5)', 'Gym / fitness centre', 600],
  ['Z05', 'Basement car park', 'Car park – enclosed, ventilated', 1500],
  ['Z06', 'Residential landlord areas', 'Residential landlord / common areas', 600],
  ['Z07', 'Production hall', 'Industrial – food & beverage processing', 3000],
  ['Z08', 'Finished-goods warehouse', 'Warehouse – heated, high-bay LED', 5000],
];
for (let r = L0; r <= L1; r++) {
  const ex = examples[r - L0];
  input(ls.getCell(`A${r}`), ex ? ex[0] : null);
  input(ls.getCell(`B${r}`), ex ? ex[1] : null);
  input(ls.getCell(`C${r}`), ex ? ex[2] : null);
  ls.getCell(`C${r}`).dataValidation = { type: 'list', allowBlank: true, formulae: [TYPES], showErrorMessage: true, error: 'Choose a type from the VA-m2 Data list' };
  input(ls.getCell(`D${r}`), ex ? ex[3] : null, '#,##0');
  calc(ls.getCell(`E${r}`), `IF(D${r}="","",D${r}*10.764)`, '#,##0');
  const m = `MATCH(C${r},${TYPES},0)`;
  calc(ls.getCell(`F${r}`), `IF(C${r}="","",INDEX(${DENS},${m},MATCH('Project Info'!$B$21,'VA-m2 Data'!$D$4:$F$4,0)))`, '#,##0.0', GREEN);
  calc(ls.getCell(`Q${r}`), `IF(C${r}="","",INDEX(${GCOL},${m}))`, null, GREEN);
  calc(ls.getCell(`G${r}`), `IF(C${r}="","",INDEX(${MULT_TAB},MATCH(Q${r},${GROUP_LIST},0),MATCH(${LREG},${MULT_HDR},0)))`, '0.00', GREEN);
  calc(ls.getCell(`H${r}`), `IF(C${r}="","",F${r}*G${r})`, '#,##0.0');
  input(ls.getCell(`I${r}`), null, '#,##0.0');
  calc(ls.getCell(`J${r}`), `IF(I${r}<>"",I${r},H${r})`, '#,##0.0');
  calc(ls.getCell(`K${r}`), `IF(N(J${r})=0,"",J${r}/10.764)`, '0.00');
  calc(ls.getCell(`L${r}`), `IF(OR(D${r}="",N(J${r})=0),0,D${r}*J${r}/1000)`, '#,##0.0;-#,##0.0;"-"');
  const occ = `INDEX(${OCOL},${m})`;
  calc(ls.getCell(`M${r}`), `IF(C${r}="","",IF(${LCODE}="NEC",INDEX(${OCC_NEC},MATCH(${occ},${OCC_LIST},0)),IF(${LCODE}="CEC",INDEX(${OCC_CEC},MATCH(${occ},${OCC_LIST},0)),"n/a")))`, '#,##0.0', GREEN);
  calc(ls.getCell(`N${r}`), `IF(ISNUMBER(M${r}),N(D${r})*M${r}/1000,0)`, '#,##0.0;-#,##0.0;"-"');
  calc(ls.getCell(`O${r}`), `IF(OR(C${r}="",NOT(ISNUMBER(M${r}))),"",IF(M${r}=0,"No code unit load",IF(J${r}>=M${r},"OK – ≥ code lighting load","BELOW code lighting load – review")))`);
  calc(ls.getCell(`P${r}`), `IF(C${r}="","",INDEX(${UCOL},${m}))`, null, GREEN);
  input(ls.getCell(`R${r}`), ex ? 'Example row – replace' : null);
}
const LT = L1 + 1;
const LS_KVA = 'L';
ls.getCell(`A${LT}`).value = 'TOTAL – area-based loads';
ls.mergeCells(`A${LT}:C${LT}`);
calc(ls.getCell(`D${LT}`), `SUM(D${L0}:D${L1})`, '#,##0');
calc(ls.getCell(`E${LT}`), `SUM(E${L0}:E${L1})`, '#,##0');
calc(ls.getCell(`J${LT}`), `IF(D${LT}=0,0,L${LT}*1000/D${LT})`, '#,##0.0');
calc(ls.getCell(`K${LT}`), `IF(D${LT}=0,0,J${LT}/10.764)`, '0.00');
calc(ls.getCell(`L${LT}`), `SUM(L${L0}:L${L1})`, '#,##0.0');
calc(ls.getCell(`N${LT}`), `SUM(N${L0}:N${L1})`, '#,##0.0');
ls.getCell(`I${LT}`).value = 'Average:';
'ABCDEFGHIJKLMNOPQR'.split('').forEach(c => { const x = ls.getCell(`${c}${LT}`); x.fill = totFill; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); x.border = box; });
ls.getCell(`A${LT + 1}`).value = 'Code lighting load (NEC Table 220.12 / CEC Section 8) is a minimum for general lighting only – the design density must normally exceed it. Use the Code Load Calc sheet for the full NEC Article 220 / CEC Section 8 service calculation.';
ls.getCell(`A${LT + 1}`).font = { name: 'Arial', size: 9, italic: true };
// Summary by load group and by use class
const GH = LT + 4;
const gh = ls.getCell(`A${GH - 1}`); gh.value = 'Summary by load group (selected region)'; gh.font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY } };
const sumHdr = (r, label) => ['B', 'D', 'J', 'L'].forEach((c, j) => { const x = ls.getCell(`${c}${r}`); x.value = [label, 'GIA (m²)', 'Avg VA/m²', 'Demand (kVA)'][j]; x.fill = hdrFill; x.border = box; x.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } }; x.alignment = { horizontal: 'center' }; });
sumHdr(GH, 'Load group');
GROUPS.forEach((g, i) => {
  const r = GH + 1 + i;
  calc(ls.getCell(`B${r}`), `${RDQ}!A${MG0 + i}`, null, GREEN); ls.mergeCells(`B${r}:C${r}`);
  calc(ls.getCell(`D${r}`), `SUMIF($Q$${L0}:$Q$${L1},B${r},$D$${L0}:$D$${L1})`, '#,##0;-#,##0;"-"');
  calc(ls.getCell(`L${r}`), `SUMIF($Q$${L0}:$Q$${L1},B${r},$L$${L0}:$L$${L1})`, '#,##0.0;-#,##0.0;"-"');
  calc(ls.getCell(`J${r}`), `IF(D${r}=0,0,L${r}*1000/D${r})`, '#,##0;-#,##0;"-"');
});
const UH = GH + GROUPS.length + 4;
const uh = ls.getCell(`A${UH - 1}`); uh.value = 'Summary by UK use class'; uh.font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY } };
sumHdr(UH, 'Use class');
USE_CLASSES.forEach((u, i) => {
  const r = UH + 1 + i;
  calc(ls.getCell(`B${r}`), `'VA-m2 Data'!O${5 + i}`, null, GREEN); ls.mergeCells(`B${r}:C${r}`);
  calc(ls.getCell(`D${r}`), `SUMIF($P$${L0}:$P$${L1},B${r},$D$${L0}:$D$${L1})`, '#,##0;-#,##0;"-"');
  calc(ls.getCell(`L${r}`), `SUMIF($P$${L0}:$P$${L1},B${r},$L$${L0}:$L$${L1})`, '#,##0.0;-#,##0.0;"-"');
  calc(ls.getCell(`J${r}`), `IF(D${r}=0,0,L${r}*1000/D${r})`, '#,##0;-#,##0;"-"');
});
ls.getColumn(1).width = 7; ls.getColumn(2).width = 28; ls.getColumn(3).width = 46; ls.getColumn(4).width = 10; ls.getColumn(5).width = 10;
'FGHIJK'.split('').forEach(c => ls.getColumn(c).width = 10); ls.getColumn('L').width = 11; ls.getColumn('M').width = 12; ls.getColumn('N').width = 11; ls.getColumn('O').width = 28; ls.getColumn('P').width = 28; ls.getColumn('Q').width = 26; ls.getColumn('R').width = 26;
ls.views = [{ state: 'frozen', xSplit: 3, ySplit: 5 }];

