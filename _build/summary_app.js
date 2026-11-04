// ---------------- 18. Summary ----------------
const sm = wb.addWorksheet('Summary');
title(sm, 'Maximum Demand Summary');
calc(sm.getCell('A2'), `'Project Info'!B6&"  |  Ref: "&'Project Info'!B7&"  |  Rev "&'Project Info'!B11&"  |  Code: "&'Project Info'!B24`, null, GREEN);
legend(sm, 3);
header(sm, 5, ['Item', 'Include?', 'Value', 'Unit', 'Basis / notes']);
const srcs = [
  ['Building Calculator (floor-by-floor)', `${BCQ}!J${FB}`, 'Floors × use density + lift schedule', 'Yes'],
  ['Area-based zones (Load Schedule, regional)', `'Load Schedule'!${LS_KVA}${LT}`, 'Σ GIA × regional design VA/m²', 'No'],
  ['Code calculated load (NEC Art. 220 / CEC Sec. 8)', CC_TOTAL, 'North American code service calculation', 'No'],
  ['Client Load Schedule (detailed, after Ku × DF × Ks)', `${CLQ}!G${DBM}`, 'Actual client equipment data', 'No'],
  ['Residential dwellings (ADMD)', `'Residential ADMD'!D${RT}`, 'Σ dwellings × ADMD', 'No'],
  ['Specific / point loads', `'Specific Loads'!F${ST}`, 'EV, plant, floodlights, process', 'Yes'],
  ['External, street and public-realm lighting', EXT_TOTAL, 'Streets, car parks, parks, open spaces', 'Yes'],
  ['Building Services Systems (system-by-system)', `${SYQ}!P${SYT}`, 'Lighting, sockets, HVAC, chillers, pumps, lifts, alarms', 'No'],
  ['Data centre (IT + UPS + cooling + ancillaries)', DC_MD, 'Data Centre sheet – Tier-based design', 'No'],
];
const SM_S0 = 6, SM_S1 = SM_S0 + srcs.length - 1; // 6..11
srcs.forEach(([label, f, note, def], i) => {
  const r = SM_S0 + i;
  const a = sm.getCell(`A${r}`); a.value = label; a.border = box; a.font = FONT;
  input(sm.getCell(`B${r}`), def); sm.getCell(`B${r}`).alignment = { horizontal: 'center' };
  sm.getCell(`B${r}`).dataValidation = { type: 'list', allowBlank: false, formulae: ['"Yes,No"'] };
  calc(sm.getCell(`C${r}`), f, '#,##0.0', GREEN);
  const u = sm.getCell(`D${r}`); u.value = 'kVA'; u.border = box; u.font = FONT; u.alignment = { horizontal: 'center' };
  const n = sm.getCell(`E${r}`); n.value = note; n.border = box; n.font = { name: 'Arial', size: 9, italic: true };
});
const SM_SUB = SM_S1 + 1, SM_SP = SM_S1 + 2, SM_SPK = SM_S1 + 3, SM_TOT = SM_S1 + 4, SM_KW = SM_S1 + 5, SM_I = SM_S1 + 6, SM_SWB = SM_S1 + 7, SM_TX = SM_S1 + 8, SM_SUP = SM_S1 + 9, SM_A = SM_S1 + 10, SM_D = SM_S1 + 11;
if (SM_TOT !== 18) throw new Error("Summary total row moved: " + SM_TOT);
const res = [
  [SM_SUB, 'Subtotal – estimated maximum demand', `SUMPRODUCT((B${SM_S0}:B${SM_S1}="Yes")*C${SM_S0}:C${SM_S1})`, 'kVA', 'Included items only', null, true],
  [SM_SP, 'Spare / future capacity allowance', `'Project Info'!B20`, '%', 'Set on Project Info', GREEN],
  [SM_SPK, 'Spare capacity', `C${SM_SUB}*C${SM_SP}`, 'kVA', ''],
  [SM_TOT, 'TOTAL DESIGN MAXIMUM DEMAND', `C${SM_SUB}+C${SM_SPK}`, 'kVA', 'Use for utility capacity application', null, true],
  [SM_KW, 'Total demand (kW)', `C${SM_TOT}*${PF}`, 'kW', 'kVA × power factor'],
  [SM_I, 'Design current per phase', `IF('Project Info'!B17=0,0,C${SM_TOT}*1000/(SQRT(3)*'Project Info'!B17))`, 'A', 'I = S ÷ (√3 × V L-L)'],
  [SM_SWB, 'Recommended main switchboard rating', `IF(C${SM_I}>MAX(${SWB}),"> "&MAX(${SWB})&" A – split supplies",INDEX(${SWB},COUNTIF(${SWB},"<"&C${SM_I})+1))`, 'A', 'Next standard rating ≥ design current'],
  [SM_TX, 'Indicative transformer size', `IF(C${SM_TOT}>MAX(${TX}),"Multiple transformers / HV",INDEX(${TX},COUNTIF(${TX},"<"&C${SM_TOT})+1))`, 'kVA', 'See Transformer & Panels for full sizing'],
  [SM_SUP, 'Indicative supply arrangement', `IF(C${SM_TOT}<=69,"LV whole-current service (≤100 A 3-phase)",IF(C${SM_TOT}<=275,"LV CT-metered supply from network",IF(C${SM_TOT}<=1000,"LV supply from dedicated substation (likely)","HV supply / customer-owned substation")))`, '', 'Indicative only – thresholds vary by utility'],
  [SM_A, 'Total floor area (included area-based sources)', `IF(B6="Yes",${BCQ}!E${FB},0)+IF(B7="Yes",'Load Schedule'!D${LT},0)`, 'm²', ''],
  [SM_D, 'Overall load density (included area-based sources, before spare)', `IF(C${SM_A}=0,0,(IF(B6="Yes",C6,0)+IF(B7="Yes",C7,0))*1000/C${SM_A})`, 'VA/m²', ''],
];
res.forEach(([r, label, f, unit, note, color, bold]) => {
  const a = sm.getCell(`A${r}`); a.value = label; a.border = box; a.font = { name: 'Arial', size: 10, bold: !!bold };
  sm.getCell(`B${r}`).border = box;
  calc(sm.getCell(`C${r}`), f, '#,##0.0', color);
  const u = sm.getCell(`D${r}`); u.value = unit; u.border = box; u.font = FONT; u.alignment = { horizontal: 'center' };
  const n = sm.getCell(`E${r}`); n.value = note; n.border = box; n.font = { name: 'Arial', size: 9, italic: true };
  if (bold) ['A', 'B', 'C', 'D', 'E'].forEach(c => { sm.getCell(`${c}${r}`).fill = totFill; sm.getCell(`${c}${r}`).font = Object.assign({}, sm.getCell(`${c}${r}`).font, { bold: true }); });
});
sm.getCell(`C${SM_SP}`).numFmt = '0%';
[SM_SWB, SM_TX, SM_A, SM_D].forEach(r => sm.getCell(`C${r}`).numFmt = '#,##0');
sm.getCell(`C${SM_SUP}`).alignment = { wrapText: true };
sm.getCell('A27').value = 'Set "Include?" to No for any source that duplicates floor area or loads entered elsewhere (e.g. the same building in both the Building Calculator and the Client Load Schedule).';
sm.getCell('A28').value = 'Early-stage estimates use rule-of-thumb densities; at detailed design use the Client Load Schedule with actual plant ratings and the diversity rules of the selected regional code.';
sm.getCell('A27').font = { name: 'Arial', size: 9, italic: true };
sm.getCell('A28').font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FFC00000' } };
sm.getColumn(1).width = 56; sm.getColumn(2).width = 10; sm.getColumn(3).width = 40; sm.getColumn(4).width = 8; sm.getColumn(5).width = 44;

// ---------------- 19. Supply Application ----------------
const sa = wb.addWorksheet('Supply Application');
title(sa, 'Utility / DNO Supply Application – Data Sheet', 'Collates the information normally requested on a new-connection application. Transfer to the utility\'s own form or portal. Green cells link to the calculations.');
legend(sa, 3);
const ssec = (r, t) => { const c = sa.getCell(`A${r}`); c.value = t; c.font = { name: 'Arial', size: 10, bold: true }; ['A', 'B', 'C'].forEach(x => sa.getCell(`${x}${r}`).fill = secFill); };
const srow = (r, label, kind, v, fmt, note) => {
  const a = sa.getCell(`A${r}`); a.value = label; a.border = box; a.font = FONT;
  if (kind === 'in') input(sa.getCell(`B${r}`), v, fmt); else calc(sa.getCell(`B${r}`), v, fmt, GREEN);
  sa.getCell(`B${r}`).alignment = { wrapText: true, horizontal: 'left' };
  if (note) { const n = sa.getCell(`C${r}`); n.value = note; n.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } }; }
};
ssec(5, 'A. APPLICANT AND SITE');
srow(6, 'Project name', 'f', `'Project Info'!B6`);
srow(7, 'Client / applicant', 'f', `'Project Info'!B8`);
srow(8, 'Site address / postcode', 'f', `'Project Info'!B9`);
srow(9, 'Utility / distribution company', 'f', `'Project Info'!B10`);
srow(10, 'Applicable wiring regulations', 'f', `'Project Info'!B24`);
srow(11, 'Contact name / telephone / email', 'in', '', null, 'Applicant or agent');
ssec(13, 'B. CONNECTION REQUEST');
srow(14, 'Required maximum import capacity (kVA)', 'f', `Summary!C${SM_TOT}`, '#,##0', 'Includes spare capacity');
srow(15, 'Maximum demand (kW)', 'f', `Summary!C${SM_KW}`, '#,##0');
srow(16, 'Design current per phase (A)', 'f', `Summary!C${SM_I}`, '#,##0');
srow(17, 'Supply voltage (V, L-L / L-N)', 'f', `'Project Info'!B17&" / "&'Project Info'!B18&" V, "&'Project Info'!B23&" Hz"`);
srow(18, 'Number of phases', 'in', 3, '0');
srow(19, 'Number of supply points / meters', 'in', 1, '0');
srow(20, 'Earthing requested', 'f', `'Project Info'!B25`);
srow(21, 'Indicative supply arrangement', 'f', `Summary!C${SM_SUP}`);
srow(22, 'Transformer (if customer substation)', 'f', `'Transformer & Panels'!B11&" kVA, "&'Transformer & Panels'!B17`);
srow(23, 'Required energisation date', 'in', '', 'dd/mm/yyyy');
srow(24, 'Phased load build-up (kVA by year)', 'in', 'Year 1: __ / Year 2: __ / Year 3: __', null, 'Utilities may connect in phases');
ssec(26, 'C. LOAD TYPES (kW unless stated)');
const lt = [
  [27, 'Heat pumps (total)', 'in', 0, '#,##0.0', ''],
  [28, 'Direct electric space heating', 'in', 0, '#,##0.0', ''],
  [29, 'Electric water heating', 'in', 0, '#,##0.0', ''],
  [30, 'EV charge points – number', 'in', 0, '0', ''],
  [31, 'EV charge points – total rating', 'in', 0, '#,##0.0', ''],
  [32, 'EV load management fitted?', 'in', 'Yes', null, 'IET Code of Practice for EV charging'],
  [33, 'Largest motor (kW) and starting method', 'in', '__ kW, VSD / soft starter / star-delta / DOL', null, 'Motor starting voltage dip assessment'],
  [34, 'Lifts – number × rating', 'f', `${BCQ}!B13&" × "&${BCQ}!B14&" kW"`, null, 'From Building Calculator'],
  [35, 'Harmonic-producing loads (VSD, UPS, LED drivers)', 'in', 0, '#,##0.0', 'UK: ENA EREC G5/5 assessment'],
  [36, 'Disturbing / fluctuating loads (welders, X-ray, crushers)', 'in', 0, '#,##0.0', 'UK: ENA EREC P28 flicker assessment'],
  [37, 'External / street lighting (kW)', 'f', `${EXQ}!N${ET}`, '#,##0.0', 'From External Lighting'],
];
lt.forEach(([r, l, k, v, f, n]) => srow(r, l, k, v, f || undefined, n));
sa.getCell('B32').dataValidation = { type: 'list', allowBlank: false, formulae: ['"Yes,No"'] };
ssec(39, 'D. GENERATION AND STORAGE');
srow(40, 'Solar PV (kWp)', 'in', 0, '#,##0.0');
srow(41, 'Battery storage (kW / kWh)', 'in', '0 kW / 0 kWh');
srow(42, 'Battery storage inverter rating (kW)', 'in', 0, '#,##0.0');
srow(43, 'Standby generator (kVA) – not parallel with grid', 'f', `'Transformer & Panels'!B${GS + 4}`, '#,##0');
srow(44, 'Export required?', 'in', 'No', null, '');
sa.getCell('B44').dataValidation = { type: 'list', allowBlank: false, formulae: ['"Yes,No"'] };
srow(45, 'Total grid-parallel generation (kW)', 'f', `N(B40)+N(B42)`, '#,##0.0', 'PV + battery inverter');
srow(46, 'Generation connection category', 'f', `IF(B45=0,"None",IF('Project Info'!B22="United Kingdom",IF(B45<=IF(B18=1,3.68,11.04),"EREC G98 (≤ 16 A per phase)","EREC G99"),"Check local utility generation rules"))`, null, 'UK: G98 ≤ 16 A/phase, otherwise G99');
ssec(48, 'E. ATTACHMENTS CHECKLIST');
const att = ['Location plan (1:1250) with site boundary', 'Site plan showing proposed intake / substation position', 'Load schedule and maximum demand calculation (this workbook)', 'Single line diagram', 'Phasing / build programme', 'Generation and storage details (if applicable)', 'Substation design / access arrangements (if applicable)'];
att.forEach((t, i) => {
  const r = 49 + i;
  const a = sa.getCell(`A${r}`); a.value = t; a.border = box; a.font = FONT;
  input(sa.getCell(`B${r}`), 'No'); sa.getCell(`B${r}`).dataValidation = { type: 'list', allowBlank: false, formulae: ['"Yes,No,N/A"'] };
});
sa.getColumn('A').width = 52; sa.getColumn('B').width = 46; sa.getColumn('C').width = 46;

// ---------------- 20. References ----------------
const rf = wb.addWorksheet('References & Notes');
title(rf, 'References, Assumptions and How to Use');
const refs = [
  ['HOW TO USE', true],
  ['1. Project Info: enter project details and select the region / design code (sets voltage, frequency, VD limits and standards).'],
  ['1a. Load Schedule: pick the design region at the top (UK, Europe, USA, Canada, Australia/NZ, Middle East, Africa, Asia). USA/Canada: also complete Code Load Calc (NEC Art. 220 / CEC Sec. 8).'],
  ['2. Choose a load-estimate route: Building Calculator (floors), Building Services Systems (system-by-system, with Pumps & Lifts sizing), Load Schedule (zones, all UK use classes), Client Load Schedule (actual equipment), Residential ADMD, Specific Loads, External Lighting.'],
  ['3. Summary: set "Include?" for each route so no load is counted twice; read the total maximum demand.'],
  ['4. Lighting Design: confirm lux levels and luminaire numbers per space; compare installed W/m² with benchmarks.'],
  ['5. Cable & Protection: size each circuit / sub-main (Ib ≤ In ≤ Iz, voltage drop, Zs). Transformer & Panels: size transformer, main switchboard, DBs and standby plant.'],
  ['6. IET Diversity: OSG Appendix A calculator for small installations (dwellings, small shops, small hotels).'],
  ['7. Supply Application: transfer the collated data to the utility / DNO application.'],
  ['6a. Data Centre: choose Tier and sizing method (white-space m², racks or IT kW); redundancy defaults follow the Tier and can be overridden. Include on Summary when the project contains a data hall.'],
  ['7a. Energy Management: tick the energy-saving measures, PV / battery and demand management; review energy, cost, carbon and power-factor correction.'],
  ['8. Compliance Checklist: review automatic results, then record the engineer status and evidence for every item before issue.'],
  [''],
  ['REFERENCES', true],
  ['BS 7671:2018 Requirements for Electrical Installations (IET Wiring Regulations, 18th Edition) as amended; IET On-Site Guide; IET Guidance Notes 1 and 3'],
  ['BSRIA BG 9/2011 Rules of Thumb; CIBSE Guide F (Energy Efficiency) and Guide K (Electricity in Buildings); CIBSE TM54; BCO Guide to Specification'],
  ['BS EN 12464-1:2021 (indoor workplaces) and BS EN 12464-2 (outdoor workplaces); SLL Code for Lighting; BS 5489-1 and BS EN 13201 (road lighting); BS EN 12193 (sports lighting)'],
  ['BS 5266-1 and BS EN 1838 (emergency lighting); BS EN 61439 (LV assemblies); BS EN 60898-1, BS EN 61009-1, BS EN 60947-2 (protective devices)'],
  ['BS EN / IEC 60076 (power transformers); ISO 8528 and BS 7698 (generating sets); BS EN 62040 (UPS); HTM 06-01 (healthcare electrical services)'],
  ['ENA Engineering Recommendations G81, G98, G99, G5/5, P28; DNO ADMD design policies; Town and Country Planning (Use Classes) Order 1987 as amended 2020'],
  ['Regional codes listed on the Regional Codes sheet (IEC 60364, I.S. 10101, DEWA, SBC 401, KAHRAMAA, SANS 10142-1, IS 732, SS 638, AS/NZS 3000, NFPA 70, CSA C22.1 etc.)'],
  [''],
  ['ASSUMPTIONS', true],
  ['1. All VA/m² densities are after diversity and apply to Gross Internal Area (GIA) unless noted; they are indicative rules of thumb, not code requirements.'],
  ['2. VA = W ÷ power factor; overall power factor set on Project Info (default 0.95).'],
  ['3. Spare capacity is applied to the total of the included sources on the Summary.'],
  ['4. Residential ADMD values already include development-wide diversity.'],
  ['5. Cable ratings are BS 7671 Appendix 4 values for copper conductors; verify against the current edition and use local tables for NEC / CEC / AS/NZS regions.'],
  ['6. Transformer fault levels assume an infinite source; obtain the actual network fault level from the utility.'],
  ['7. Switchboard, transformer and supply-arrangement outputs are indicative; the utility determines the point of connection and supply arrangement.'],
  ['8. Example values on every input sheet are illustrative and must be replaced with client / project data.'],
];
refs.forEach(([t, s], i) => { const c = rf.getCell(`A${3 + i}`); c.value = t; c.font = { name: 'Arial', size: 10, bold: !!s }; c.alignment = { wrapText: true }; if (s) c.fill = secFill; });
rf.getColumn(1).width = 150;

// ---------------- 19a. Compliance Checklist ----------------
const ck = wb.addWorksheet('Compliance Checklist');
title(ck, 'Design Compliance Checklist – by Regional Code', 'Automatic checks summarise the workbook results; every item still needs the designer\'s confirmation. Clause references are for guidance – confirm against the edition adopted locally. This checklist does not replace certification by a competent person.');
legend(ck, 3);
ck.getCell('A5').value = 'Project region'; ck.getCell('A5').font = { name: 'Arial', size: 10, bold: true };
calc(ck.getCell('B5'), `'Project Info'!B22&"  →  density region: "&'Project Info'!B35&"  |  "&'Project Info'!B24`, null, GREEN); ck.mergeCells('B5:F5');
ck.getCell('A6').value = 'Code family used for references'; ck.getCell('A6').font = { name: 'Arial', size: 10, bold: true };
calc(ck.getCell('B6'), `CHOOSE(MATCH('Project Info'!B36,{"IEC","NEC","CEC","ASNZS"},0),"BS 7671 / IEC 60364 family (+ local regulations)","NFPA 70 – NEC family (+ local adoption)","CSA C22.1 – CEC","AS/NZS 3000")`, null, GREEN); ck.mergeCells('B6:F6');
const FAM = `MATCH('Project Info'!$B$36,{"IEC","NEC","CEC","ASNZS"},0)`;
header(ck, 8, ['No.', 'Requirement / design check', 'Regulation reference (selected region)', 'Workbook result', 'Auto status', 'Engineer status', 'Engineer notes / evidence']);
ck.getRow(8).height = 30;
const CPQ = "'Cable & Protection'";
const chk = (pat) => `COUNTIF(${CPQ}!${K['Check']}${CP0}:${K['Check']}${CP1},"${pat}")`;
const ITEMS = [
  ['Supply characteristics assessed (voltage, frequency, earthing, PFC)', ['313.1; 312', 'NEC 110.4; utility service requirements', 'CEC Rule 2-100; utility', 'AS/NZS 3000 Clause 2.2.1'],
    `'Project Info'!B17&"/"&'Project Info'!B18&" V, "&'Project Info'!B23&" Hz, "&'Project Info'!B25`, `IF(N('Project Info'!B17)>0,"OK","CHECK")`],
  ['Maximum demand assessed with diversity / demand factors', ['311.1; IET OSG Appendix A', 'NEC Article 220', 'CEC Rules 8-200 to 8-210', 'AS/NZS 3000 Clause 2.2.2; Appendix C'],
    `TEXT(Summary!C${SM_TOT},"#,##0")&" kVA (incl. spare)"`, `IF(Summary!C${SM_TOT}>0,"OK","CHECK")`],
  ['Code-prescribed calculated load (where required)', ['Not prescribed – rules of thumb / ADMD', 'NEC 220.12, 220.42, 220.47 – Code Load Calc sheet', 'CEC Section 8 – Code Load Calc sheet', 'Appendix C (assessment)'],
    `IF(OR('Project Info'!B36="NEC",'Project Info'!B36="CEC"),TEXT(${CC_TOTAL},"#,##0")&" kVA calculated; "&COUNTIF('Load Schedule'!$O$6:$O$55,"BELOW*")&" zones below code lighting load","Not applicable")`,
    `IF(OR('Project Info'!B36="NEC",'Project Info'!B36="CEC"),IF(COUNTIF('Load Schedule'!$O$6:$O$55,"BELOW*")=0,"OK","ACTION"),"N/A")`],
  ['Current-carrying capacity: Ib ≤ In ≤ Iz with correction factors', ['433.1.1; Appendix 4', 'NEC 310.14, 310.16; 240.4', 'CEC Rule 4-004; Tables 1–4; 14-104', 'AS/NZS 3000 Clause 3.4; AS/NZS 3008.1'],
    `${chk('FAIL – Iz*')}&" circuit(s) failing"`, `IF(${chk('FAIL – Iz*')}=0,"OK","ACTION")`],
  ['Voltage drop within limits', ['525; Appendix 4 Table 4Ab', 'NEC 210.19(A) / 215.2 informational notes', 'CEC Rule 8-102', 'AS/NZS 3000 Clause 3.6'],
    `${chk('FAIL – voltage*')}&" circuit(s) failing; limits "&'Project Info'!B26&"% / "&'Project Info'!B27&"%"`, `IF(${chk('FAIL – voltage*')}=0,"OK","ACTION")`],
  ['Automatic disconnection – earth fault loop impedance', ['411.3.2; Table 41.3; 411.4', 'NEC 250.4(A)(5); 250.122', 'CEC Section 10; Rule 14-012', 'AS/NZS 3000 Clause 5.7; 8.3.9'],
    `${chk('FAIL – Zs*')}&" failing; "&${chk('PASS – ADS*')}&" relying on RCD"`, `IF(${chk('FAIL – Zs*')}=0,"OK","ACTION")`],
  ['Protective / earthing conductor sizing', ['543.1; Table 54.7; 542–544', 'NEC Table 250.122; 250.66', 'CEC Table 16; Rule 10-614', 'AS/NZS 3000 Clause 5.3; Table 5.1'],
    `${chk('FAIL – cpc*')}&" cpc failures; main earthing "&'Transformer & Panels'!B${ES + 3}&" mm², bonding "&'Transformer & Panels'!B${ES + 4}&" mm²"`, `IF(${chk('FAIL – cpc*')}=0,"OK","ACTION")`],
  ['Breaking capacity ≥ prospective fault current', ['434.5.1; 536.1', 'NEC 110.9; 110.10', 'CEC Rule 14-012', 'AS/NZS 3000 Clause 2.5.5'],
    `${chk('FAIL – breaking*')}&" failing; LV fault level "&TEXT('Transformer & Panels'!B16,"0.0")&" kA"`, `IF(${chk('FAIL – breaking*')}=0,"OK","ACTION")`],
  ['Selectivity / discrimination of protective devices', ['536.4', 'NEC 240.12; 700.32; 701.32', 'CEC Rule 14-012; Section 46', 'AS/NZS 3000 Clause 2.5.7'], `"Manufacturer selectivity tables"`, `"MANUAL"`],
  ['Additional protection – RCD / GFCI for sockets and specified locations', ['411.3.3; 411.3.4; 522.6.202; 701', 'NEC 210.8 (GFCI)', 'CEC Rule 26-704', 'AS/NZS 3000 Clause 2.6.3'], `'Transformer & Panels'!B${GSP + 3}`, `"MANUAL"`],
  ['Arc-fault protection (AFDD / AFCI)', ['421.1.7', 'NEC 210.12 (AFCI)', 'CEC Rule 26-722', 'Risk assessment (not mandatory)'], `'Transformer & Panels'!B${GSP + 5}`, `"MANUAL"`],
  ['Surge protection', ['Section 443; 534', 'NEC 230.67; Article 242', 'CEC – confirm applicable rule', 'AS/NZS 3000 Clause 2.9; AS/NZS 1768'], `'Transformer & Panels'!B${GSP + 6}`, `"MANUAL"`],
  ['Isolation, switching and emergency stopping', ['Chapter 46; 537', 'NEC 430 Part IX; 404', 'CEC Rule 28-600', 'AS/NZS 3000 Clause 2.3'], `"See Pumps & Lifts / specification"`, `"MANUAL"`],
  ['Lighting levels to the regional lighting standard', ['BS EN 12464-1:2021', 'IES Lighting Handbook / RP series', 'IES / CSA', 'AS/NZS 1680 series'],
    `SUMPRODUCT(('Lighting Design'!$C$10:$C$49<>"")*('Lighting Design'!$Q$10:$Q$49<'Lighting Design'!$I$10:$I$49))&" space(s) below design lux"`, `IF(SUMPRODUCT(('Lighting Design'!$C$10:$C$49<>"")*('Lighting Design'!$Q$10:$Q$49<'Lighting Design'!$I$10:$I$49))=0,"OK","ACTION")`],
  ['Lighting power density / energy code', ['Building Regulations Part L (England) / national', 'ASHRAE 90.1 / IECC', 'NECB', 'NCC Section J'],
    `COUNTIF('Lighting Design'!$R$10:$R$49,"Above*")&" space(s) above W/m² benchmark"`, `IF(COUNTIF('Lighting Design'!$R$10:$R$49,"Above*")=0,"OK","ACTION")`],
  ['Emergency / escape lighting', ['BS 5266-1; BS EN 1838', 'NEC Art. 700; NFPA 101 §7.9', 'CEC Section 46; NBC 3.2.7', 'AS/NZS 2293.1'], `"See Lighting Data – emergency table"`, `"MANUAL"`],
  ['Life-safety and fire-fighting supplies (fire pumps, smoke fans, fire lifts)', ['Chapter 56; BS 8519; BS EN 12845', 'NEC Art. 695; 700', 'CEC Section 32; Section 46', 'AS/NZS 3000 Clause 7.2'], `"Fire pump: "&'Pumps & Lifts'!M13`, `"MANUAL"`],
  ['Transformer loading and switchgear ratings', ['BS EN 60076; BS EN 61439-2', 'NEC Art. 450; UL 891', 'CEC Section 26; CSA C22.2 No. 47', 'AS 60076; AS/NZS 61439'],
    `'Transformer & Panels'!B8&" × "&'Transformer & Panels'!B11&" kVA; Icw "&'Transformer & Panels'!B24&" kA"`, `IF(ISNUMBER('Transformer & Panels'!B11),"OK","ACTION")`],
  ['Substation / switchroom space, access and ventilation', ['729; ENA TS 41-24; Building Regs', 'NEC 110.26; 110.31; 450.21', 'CEC Rules 2-308, 2-310; 26-240', 'AS 2067; AS/NZS 3000 Clause 2.10'],
    `TEXT('Transformer & Panels'!B${TS + 18},"0")&" m² total; vent "&TEXT('Transformer & Panels'!B${TS + 21},"0.00")&" m³/s"`, `"MANUAL"`],
  ['EV charging installation', ['Section 722; IET Code of Practice', 'NEC Article 625; 220.57', 'CEC Section 86', 'AS/NZS 3000 (EV provisions); AS IEC 61851'], `"See Specific Loads / Building Services Systems"`, `"MANUAL"`],
  ['Embedded generation / storage connection', ['ENA EREC G98 / G99; Sections 551, 712, 826', 'NEC Art. 690, 705, 706; IEEE 1547', 'CEC Sections 64, 84', 'AS/NZS 4777; AS/NZS 5033; AS/NZS 5139'], `'Supply Application'!B46`, `"MANUAL"`],
  ['Utility connection application data complete', ['ENA G81 / DNO process', 'Utility service requirements; NEC Art. 230', 'Utility; CEC Section 6', 'State Service & Installation Rules'],
    `COUNTIF('Supply Application'!B49:B55,"Yes")&" of 7 attachments ready"`, `IF(COUNTIF('Supply Application'!B49:B55,"No")=0,"OK","ACTION")`],
  ['Data centre resilience and design (where applicable)', ['EN 50600-2-2; Uptime Tier Standard; BICSI 002', 'ANSI/TIA-942-B; NEC Art. 645; NFPA 75', 'TIA-942-B; CSA C22.1 – confirm applicable rules', 'AS/NZS 3000; TIA-942-B'],
    `${DCQ}!B6&"; IT "&TEXT(${DCQ}!B20,"#,##0")&" kW; peak PUE "&TEXT(${DCQ}!B60,"0.00")`, `"MANUAL"`],
  ['Energy metering, monitoring and management', ['Building Regulations Part L; ISO 50001; ISO 52120-1 (BACS)', 'ASHRAE 90.1 §8.4.3 / IECC – confirm edition', 'NECB – confirm edition', 'NCC Section J – energy monitoring'],
    `TEXT(${EMQ}!B${EM_EUI},"0")&" kWh/m²/yr net; MD "&TEXT(${EMQ}!B${EM_MDA},"#,##0")&" kVA after demand management"`, `IF(${EMQ}!B${EM_EUI}<=${EMQ}!B${EM_BM},"OK","ACTION")`],
  ['Inspection, testing and certification', ['BS 7671 Part 6; EIC / MWC', 'AHJ inspection; NEC 90.4', 'CEC Rule 2-004 (permit / inspection)', 'AS/NZS 3000 Section 8; AS/NZS 3017 / 3019'], `"Required at completion"`, `"MANUAL"`],
];
const CK0 = 9;
ITEMS.forEach(([req, refs, res, st], i) => {
  const r = CK0 + i;
  const a = ck.getCell(`A${r}`); a.value = i + 1; a.border = box; a.font = FONT; a.alignment = { horizontal: 'center', vertical: 'top' };
  const b = ck.getCell(`B${r}`); b.value = req; b.border = box; b.font = FONT; b.alignment = { wrapText: true, vertical: 'top' };
  calc(ck.getCell(`C${r}`), `CHOOSE(${FAM},"${refs.join('","')}")`); ck.getCell(`C${r}`).alignment = { wrapText: true, vertical: 'top' }; ck.getCell(`C${r}`).font = { name: 'Arial', size: 9 };
  calc(ck.getCell(`D${r}`), res); ck.getCell(`D${r}`).alignment = { wrapText: true, vertical: 'top' }; ck.getCell(`D${r}`).font = { name: 'Arial', size: 9 };
  calc(ck.getCell(`E${r}`), st); ck.getCell(`E${r}`).alignment = { horizontal: 'center', vertical: 'top' }; ck.getCell(`E${r}`).font = { name: 'Arial', size: 10, bold: true };
  input(ck.getCell(`F${r}`), 'Not checked'); ck.getCell(`F${r}`).dataValidation = { type: 'list', allowBlank: false, formulae: ['"Not checked,Complies,Complies with departure,Action required,Not applicable"'] };
  input(ck.getCell(`G${r}`), null); ck.getCell(`G${r}`).alignment = { wrapText: true, vertical: 'top' };
  ck.getRow(r).height = 30;
});
const CK1 = CK0 + ITEMS.length - 1;
ck.addConditionalFormatting({ ref: `E${CK0}:E${CK1}`, rules: [
  { type: 'cellIs', operator: 'equal', formulae: ['"OK"'], style: { font: { color: { argb: 'FF006100' } }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFC6EFCE' } } } },
  { type: 'cellIs', operator: 'equal', formulae: ['"ACTION"'], style: { font: { color: { argb: 'FF9C0006' } }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } },
  { type: 'cellIs', operator: 'equal', formulae: ['"CHECK"'], style: { font: { color: { argb: 'FF9C5700' } }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFEB9C' } } } },
] });
const CKS = CK1 + 2;
[['Automatic checks – OK', `COUNTIF(E${CK0}:E${CK1},"OK")`], ['Automatic checks – ACTION', `COUNTIF(E${CK0}:E${CK1},"ACTION")`], ['Manual items', `COUNTIF(E${CK0}:E${CK1},"MANUAL")`], ['Engineer status – not yet checked', `COUNTIF(F${CK0}:F${CK1},"Not checked")`], ['Engineer status – action required', `COUNTIF(F${CK0}:F${CK1},"Action required")`]].forEach(([l, f], i) => {
  const a = ck.getCell(`B${CKS + i}`); a.value = l; a.border = box; a.font = { name: 'Arial', size: 10, bold: true };
  calc(ck.getCell(`D${CKS + i}`), f, '0');
});
ck.getCell(`B${CKS + 6}`).value = 'Designer: ____________________   Signature: ____________   Date: __________      Checker: ____________________   Date: __________';
ck.getCell(`B${CKS + 6}`).font = { name: 'Arial', size: 10 };
ck.getColumn('A').width = 6; ck.getColumn('B').width = 52; ck.getColumn('C').width = 46; ck.getColumn('D').width = 46; ck.getColumn('E').width = 11; ck.getColumn('F').width = 22; ck.getColumn('G').width = 40;
ck.views = [{ state: 'frozen', ySplit: 8 }];


// ---------------- Sheet order ----------------
const ORDER = ['Project Info', 'Summary', 'Regional Codes', 'Building Calculator', 'Building Services Systems', 'Pumps & Lifts', 'Data Centre', 'Load Schedule', 'Code Load Calc', 'Client Load Schedule', 'Residential ADMD', 'Specific Loads', 'External Lighting', 'Lighting Design', 'Cable & Protection', 'Transformer & Panels', 'Energy Management', 'Supply Application', 'Compliance Checklist', 'IET Diversity', 'Diversity Factors', 'Service Breakdown', 'VA-m2 Data', 'Regional Densities', 'Lighting Data', 'Cable Data', 'References & Notes'];
ORDER.forEach((n, i) => { const ws = wb.getWorksheet(n); if (ws) ws.orderNo = i; });
wb.worksheets.forEach(ws => { ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 }; });
wb.views = [{ activeTab: 0 }];

wb.xlsx.writeFile(OUT).then(() => console.log('written', OUT, wb.worksheets.map(w => w.name).join(' | ')));
