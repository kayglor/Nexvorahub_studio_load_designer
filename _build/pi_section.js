// ---------------- 0. Regional Codes ----------------
const rc = wb.addWorksheet('Regional Codes');
title(rc, 'Regional Design Codes and Supply Parameters', 'Select the region on Project Info. Voltage, frequency, voltage-drop limits and the applicable standards are looked up from this table. Edit any blue value to suit local utility requirements.');
legend(rc, 3);
header(rc, 5, ['Region / country', 'Wiring regulations', 'Voltage L-N (V)', 'Voltage L-L (V)', 'Frequency (Hz)', 'Typical earthing', 'VD limit – lighting (%)', 'VD limit – other (%)', 'Lighting standard', 'Emergency lighting', 'Assemblies / switchgear', 'Protective devices', 'Load calculation method', 'Supply authority / connection', 'Demand / diversity factor set', 'Load density region', 'Code family']);
rc.getRow(5).height = 42;
const REG = [
  ['United Kingdom', 'BS 7671:2018 (18th Edition) as amended', 230, 400, 50, 'TN-C-S (PME) / TN-S', 3, 5, 'BS EN 12464-1 / SLL Code', 'BS 5266-1 / BS EN 1838', 'BS EN 61439', 'BS EN 60898-1 / BS EN 60947-2', 'VA/m² + ADMD; IET OSG App. A', 'DNO / IDNO – ENA G81, G98/G99', 'UK / IEC (BS 7671 – IET OSG)', 'United Kingdom'],
  ['Ireland', 'I.S. 10101:2020 (NSAI)', 230, 400, 50, 'TN-C-S', 3, 5, 'I.S. EN 12464-1', 'I.S. 3217', 'I.S. EN 61439', 'I.S. EN 60898-1', 'VA/m²; ESB Networks guidance', 'ESB Networks', 'UK / IEC (BS 7671 – IET OSG)', 'Europe (IEC)'],
  ['European Union (IEC)', 'IEC 60364 / HD 60364 (national adoption)', 230, 400, 50, 'TN-S / TT', 3, 5, 'EN 12464-1', 'EN 1838 / EN 50172', 'EN 61439', 'EN 60898-1 / EN 60947-2', 'IEC 60364-5-52; national rules', 'National DSO', 'UK / IEC (BS 7671 – IET OSG)', 'Europe (IEC)'],
  ['United Arab Emirates', 'DEWA / ADDC / SEWA Regulations for Electrical Installations', 230, 400, 50, 'TN-S', 4, 4, 'EN 12464-1 / CIBSE', 'UAE Fire & Life Safety Code', 'IEC 61439', 'IEC 60898-1 / IEC 60947-2', 'Authority connected-load & diversity tables', 'DEWA / ADDC / SEWA / FEWA', 'Middle East (DEWA / ADDC / SBC / KAHRAMAA)', 'Middle East (GCC)'],
  ['Saudi Arabia', 'SBC 401 – Saudi Building Code (Electrical)', 230, 400, 60, 'TN-S', 3, 5, 'SBC 401 / EN 12464-1', 'SBC 801', 'IEC 61439', 'IEC 60898-1 / IEC 60947-2', 'SBC 401 demand factors', 'Saudi Electricity Company (SEC)', 'Middle East (DEWA / ADDC / SBC / KAHRAMAA)', 'Middle East (GCC)'],
  ['Qatar', 'KAHRAMAA Regulations for the Installation of Electrical Wiring', 240, 415, 50, 'TN-S', 3, 5, 'EN 12464-1 / CIBSE', 'QCDD requirements', 'IEC 61439', 'IEC 60898-1 / IEC 60947-2', 'KAHRAMAA load & diversity guidance', 'KAHRAMAA', 'Middle East (DEWA / ADDC / SBC / KAHRAMAA)', 'Middle East (GCC)'],
  ['Nigeria', 'NESIS (NEMSA) / BS 7671 practice', 230, 415, 50, 'TN-S / TT', 3, 5, 'BS EN 12464-1', 'BS 5266-1', 'IEC 61439', 'IEC 60898-1 / IEC 60947-2', 'BS 7671 / IET method', 'DisCo (NERC regulated)', 'Africa (SANS 10142 / NRS 034 / BS 7671)', 'Africa'],
  ['Ghana', 'Ghana Electrical Wiring Regulations (L.I. 2008)', 230, 400, 50, 'TN-S / TT', 3, 5, 'BS EN 12464-1', 'BS 5266-1', 'IEC 61439', 'IEC 60898-1', 'BS 7671 / IET method', 'ECG / NEDCo', 'Africa (SANS 10142 / NRS 034 / BS 7671)', 'Africa'],
  ['Kenya', 'KS 662 / BS 7671 practice', 240, 415, 50, 'TN-S / TT', 3, 5, 'BS EN 12464-1', 'BS 5266-1', 'IEC 61439', 'IEC 60898-1', 'BS 7671 / IET method', 'Kenya Power (KPLC)', 'Africa (SANS 10142 / NRS 034 / BS 7671)', 'Africa'],
  ['South Africa', 'SANS 10142-1', 230, 400, 50, 'TN-C-S', 5, 5, 'SANS 10114-1', 'SANS 10114-2', 'SANS 61439', 'SANS 60898 / SANS 60947-2', 'NRS 034 ADMD', 'Eskom / municipality', 'Africa (SANS 10142 / NRS 034 / BS 7671)', 'Africa'],
  ['India', 'IS 732 / NBC 2016 Part 8 / CEA Regulations', 230, 415, 50, 'TN-S / TT', 3, 5, 'IS 3646 / NBC Part 8', 'NBC Part 4', 'IS/IEC 61439', 'IS/IEC 60898-1 / IS/IEC 60947-2', 'NBC Part 8 load densities', 'State DISCOM', 'South & SE Asia (NBC India / SS 638 / MS IEC)', 'South & SE Asia'],
  ['Singapore', 'SS 638:2018 Code of Practice for Electrical Installations', 230, 400, 50, 'TT / TN-S', 4, 4, 'SS 531', 'SS 563', 'IEC 61439', 'IEC 60898-1 / IEC 60947-2', 'SS 638 / SP Group guidance', 'SP Group', 'South & SE Asia (NBC India / SS 638 / MS IEC)', 'South & SE Asia'],
  ['Hong Kong', 'EMSD Code of Practice for the Electricity (Wiring) Regulations', 220, 380, 50, 'TT / TN-S', 4, 4, 'CIBSE / EN 12464-1', 'FSD Code of Practice', 'IEC 61439', 'IEC 60898-1 / IEC 60947-2', 'EMSD CoP / utility guidance', 'CLP Power / HK Electric', 'South & SE Asia (NBC India / SS 638 / MS IEC)', 'South & SE Asia'],
  ['Malaysia', 'MS IEC 60364 / Energy Commission guidelines', 230, 400, 50, 'TT', 4, 4, 'MS 1525 / EN 12464-1', 'UBBL / MS IEC 60598-2-22', 'MS IEC 61439', 'MS IEC 60898-1', 'TNB Electricity Supply Application Handbook', 'Tenaga Nasional Berhad (TNB)', 'South & SE Asia (NBC India / SS 638 / MS IEC)', 'South & SE Asia'],
  ['Australia', 'AS/NZS 3000:2018 + state regulations (cables AS/NZS 3008.1)', 230, 400, 50, 'MEN (TN-C-S)', 5, 5, 'AS/NZS 1680', 'AS/NZS 2293', 'AS/NZS 61439', 'AS/NZS 60898', 'AS/NZS 3000 Appendix C', 'DNSP (local distributor)', 'Australia / NZ (AS/NZS 3000 App. C)', 'Australia / NZ'],
  ['United States', 'NFPA 70 – National Electrical Code (NEC)', 277, 480, 60, 'Solidly grounded wye', 3, 5, 'IES / ASHRAE 90.1 LPD', 'NFPA 101 / NEC Art. 700', 'UL 891 / UL 67', 'UL 489', 'NEC Article 220 (VA/ft²)', 'Local utility', 'North America (NEC Art. 220 / CEC Sec. 8)', 'USA (NEC)'],
  ['Canada', 'CSA C22.1 – Canadian Electrical Code', 347, 600, 60, 'Solidly grounded wye', 3, 5, 'IES / NECB', 'NBC / CSA C22.2 No. 141', 'CSA C22.2 No. 244 / No. 29', 'CSA C22.2 No. 5', 'CEC Section 8', 'Local utility', 'North America (NEC Art. 220 / CEC Sec. 8)', 'Canada (CEC)'],
];

const FAMX = { 'United States': 'NEC', 'Canada': 'CEC', 'Australia': 'ASNZS' };
REG.forEach(r => r.push(FAMX[r[0]] || 'IEC'));
const SETN = { IEC: 'UK / IEC (BS 7671 – IET OSG)', ME: 'Middle East (DEWA / ADDC / SBC / KAHRAMAA)', NA: 'North America (NEC Art. 220 / CEC Sec. 8)', AU: 'Australia / NZ (AS/NZS 3000 App. C)', AF: 'Africa (SANS 10142 / NRS 034 / BS 7671)', AS: 'South & SE Asia (NBC India / SS 638 / MS IEC)' };
const TPL = {
  EU: ['TN-S / TT', 3, 5, 'EN 12464-1', 'EN 1838 / EN 50172', 'EN 61439', 'EN 60898-1 / EN 60947-2', 'IEC 60364 / national demand-factor rules', SETN.IEC, 'Europe (IEC)', 'IEC'],
  BS: ['TN-S / TT', 3, 5, 'BS EN 12464-1', 'BS 5266-1', 'IEC 61439', 'IEC 60898-1 / IEC 60947-2', 'BS 7671 / IET method', SETN.AF, 'Africa', 'IEC'],
  GCC: ['TN-S', 4, 4, 'EN 12464-1 / CIBSE', 'Civil defence / fire code', 'IEC 61439', 'IEC 60898-1 / IEC 60947-2', 'Authority connected-load & demand factors', SETN.ME, 'Middle East (GCC)', 'IEC'],
  ASIA: ['TN-S / TT', 3, 5, 'EN 12464-1 / national code', 'National fire code', 'IEC 61439', 'IEC 60898-1 / IEC 60947-2', 'Connected load × demand factors', SETN.AS, 'South & SE Asia', 'IEC'],
  EASIA: ['TN-S / TT', 3, 5, 'National lighting standard', 'National fire code', 'IEC 61439 (national adoption)', 'IEC 60898-1 / IEC 60947-2', 'Demand-factor method (national)', SETN.AS, 'East Asia', 'IEC'],
  NECA: ['Solidly grounded wye', 3, 5, 'IES / national', 'NFPA 101 / national', 'UL 891 / IEC 61439', 'UL 489 / IEC 60947-2', 'NEC-based Article 220 method (national adoption)', SETN.NA, 'Latin America & Caribbean', 'NEC'],
  NEC: ['Solidly grounded wye', 3, 5, 'IES / ASHRAE 90.1 LPD', 'NFPA 101 / NEC Art. 700', 'UL 891 / UL 67', 'UL 489', 'NEC Article 220 (VA/ft²)', SETN.NA, 'USA (NEC)', 'NEC'],
  LATIEC: ['TN-S / TT', 4, 4, 'ISO/CIE 8995-1 / EN 12464-1', 'National fire code', 'IEC 61439', 'IEC 60898-1 / IEC 60947-2', 'IEC 60364-based national demand factors', SETN.IEC, 'Latin America & Caribbean', 'IEC'],
  ASNZS: ['MEN (TN-C-S)', 5, 5, 'AS/NZS 1680', 'AS/NZS 2293', 'AS/NZS 61439', 'AS/NZS 60898', 'AS/NZS 3000 Appendix C', SETN.AU, 'Australia / NZ', 'ASNZS'],
};
const EXTRA = [
  ['France', 'NF C 15-100', 230, 400, 50, 'Enedis / local DSO', 'EU'], ['Germany', 'DIN VDE 0100 (HD 60364); VDE-AR-N 4100', 230, 400, 50, 'Local Netzbetreiber (DSO)', 'EU'],
  ['Spain', 'REBT (RD 842/2002) + ITC-BT', 230, 400, 50, 'i-DE / e-distribución / local DSO', 'EU'], ['Italy', 'CEI 64-8', 230, 400, 50, 'e-distribuzione / local DSO', 'EU'],
  ['Netherlands', 'NEN 1010', 230, 400, 50, 'Netbeheerder (Liander, Stedin, Enexis)', 'EU'], ['Belgium', 'AREI / RGIE', 230, 400, 50, 'Fluvius / ORES / Sibelga', 'EU'],
  ['Portugal', 'RTIEBT', 230, 400, 50, 'E-REDES', 'EU'], ['Switzerland', 'NIN SN 411000', 230, 400, 50, 'Local utility (VSE rules)', 'EU'],
  ['Austria', 'ÖVE/ÖNORM E 8101', 230, 400, 50, 'Local Netzbetreiber', 'EU'], ['Sweden', 'SS 436 40 00', 230, 400, 50, 'Local DSO (elnätsföretag)', 'EU'],
  ['Norway', 'NEK 400', 230, 400, 50, 'Local DSO (230 V IT networks common)', 'EU', { 0: 'IT / TN' }], ['Denmark', 'DS/HD 60364 (Elinstallationsbekendtgørelsen)', 230, 400, 50, 'Local DSO', 'EU'],
  ['Finland', 'SFS 6000', 230, 400, 50, 'Local DSO', 'EU'], ['Poland', 'PN-HD 60364', 230, 400, 50, 'OSD (local DSO)', 'EU'],
  ['Czech Republic', 'ČSN 33 2000 (HD 60364)', 230, 400, 50, 'ČEZ Distribuce / EG.D / PREdistribuce', 'EU'], ['Greece', 'ELOT HD 384', 230, 400, 50, 'HEDNO', 'EU'],
  ['Romania', 'I7-2011 / SR HD 60364', 230, 400, 50, 'Distribution operator', 'EU'], ['Turkey', 'Elektrik İç Tesisleri Yönetmeliği / TS HD 60364', 230, 400, 50, 'Distribution company (EDAŞ)', 'EU'],
  ['Russia', 'PUE (7th ed.) / GOST R 50571 (IEC 60364)', 230, 400, 50, 'Regional grid company (Rosseti)', 'EU', { 3: 'SP 52.13330', 4: 'SP 52.13330', 5: 'GOST IEC 61439', 6: 'GOST IEC 60898-1' }],
  ['Kuwait', 'MEW Regulations for Electrical Installations', 240, 415, 50, 'Ministry of Electricity & Water (MEW)', 'GCC'], ['Oman', 'Nama / OES Distribution Regulations', 240, 415, 50, 'Nama Electricity Distribution Co.', 'GCC'],
  ['Bahrain', 'EWA Regulations for Electrical Installations', 230, 400, 50, 'Electricity & Water Authority (EWA)', 'GCC'], ['Jordan', 'Jordanian Electrical Code (IEC 60364-based)', 230, 400, 50, 'JEPCO / EDCO / IDECO', 'GCC'],
  ['Egypt', 'Egyptian Code for Electrical Installations', 220, 380, 50, 'Distribution company (EEHC)', 'GCC'], ['Israel', 'Electricity Law regulations / SI 60364', 230, 400, 50, 'Israel Electric Corporation', 'GCC'],
  ['Iraq', 'Iraqi Electrical Code / IEC 60364', 230, 400, 50, 'Ministry of Electricity', 'GCC'], ['Lebanon', 'IEC 60364 / NF C 15-100 practice', 230, 400, 50, 'Électricité du Liban', 'GCC'],
  ['Tanzania', 'TBS / BS 7671 practice', 230, 400, 50, 'TANESCO', 'BS'], ['Uganda', 'UNBS / BS 7671 practice', 240, 415, 50, 'UEDCL / Umeme', 'BS'],
  ['Rwanda', 'RSB / IEC 60364', 230, 400, 50, 'REG / EUCL', 'BS'], ['Ethiopia', 'EBCS / IEC 60364', 220, 380, 50, 'Ethiopian Electric Utility', 'BS'],
  ['Zambia', 'ZS / BS 7671 practice', 230, 400, 50, 'ZESCO', 'BS'], ['Zimbabwe', 'SAZ / BS 7671 practice', 220, 380, 50, 'ZETDC', 'BS'],
  ['Botswana', 'BOBS / SANS 10142-1 practice', 230, 400, 50, 'Botswana Power Corporation', 'BS'], ['Namibia', 'SANS 10142-1 (adopted)', 230, 400, 50, 'NamPower / REDs', 'BS'],
  ['Malawi', 'MBS / BS 7671 practice', 230, 400, 50, 'ESCOM', 'BS'], ['Cameroon', 'NF C 15-100 practice', 220, 380, 50, 'ENEO', 'BS'],
  ['Senegal', 'NF C 15-100 practice', 230, 400, 50, 'Senelec', 'BS'], ['Cote d\'Ivoire', 'NF C 15-100 practice', 220, 380, 50, 'CIE', 'BS'],
  ['Morocco', 'NF C 15-100 / NM 06.1.100', 220, 380, 50, 'ONEE / Lydec / Redal', 'BS'], ['Algeria', 'NF C 15-100 practice', 230, 400, 50, 'Sonelgaz', 'BS'],
  ['Tunisia', 'NF C 15-100 practice', 230, 400, 50, 'STEG', 'BS'], ['Sierra Leone', 'BS 7671 practice', 230, 400, 50, 'EDSA', 'BS'],
  ['Mauritius', 'BS 7671 practice', 230, 400, 50, 'Central Electricity Board', 'BS'],
  ['Pakistan', 'Pakistan Electric & Telecom Safety Code / IEC 60364', 230, 400, 50, 'DISCO (LESCO, K-Electric …)', 'ASIA'], ['Bangladesh', 'BNBC 2020 Part 8', 220, 380, 50, 'DESCO / DPDC / BPDB', 'ASIA'],
  ['Sri Lanka', 'SLS 1522 / IEE practice', 230, 400, 50, 'CEB / LECO', 'ASIA'], ['Nepal', 'NBC 207 / IEC 60364', 230, 400, 50, 'Nepal Electricity Authority', 'ASIA'],
  ['Indonesia', 'PUIL 2011 (SNI 0225)', 230, 400, 50, 'PLN', 'ASIA'], ['Thailand', 'EIT Electrical Installation Standard', 230, 400, 50, 'MEA / PEA', 'ASIA'],
  ['Vietnam', 'TCVN 7447 (IEC 60364) / QCVN', 220, 380, 50, 'EVN', 'ASIA'], ['Brunei', 'IEE / BS 7671 practice', 240, 415, 50, 'DES / BPC', 'ASIA'],
  ['Maldives', 'BS 7671 practice', 230, 400, 50, 'STELCO / FENAKA', 'ASIA'],
  ['Philippines', 'Philippine Electrical Code (PEC, NEC-based)', 230, 400, 60, 'Meralco / electric cooperative', 'NECA', { 9: 'South & SE Asia' }],
  ['China', 'GB 50054 / GB/T 16895 (IEC 60364)', 220, 380, 50, 'State Grid / China Southern Power Grid', 'EASIA', { 3: 'GB 50034', 4: 'GB 51309', 5: 'GB/T 7251 (IEC 61439)', 6: 'GB/T 10963 (IEC 60898)', 7: 'GB 50052 / JGJ 16 demand factors' }],
  ['Japan', 'Technical Standards for Electrical Equipment; JIS C 60364', 100, 200, 50, 'Regional EPCO (50 Hz east / 60 Hz west)', 'EASIA', { 0: 'TT', 3: 'JIS Z 9110', 4: 'Building Standard Law', 5: 'JIS C 8480', 6: 'JIS C 8211', 7: 'JEAC 8001 demand factors' }],
  ['South Korea', 'Korea Electro-technical Code (KEC)', 220, 380, 60, 'KEPCO', 'EASIA', { 3: 'KS A 3011', 7: 'KEC demand factors' }],
  ['Taiwan', 'Electrical Installation Rules for Customers', 110, 220, 60, 'Taipower', 'EASIA', { 3: 'CNS 12112' }],
  ['Mongolia', 'IEC 60364 / PUE practice', 230, 400, 50, 'Ulaanbaatar Electricity Distribution', 'EASIA'],
  ['Mexico', 'NOM-001-SEDE-2012 (NEC-based)', 127, 220, 60, 'CFE', 'NECA'], ['Colombia', 'RETIE / NTC 2050 (NEC-based)', 120, 208, 60, 'Distribution operator (Enel, EPM …)', 'NECA'],
  ['Ecuador', 'NEC-SB-IE / NEC practice', 120, 208, 60, 'CNEL / EEQ', 'NECA'], ['Costa Rica', 'Código Eléctrico (NEC adopted)', 120, 240, 60, 'ICE / CNFL', 'NECA'],
  ['Panama', 'NEC (adopted)', 120, 240, 60, 'ENSA / Naturgy', 'NECA'], ['Dominican Republic', 'NEC practice / national regulations', 120, 240, 60, 'EDENORTE / EDESUR / EDEESTE', 'NECA'],
  ['Jamaica', 'NEC (adopted, JS 1)', 110, 220, 50, 'JPS', 'NECA'], ['Trinidad and Tobago', 'TTBS / NEC practice', 115, 230, 60, 'T&TEC', 'NECA'],
  ['Bahamas', 'NEC (adopted)', 120, 240, 60, 'BPL', 'NECA'], ['Puerto Rico', 'NFPA 70 – NEC', 120, 240, 60, 'LUMA Energy', 'NEC'],
  ['Brazil', 'ABNT NBR 5410', 127, 220, 60, 'Local distributor (Enel, CPFL, Light …)', 'LATIEC', { 3: 'ABNT NBR ISO/CIE 8995-1', 4: 'ABNT NBR 10898', 5: 'ABNT NBR IEC 61439', 6: 'ABNT NBR NM 60898' }],
  ['Argentina', 'AEA 90364 (IEC 60364)', 220, 380, 50, 'EDENOR / EDESUR / local distributor', 'LATIEC'], ['Chile', 'NCh Elec. 4/2003 / RIC (SEC)', 220, 380, 50, 'Enel / CGE', 'LATIEC'],
  ['Peru', 'Código Nacional de Electricidad – Utilización', 220, 380, 60, 'Luz del Sur / Enel / local', 'LATIEC'], ['Uruguay', 'UTE Reglamento de Baja Tensión', 230, 400, 50, 'UTE', 'LATIEC'],
  ['New Zealand', 'AS/NZS 3000:2018 (Electricity (Safety) Regulations 2010)', 230, 400, 50, 'Local EDB (Vector, Orion …)', 'ASNZS'], ['Fiji', 'AS/NZS 3000 (adopted)', 240, 415, 50, 'Energy Fiji Limited', 'ASNZS'],
  ['Papua New Guinea', 'AS/NZS 3000 (adopted)', 240, 415, 50, 'PNG Power', 'ASNZS'],
];
EXTRA.forEach(([n, regs, vln, vll, hz, auth, t, ov]) => {
  const tp = TPL[t].slice(); if (ov) Object.entries(ov).forEach(([k, v]) => tp[+k] = v);
  REG.push([n, regs, vln, vll, hz, tp[0], tp[1], tp[2], tp[3], tp[4], tp[5], tp[6], tp[7], auth, tp[8], tp[9], tp[10]]);
});
REG.sort((a, b) => a[0].localeCompare(b[0]));
const RG0 = 6, RG1 = RG0 + REG.length - 1;
REG.forEach((row, i) => {
  const r = RG0 + i;
  row.forEach((v, j) => {
    const c = rc.getCell(r, j + 1);
    if ([2, 3, 4, 6, 7].includes(j)) { input(c, v, j >= 6 ? '0.0' : '0'); c.alignment = { horizontal: 'center' }; }
    else { c.value = v; c.border = box; c.font = { name: 'Arial', size: j === 0 ? 10 : 9, bold: j === 0 }; c.alignment = { wrapText: true, vertical: 'top' }; }
  });
  rc.getRow(r).height = 30;
});
rc.getCell(`A${RG1 + 2}`).value = 'Notes: (0) Country data are typical LV utilisation voltages and principal wiring codes – confirm with the local utility / authority. (1) North American defaults are commercial 277/480 V (USA) and 347/600 V (Canada); dwellings use 120/240 V – override on Project Info. (1a) Earlier note: regions often use 277/480 V for commercial buildings – enter an override on Project Info. (2) Cable current ratings in this workbook are BS 7671 / IEC 60364-5-52 values (mm²); for NEC/CEC regions use NEC Table 310.16 / CEC Table 2 (AWG/kcmil) and AS/NZS 3008.1 in Australia/NZ. (3) Voltage-drop limits are from origin of installation to load; confirm with the local code and utility.';
rc.getCell(`A${RG1 + 2}`).font = { name: 'Arial', size: 9, italic: true };
rc.getCell(`A${RG1 + 2}`).alignment = { wrapText: true };
rc.mergeCells(`A${RG1 + 2}:Q${RG1 + 4}`);
rc.getColumn(1).width = 24; rc.getColumn(2).width = 36; [3, 4, 5].forEach(i => rc.getColumn(i).width = 10); rc.getColumn(6).width = 18;
[7, 8].forEach(i => rc.getColumn(i).width = 11); [9, 10, 11, 12, 13, 14, 15, 16].forEach(i => rc.getColumn(i).width = 26); rc.getColumn(17).width = 12;
rc.views = [{ state: 'frozen', xSplit: 1, ySplit: 5 }];
const RCQ = "'Regional Codes'";
const REG_LIST = `${RCQ}!$A$${RG0}:$A$${RG1}`;
const RL = (col) => `INDEX(${RCQ}!$${col}$${RG0}:$${col}$${RG1},MATCH('Project Info'!$B$22,${REG_LIST},0))`;

// ---------------- 1. Project Info ----------------
const pi = wb.addWorksheet('Project Info');
title(pi, 'Electrical Load Assessment – Design Report', 'Early-stage maximum demand, equipment and cable sizing. Select the region first; replace estimates with a detailed load schedule at detailed design.');
legend(pi, 3);
const piRows = [
  [5, 'PROJECT DETAILS', null, null, true],
  [6, 'Project name', 'Example: Riverside Mixed-Use Development'],
  [7, 'Project reference', 'EX-0001'],
  [8, 'Client', 'Example Client Ltd'],
  [9, 'Site address / postcode', 'Example Street, London'],
  [10, 'Utility / DNO', 'Enter the local distribution company'],
  [11, 'Report revision', 'P01'],
  [12, 'Date', new Date(2026, 9, 18)],
  [13, 'Prepared by', ''],
  [14, 'Checked by', ''],
  [16, 'DESIGN PARAMETERS', null, null, true],
  [19, 'Power factor (overall)', 0.95, 'Typical 0.9–0.95 for modern LED/VSD-dominated loads'],
  [20, 'Spare / future capacity allowance', 0.2, 'Typical 15–25%; higher where heat/EV electrification is expected'],
  [21, 'Density basis for lookup', 'Typical', 'Choose Min / Typical / Max from the VA-m2 Data sheet'],
  [22, 'Region / design code', 'United Kingdom', 'Sets voltages, frequency, VD limits and standards (Regional Codes sheet)'],
];
piRows.forEach(([r, label, val, note, sec]) => {
  const a = pi.getCell(`A${r}`); a.value = label; a.font = { name: 'Arial', size: 10, bold: !!sec };
  if (sec) { ['A', 'B', 'C', 'D'].forEach(col => pi.getCell(`${col}${r}`).fill = secFill); return; }
  a.border = box;
  input(pi.getCell(`B${r}`), val);
  if (note) { pi.getCell(`C${r}`).value = note; pi.getCell(`C${r}`).font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } }; }
});
pi.getCell('D16').value = 'Override (optional)'; pi.getCell('D16').font = { name: 'Arial', size: 10, bold: true };
// Region-derived rows with optional override in column D
const derived = [
  [17, 'Supply voltage, line–line (V)', 'D', '0', true],
  [18, 'Supply voltage, line–neutral (V)', 'C', '0', true],
  [23, 'Supply frequency (Hz)', 'E', '0', true],
  [24, 'Wiring regulations', 'B', null, false],
  [25, 'Typical earthing system', 'F', null, true],
  [26, 'Voltage drop limit – lighting (%)', 'G', '0.0', true],
  [27, 'Voltage drop limit – other uses (%)', 'H', '0.0', true],
  [28, 'Lighting design standard', 'I', null, false],
  [29, 'Emergency lighting standard', 'J', null, false],
  [30, 'Switchgear / assemblies standard', 'K', null, false],
  [31, 'Protective device standard', 'L', null, false],
  [32, 'Load calculation method', 'M', null, false],
  [33, 'Supply authority', 'N', null, false],
  [34, 'Demand / diversity factor set', 'O', null, true],
  [35, 'Load density region (Load Schedule / Code Load Calc)', 'P', null, true],
  [36, 'Code family (compliance references, code calculation)', 'Q', null, true],
];
derived.forEach(([r, label, col, fmt, ovr]) => {
  const a = pi.getCell(`A${r}`); a.value = label; a.border = box; a.font = FONT;
  const f = ovr ? `IF(D${r}<>"",D${r},${RL(col)})` : RL(col);
  calc(pi.getCell(`B${r}`), f, fmt || undefined, GREEN);
  pi.getCell(`B${r}`).alignment = { wrapText: true };
  if (ovr) input(pi.getCell(`D${r}`), null, fmt || undefined);
});
pi.getCell('C17').value = 'From region; override e.g. 480 V for US commercial';
pi.getCell('C26').value = 'From origin of installation to load';
['C17', 'C26'].forEach(a => pi.getCell(a).font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } });
pi.getCell('B12').numFmt = 'dd/mm/yyyy';
pi.getCell('B19').numFmt = '0.00';
pi.getCell('B20').numFmt = '0%';
pi.getCell('B21').dataValidation = { type: 'list', allowBlank: false, formulae: ['"Min,Typical,Max"'] };
pi.getCell('B22').dataValidation = { type: 'list', allowBlank: false, formulae: [REG_LIST] };
pi.getCell('B19').dataValidation = { type: 'decimal', operator: 'between', formulae: [0.5, 1], showErrorMessage: true, error: 'Power factor must be between 0.5 and 1.0' };
pi.getColumn(1).width = 36; pi.getColumn(2).width = 46; pi.getColumn(3).width = 58; pi.getColumn(4).width = 20;
pi.getCell('D34').dataValidation = { type: 'list', allowBlank: true, formulae: ["'Diversity Factors'!$W$8:$W$13"] };
pi.getCell('C34').value = 'From region; override to apply another code\'s factors'; pi.getCell('C34').font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } };
pi.getCell('D35').dataValidation = { type: 'list', allowBlank: true, formulae: ["'Regional Densities'!$A$6:$A$15"] };
pi.getCell('D36').dataValidation = { type: 'list', allowBlank: true, formulae: ['"IEC,NEC,CEC,ASNZS"'] };
pi.getCell('C36').value = 'IEC = BS 7671 / IEC 60364 family; NEC = NFPA 70 family; CEC = CSA C22.1; ASNZS = AS/NZS 3000'; pi.getCell('C36').font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } };
pi.getCell('C35').value = 'From region; override to design to another region\'s practice'; pi.getCell('C35').font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF595959' } };
// Design basis and revision log
['A37', 'B37', 'C37', 'D37'].forEach(a => pi.getCell(a).fill = secFill);
pi.getCell('A37').value = 'DESIGN BASIS'; pi.getCell('A37').font = { name: 'Arial', size: 10, bold: true };
[[38, 'Design stage', 'Concept / RIBA Stage 2', '"Feasibility,Concept / RIBA Stage 2,Developed / RIBA Stage 3,Technical / RIBA Stage 4,Construction issue,As-built"'],
 [39, 'Building classification / main use', 'Mixed-use', null], [40, 'Gross internal area (m²)', null, null], [41, 'Number of storeys (above / below ground)', '8 / 1', null],
 [42, 'Ambient design temperature (°C)', 30, null], [43, 'Altitude (m) / special environment', '< 1,000 m, normal', null],
 [44, 'Client brief / employer\'s requirements reference', '', null], [45, 'Designer (competent person) and qualification', '', null]].forEach(([r, l, v, list]) => {
  const a = pi.getCell(`A${r}`); a.value = l; a.border = box; a.font = FONT;
  input(pi.getCell(`B${r}`), v);
  if (list) pi.getCell(`B${r}`).dataValidation = { type: 'list', allowBlank: true, formulae: [list] };
});
['A47', 'B47', 'C47', 'D47'].forEach(a => pi.getCell(a).fill = secFill);
pi.getCell('A47').value = 'REVISION HISTORY'; pi.getCell('A47').font = { name: 'Arial', size: 10, bold: true };
['Rev', 'Date', 'Description', 'By / checked'].forEach((h, i) => { const x = pi.getCell(48, 1 + i); x.value = h; x.fill = hdrFill; x.border = box; x.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } }; });
for (let r = 49; r <= 56; r++) { for (let c = 1; c <= 4; c++) input(pi.getCell(r, c), r === 49 ? ['P01', new Date(2026, 9, 21), 'First issue for comment', ''][c - 1] : null); pi.getCell(r, 2).numFmt = 'dd/mm/yyyy'; }
const VDL = `'Project Info'!$B$26`, VDO = `'Project Info'!$B$27`;

