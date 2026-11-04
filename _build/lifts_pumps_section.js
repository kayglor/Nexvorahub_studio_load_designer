// ---------------- 18b. Pumps & Lifts ----------------
const pl = wb.addWorksheet('Pumps & Lifts');
title(pl, 'Pump and Lift (Elevator) Sizing – Motor Rating, Starter / Control Type and kVA', 'Pumps: P = ρ·g·Q·H ÷ η; starter and control selected by service and motor size. Lifts: rated load from persons (EN 81-20, 75 kg/person); motor kW from load, speed and counterweight; handling capacity from persons per trip.');
legend(pl, 3);
const IEC = [0.37, 0.55, 0.75, 1.1, 1.5, 2.2, 3, 4, 5.5, 7.5, 11, 15, 18.5, 22, 30, 37, 45, 55, 75, 90, 110, 132, 160, 200, 250, 315];
IEC.forEach((v, i) => input(pl.getCell(`AC${6 + i}`), v, '0.00'));
pl.getCell('AC5').value = 'IEC motor kW'; pl.getCell('AC5').font = { name: 'Arial', size: 9, bold: true };
const IECR = `$AC$6:$AC$${5 + IEC.length}`;
const LOADS = [320, 450, 630, 800, 1000, 1275, 1600, 2000, 2500, 3000, 4000];
LOADS.forEach((v, i) => input(pl.getCell(`AD${6 + i}`), v, '0'));
pl.getCell('AD5').value = 'Lift rated loads (kg)'; pl.getCell('AD5').font = { name: 'Arial', size: 9, bold: true };
const LOADR = `$AD$6:$AD$${5 + LOADS.length}`;
const SVC = ['Domestic water booster', 'Chilled water', 'Condenser water', 'LTHW heating', 'Sewage / sump', 'Fire – sprinkler', 'Fire – wet riser / hydrant', 'Irrigation', 'Other'];
SVC.forEach((v, i) => { const c = pl.getCell(`AE${6 + i}`); c.value = v; c.font = FONT; });
pl.getCell('AE5').value = 'Pump services'; pl.getCell('AE5').font = { name: 'Arial', size: 9, bold: true };
const SVCR = `$AE$6:$AE$${5 + SVC.length}`;
const CTRL = [
  'Packaged duty/standby set; pressure transducer; auto changeover and run-hour rotation; low-water cut-out; BMS status/fault',
  'BMS ΔP / temperature control; duty/standby auto changeover; VSD 0–10 V speed signal',
  'BMS temperature control; duty/standby auto changeover; VSD speed signal; flow switch',
  'BMS ΔP / weather-compensated control; duty/standby; VSD speed signal',
  'Float / ultrasonic level control; duty/assist/standby; high-level alarm to BMS and local beacon',
  'Auto start on pressure drop; manual stop only; no overload trip; run/fail alarms to fire panel; fire-rated cable (BS 8519)',
  'Auto start on pressure drop / hydrant demand; manual stop only; alarms to fire panel; fire-rated cable (BS 8519)',
  'Timer / soil-moisture controller; local isolator; BMS status',
  'Local control panel; BMS run/fault status',
];

sub(pl, 'A5', 'A. PUMP SIZING AND STARTER / CONTROL SELECTION');
pl.getCell('A6').value = 'Service margin on shaft power'; pl.getCell('A6').font = FONT; input(pl.getCell('C6'), 1.15, '0.00');
pl.getCell('D6').value = 'Typical 1.10–1.20'; pl.getCell('D6').font = { name: 'Arial', size: 9, italic: true };
const PH = ['Ref', 'Service', 'Flow (l/s)', 'Head (m)', 'Pump efficiency', 'Duty qty', 'Standby qty', 'Hydraulic power (kW)', 'Shaft power (kW)', 'Required motor (kW)', 'Selected motor (kW)', 'Motor efficiency (IE3)', 'Starter / drive (auto)', 'Starter override', 'Applied starter / drive', 'Motor PF', 'Full-load current (A)', 'Starting current (× FLC)', 'Starting kVA (each)', 'Control system specification', 'Duty demand (kVA)', 'Notes'];
header(pl, 8, PH); pl.getRow(8).height = 45;
const PK = {}; PH.forEach((h, i) => PK[h] = colL(i));
const P0 = 9, P1 = 23;
const pEx = [
  ['P01', 'Domestic water booster', 3, 45, 0.65, 1, 1],
  ['P02', 'Chilled water', 25, 25, 0.75, 2, 1],
  ['P03', 'LTHW heating', 10, 15, 0.7, 1, 1],
  ['P04', 'Sewage / sump', 5, 12, 0.5, 1, 1],
  ['P05', 'Fire – sprinkler', 30, 70, 0.7, 1, 1],
];
for (let r = P0; r <= P1; r++) {
  const e = pEx[r - P0] || [];
  const c = (k) => pl.getCell(`${PK[k]}${r}`), R = (k) => `${PK[k]}${r}`;
  input(c('Ref'), e[0] || null); input(c('Service'), e[1] || null);
  c('Service').dataValidation = { type: 'list', allowBlank: true, formulae: [`'Pumps & Lifts'!${SVCR}`] };
  input(c('Flow (l/s)'), e[2] !== undefined ? e[2] : null, '#,##0.0'); input(c('Head (m)'), e[3] !== undefined ? e[3] : null, '#,##0.0');
  input(c('Pump efficiency'), e[4] !== undefined ? e[4] : null, '0.00'); input(c('Duty qty'), e[5] !== undefined ? e[5] : null, '0'); input(c('Standby qty'), e[6] !== undefined ? e[6] : null, '0');
  calc(c('Hydraulic power (kW)'), `IF(OR(${R('Flow (l/s)')}="",${R('Head (m)')}=""),0,9.81*${R('Flow (l/s)')}*${R('Head (m)')}/1000)`, '#,##0.00;-#,##0.00;"-"');
  calc(c('Shaft power (kW)'), `IF(N(${R('Pump efficiency')})=0,0,${R('Hydraulic power (kW)')}/${R('Pump efficiency')})`, '#,##0.00;-#,##0.00;"-"');
  calc(c('Required motor (kW)'), `${R('Shaft power (kW)')}*$C$6`, '#,##0.00;-#,##0.00;"-"');
  calc(c('Selected motor (kW)'), `IF(${R('Required motor (kW)')}=0,"",IF(${R('Required motor (kW)')}>MAX(${IECR}),"> "&MAX(${IECR}),INDEX(${IECR},COUNTIF(${IECR},"<"&${R('Required motor (kW)')})+1)))`, '0.00');
  const kw = `N(${R('Selected motor (kW)')})`;
  calc(c('Motor efficiency (IE3)'), `IF(${kw}=0,"",IF(${kw}<=1.1,0.82,IF(${kw}<=4,0.87,IF(${kw}<=11,0.9,IF(${kw}<=37,0.93,IF(${kw}<=110,0.95,0.96))))))`, '0.00');
  const sv = R('Service');
  calc(c('Starter / drive (auto)'), `IF(OR(${sv}="",${kw}=0),"",IF(LEFT(${sv},4)="Fire",IF(${kw}<=30,"DOL – fire pump controller (BS EN 12845)","Star-delta – fire pump controller (BS EN 12845)"),IF(OR(${sv}="Domestic water booster",${sv}="Chilled water",${sv}="Condenser water",${sv}="LTHW heating"),"VSD (variable speed drive)",IF(${sv}="Sewage / sump",IF(${kw}<=7.5,"DOL","Soft starter"),IF(${kw}<=7.5,"DOL",IF(${kw}<=30,"Star-delta","Soft starter / VSD"))))))`);
  input(c('Starter override'), null);
  c('Starter override').dataValidation = { type: 'list', allowBlank: true, formulae: ['"DOL,Star-delta,Soft starter,VSD (variable speed drive)"'] };
  calc(c('Applied starter / drive'), `IF(${R('Starter override')}<>"",${R('Starter override')},${R('Starter / drive (auto)')})`);
  const st = R('Applied starter / drive');
  calc(c('Motor PF'), `IF(${st}="","",IF(LEFT(${st},3)="VSD",0.95,0.85))`, '0.00');
  calc(c('Full-load current (A)'), `IF(${kw}=0,"",${kw}*1000/(SQRT(3)*${VLL}*${R('Motor PF')}*${R('Motor efficiency (IE3)')}))`, '#,##0.0');
  calc(c('Starting current (× FLC)'), `IF(${st}="","",IF(LEFT(${st},3)="DOL",6.5,IF(LEFT(${st},4)="Star",2.5,IF(LEFT(${st},4)="Soft",3,IF(LEFT(${st},3)="VSD",1.1,6.5)))))`, '0.0');
  calc(c('Starting kVA (each)'), `IF(${kw}=0,"",SQRT(3)*${VLL}*${R('Full-load current (A)')}*${R('Starting current (× FLC)')}/1000)`, '#,##0.0');
  calc(c('Control system specification'), `IF(${sv}="","",CHOOSE(MATCH(${sv},${SVCR},0),"${CTRL.join('","')}"))`);
  calc(c('Duty demand (kVA)'), `IF(${kw}=0,0,N(${R('Duty qty')})*${kw}/(${R('Motor efficiency (IE3)')}*${R('Motor PF')}))`, '#,##0.0;-#,##0.0;"-"');
  input(c('Notes'), null);
}
const PT = P1 + 1;
pl.getCell(`A${PT}`).value = 'Duty demand – normal services (kVA)'; pl.mergeCells(`A${PT}:D${PT}`);
calc(pl.getCell(`${PK['Duty demand (kVA)']}${PT}`), `SUMPRODUCT((LEFT(B${P0}:B${P1},4)<>"Fire")*${PK['Duty demand (kVA)']}${P0}:${PK['Duty demand (kVA)']}${P1})`, '#,##0.0');
pl.getCell(`A${PT + 1}`).value = 'Duty demand – fire / life-safety pumps (kVA)'; pl.mergeCells(`A${PT + 1}:D${PT + 1}`);
calc(pl.getCell(`${PK['Duty demand (kVA)']}${PT + 1}`), `SUMPRODUCT((LEFT(B${P0}:B${P1},4)="Fire")*${PK['Duty demand (kVA)']}${P0}:${PK['Duty demand (kVA)']}${P1})`, '#,##0.0');
pl.getCell(`A${PT + 2}`).value = 'Largest starting kVA (generator / voltage-dip check)'; pl.mergeCells(`A${PT + 2}:D${PT + 2}`);
calc(pl.getCell(`${PK['Starting kVA (each)']}${PT + 2}`), `MAX(${PK['Starting kVA (each)']}${P0}:${PK['Starting kVA (each)']}${P1})`, '#,##0.0');
[PT, PT + 1, PT + 2].forEach(r => PH.forEach((_, i) => { const x = pl.getCell(r, i + 1); x.fill = totFill; x.border = box; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); }));

// ---- Lifts ----
const LS0 = PT + 5;
sub(pl, `A${LS0}`, 'B. LIFT (ELEVATOR) SIZING – CAPACITY, kVA AND HANDLING CAPACITY');
pl.getCell(`A${LS0 + 1}`).value = 'Target handling capacity (% population / 5 min)'; pl.getCell(`A${LS0 + 1}`).font = FONT; input(pl.getCell(`C${LS0 + 1}`), 0.12, '0%');
pl.getCell(`D${LS0 + 1}`).value = 'Offices 12–15%; hotels 10–12%; residential 5–8% (CIBSE Guide D)'; pl.getCell(`D${LS0 + 1}`).font = { name: 'Arial', size: 9, italic: true };
pl.getCell(`A${LS0 + 2}`).value = 'Target average interval (s)'; pl.getCell(`A${LS0 + 2}`).font = FONT; input(pl.getCell(`C${LS0 + 2}`), 30, '0');
pl.getCell(`D${LS0 + 2}`).value = 'Offices ≤ 30 s; residential ≤ 60 s'; pl.getCell(`D${LS0 + 2}`).font = { name: 'Arial', size: 9, italic: true };
const HC_T = `$C$${LS0 + 1}`, INT_T = `$C$${LS0 + 2}`;
const LH = ['Ref', 'Lift / group', 'Lift type', 'Lifts in group', 'Persons per car', 'Rated load (kg)', 'Floors served', 'Floor height (m)', 'Travel (m)', 'Speed override (m/s)', 'Design speed (m/s)', 'Counterweight balance', 'System efficiency', 'Motor power (kW)', 'Selected motor (kW)', 'Running kVA per lift', 'Peak (acceleration) kVA per lift', 'Group demand (kVA)', 'Population served', 'Passengers per trip (80%)', 'Round trip time (s)', 'Handling capacity (persons / 5 min)', 'HC % of population', 'Average interval (s)', 'Check', 'Notes'];
const LHr = LS0 + 4;
header(pl, LHr, LH); pl.getRow(LHr).height = 45;
const LK = {}; LH.forEach((h, i) => LK[h] = colL(i));
const LTYPES = ['Passenger – traction MRL (VVVF)', 'Passenger – hydraulic', 'Goods / service – traction', 'Fire-fighting lift (BS EN 81-72)', 'Bed / passenger – traction'];
LTYPES.forEach((v, i) => { const c = pl.getCell(`AF${6 + i}`); c.value = v; c.font = FONT; });
pl.getCell('AF5').value = 'Lift types'; pl.getCell('AF5').font = { name: 'Arial', size: 9, bold: true };
const LTR = `'Pumps & Lifts'!$AF$6:$AF$${5 + LTYPES.length}`;
const L0r = LHr + 1, L1r = L0r + 9;
const lEx = [
  ['L01', 'Main passenger group', 'Passenger – traction MRL (VVVF)', 3, 13, 9, 3.5, 600],
  ['L02', 'Goods / service lift', 'Goods / service – traction', 1, 21, 9, 3.5, 0],
  ['L03', 'Fire-fighting lift', 'Fire-fighting lift (BS EN 81-72)', 1, 13, 9, 3.5, 0],
  ['L04', 'Car park shuttle', 'Passenger – hydraulic', 1, 8, 2, 3.0, 50],
];
for (let r = L0r; r <= L1r; r++) {
  const e = lEx[r - L0r] || [];
  const c = (k) => pl.getCell(`${LK[k]}${r}`), R = (k) => `${LK[k]}${r}`;
  input(c('Ref'), e[0] || null); input(c('Lift / group'), e[1] || null); input(c('Lift type'), e[2] || null);
  c('Lift type').dataValidation = { type: 'list', allowBlank: true, formulae: [LTR] };
  input(c('Lifts in group'), e[3] !== undefined ? e[3] : null, '0'); input(c('Persons per car'), e[4] !== undefined ? e[4] : null, '0');
  input(c('Floors served'), e[5] !== undefined ? e[5] : null, '0'); input(c('Floor height (m)'), e[6] !== undefined ? e[6] : null, '0.0');
  input(c('Speed override (m/s)'), null, '0.00'); input(c('Population served'), e[7] !== undefined ? e[7] : null, '#,##0');
  const hyd = `${R('Lift type')}="Passenger – hydraulic"`;
  calc(c('Rated load (kg)'), `IF(N(${R('Persons per car')})=0,"",IF(${R('Persons per car')}*75>MAX(${LOADR}),${R('Persons per car')}*75,INDEX(${LOADR},COUNTIF(${LOADR},"<"&${R('Persons per car')}*75)+1)))`, '#,##0');
  calc(c('Travel (m)'), `MAX(N(${R('Floors served')})-1,0)*N(${R('Floor height (m)')})`, '#,##0.0');
  calc(c('Design speed (m/s)'), `IF(${R('Lift type')}="","",IF(${R('Speed override (m/s)')}<>"",${R('Speed override (m/s)')},IF(${hyd},0.63,IF(${R('Travel (m)')}<=15,1,IF(${R('Travel (m)')}<=35,1.6,IF(${R('Travel (m)')}<=75,2.5,4))))))`, '0.00');
  calc(c('Counterweight balance'), `IF(${R('Lift type')}="","",IF(${hyd},0,0.5))`, '0.00');
  calc(c('System efficiency'), `IF(${R('Lift type')}="","",IF(${hyd},0.5,0.75))`, '0.00');
  const ok = `AND(${R('Lift type')}<>"",ISNUMBER(${R('Rated load (kg)')}))`;
  calc(c('Motor power (kW)'), `IF(${ok},IF(${hyd},2*${R('Rated load (kg)')},(1-${R('Counterweight balance')})*${R('Rated load (kg)')})*9.81*${R('Design speed (m/s)')}/(${R('System efficiency')}*1000),0)`, '#,##0.0;-#,##0.0;"-"');
  calc(c('Selected motor (kW)'), `IF(${R('Motor power (kW)')}=0,"",IF(${R('Motor power (kW)')}>MAX(${IECR}),"> "&MAX(${IECR}),INDEX(${IECR},COUNTIF(${IECR},"<"&${R('Motor power (kW)')})+1)))`, '0.0');
  const mk = `N(${R('Selected motor (kW)')})`;
  calc(c('Running kVA per lift'), `IF(${mk}=0,0,${mk}/(0.9*IF(${hyd},0.85,0.95)))`, '#,##0.0;-#,##0.0;"-"');
  calc(c('Peak (acceleration) kVA per lift'), `${R('Running kVA per lift')}*IF(${hyd},3,1.8)`, '#,##0.0;-#,##0.0;"-"');
  calc(c('Group demand (kVA)'), `IF(N(${R('Lifts in group')})=0,0,${R('Running kVA per lift')}*(1+IF(${R('Lifts in group')}>=2,0.75,0)+0.6*MAX(${R('Lifts in group')}-2,0)))`, '#,##0.0;-#,##0.0;"-"');
  calc(c('Passengers per trip (80%)'), `0.8*N(${R('Persons per car')})`, '0.0');
  const N_ = `MAX(N(${R('Floors served')})-1,1)`, P_ = R('Passengers per trip (80%)');
  calc(c('Round trip time (s)'), `IF(OR(NOT(${ok}),N(${R('Design speed (m/s)')})=0),"",2*${R('Travel (m)')}/${R('Design speed (m/s)')}+(${N_}*(1-((${N_}-1)/${N_})^${P_})+1)*10+2*${P_}*1.2)`, '0');
  calc(c('Handling capacity (persons / 5 min)'), `IF(N(${R('Round trip time (s)')})=0,"",300*${P_}*N(${R('Lifts in group')})/${R('Round trip time (s)')})`, '0');
  calc(c('HC % of population'), `IF(OR(N(${R('Population served')})=0,${R('Handling capacity (persons / 5 min)')}=""),"",${R('Handling capacity (persons / 5 min)')}/${R('Population served')})`, '0.0%');
  calc(c('Average interval (s)'), `IF(OR(N(${R('Round trip time (s)')})=0,N(${R('Lifts in group')})=0),"",${R('Round trip time (s)')}/${R('Lifts in group')})`, '0');
  calc(c('Check'), `IF(${R('HC % of population')}="",IF(${ok},"Service / fire lift – capacity n/a",""),IF(AND(${R('HC % of population')}>=${HC_T},${R('Average interval (s)')}<=${INT_T}),"OK",IF(${R('HC % of population')}<${HC_T},"Increase lifts / car size","Interval long – add lifts / speed")))`);
  input(c('Notes'), null);
}
const LTt = L1r + 1;
pl.getCell(`A${LTt}`).value = 'TOTAL – lift group demand (kVA)'; pl.mergeCells(`A${LTt}:D${LTt}`);
calc(pl.getCell(`${LK['Group demand (kVA)']}${LTt}`), `SUM(${LK['Group demand (kVA)']}${L0r}:${LK['Group demand (kVA)']}${L1r})`, '#,##0.0');
calc(pl.getCell(`${LK['Peak (acceleration) kVA per lift']}${LTt}`), `MAX(${LK['Peak (acceleration) kVA per lift']}${L0r}:${LK['Peak (acceleration) kVA per lift']}${L1r})`, '#,##0.0');
LH.forEach((_, i) => { const x = pl.getCell(LTt, i + 1); x.fill = totFill; x.border = box; x.font = Object.assign({}, x.font || FONT, { name: 'Arial', bold: true }); });
pl.getCell(`A${LTt + 1}`).value = 'Method: rated load = persons × 75 kg rounded up to a standard EN 81-20 load; traction motor kW = (1 − balance) × load × g × v ÷ η; hydraulic kW = (car + load) × g × v ÷ η with car mass ≈ rated load. Group demand applies 100% largest + 75% second + 60% others. RTT = 2H/v + (stops + 1) × 10 s + 2P × 1.2 s (simplified CIBSE Guide D method). Confirm with the lift manufacturer.';
pl.getCell(`A${LTt + 1}`).font = { name: 'Arial', size: 9, italic: true }; pl.getCell(`A${LTt + 1}`).alignment = { wrapText: true };
pl.mergeCells(LTt + 1, 1, LTt + 3, 20);
for (let i = 1; i <= 26; i++) pl.getColumn(i).width = 11;
pl.getColumn('B').width = 24; pl.getColumn('C').width = 30; pl.getColumn(PK['Starter / drive (auto)']).width = 30; pl.getColumn(PK['Applied starter / drive']).width = 30; pl.getColumn(PK['Control system specification']).width = 60;
pl.getColumn('AC').width = 12; pl.getColumn('AD').width = 12; pl.getColumn('AE').width = 26; pl.getColumn('AF').width = 32;
pl.views = [{ state: 'frozen', xSplit: 2 }];
const PLQ = "'Pumps & Lifts'";

