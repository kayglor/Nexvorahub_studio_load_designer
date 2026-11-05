const out = {};
out.md = +R.md.toFixed(1); out.areaKva = +R.Z.kva.toFixed(1); out.zones = R.Z.rows.map(r => [r.z.type.slice(0, 24), r.dens, +r.kva.toFixed(1)]); out.equip = +R.Lo.kva.toFixed(1);
out.lifts = R.Li.rows.map(r => [r.load, r.v, +r.kw.toFixed(2), r.motor, +r.grp.toFixed(1), r.hcp && +r.hcp.toFixed(3), Math.round(r.interval)]);
out.pumps = R.P.rows.map(r => [r.motor, r.starter, +r.kva.toFixed(1), +r.startKva.toFixed(1)]);
out.circ = R.C.map(r => [r.k.ref, +r.ib.toFixed(1), r.inr, r.csa, r.Iz && +r.Iz.toFixed(1), r.vd && +r.vd.toFixed(2), r.zs && +r.zs.toFixed(3), r.cpc, r.status]);
out.tx = { n: R.nTx, kva: R.txKva, flc: R.flc && Math.round(R.flc), fl: R.flP && +R.flP.toFixed(1), inc: R.incomer, icw: R.icw, gen: R.gen, met: R.met, bond: R.bond, room: R.room && [R.room.L, R.room.W, +R.room.swA.toFixed(1)] };
out.dc = (() => { const d = R.Dfull; return [d.racks, d.it, Math.round(d.fac), Math.round(d.md), +d.pueP.toFixed(2), d.upsI, d.genI, d.txI, d.crah, d.chill, d.rackPdu]; })();
out.energy = [R.E.base, R.E.savKwh, R.E.imp, R.E.eui, R.E.bank, R.E.mdAfter].map(Math.round);
out.checks = CHECKS.map(c => c[2](R)[1]).join(',');
const errs = [];
for (const [id] of VIEWS) { try { view = id; render(); if (!document.querySelector('#main').innerHTML.length) errs.push(id + ': empty'); } catch (e) { errs.push(id + ': ' + e.stack.split('\n').slice(0, 2).join(' ')); } }
out.viewErrors = errs;
const sw = [];
for (const c of COUNTRIES.map(c => c.name)) {
  S.project.country = c;
  try {
    R = calcAll(); for (const [id] of VIEWS) { view = id; render(); } reportDoc(); csvSchedules();
    if (['United States', 'Canada', 'United Arab Emirates', 'Japan', 'Brazil', 'Australia', 'Germany', 'Nigeria', 'India', 'China'].includes(c)) sw.push([c, R.X.vln + '/' + R.X.vll, R.X.dreg, R.X.fam, Math.round(R.Z.kva), Math.round(R.md), R.C.map(x => x.status.replace('FAIL – ', 'F:')).join('|')]);
  } catch (e) { sw.push([c, 'ERR ' + e.message]); }
}
out.countriesTested = COUNTRIES.length; out.sweep = sw;
S.project.country = 'United Kingdom'; S.dc.include = true; R = calcAll(); out.withDC = Math.round(R.md); render();
for (const t of TIERS) { S.dc.tier = t[0]; R = calcAll(); out[t[0].slice(0, 8)] = [R.D.upsI, R.D.genI, R.D.txI, R.D.crah, R.D.paths]; }
out.reportLen = reportDoc().length; out.csvLines = csvSchedules().split('\n').length;

// ---- AI data centre power planning (ai-dc-data.json) ----
const aiFail = [], near = (a, b, tol = 1e-6) => Math.abs(a - b) <= tol * Math.max(1, Math.abs(b));
const chk = (name, ok, got) => { if (!ok) aiFail.push(name + (got !== undefined ? ' got ' + JSON.stringify(got) : '')); };
// library integrity: every real platform has typical kW, liquid share, architecture, source URL and as-of date
AI_PLAT.slice(1).forEach(p => chk('platform data ' + p.name, p.kw[0] <= p.kw[1] && p.kw[1] <= p.kw[2] && p.liquid >= 0 && p.liquid <= 1 && /^https:\/\//.test(p.url) && isFinite(Date.parse(p.asOf)) && ['ac', 'dc'].includes(p.arch) && p.podRacks > 0));
chk('custom platform first', AI_PLAT[0].name === AI_CUSTOM && AI_PLAT[0].kw[1] == null);
chk('library version/date', !!AIDC.version && isFinite(Date.parse(AIDC.asOf)) && AIDC.staleMonths === 6);
// stale flag: older than 6 months
const ref = new Date('2026-11-05T00:00:00Z');
chk('stale 10 months', aiStale('2026-01-01', ref) === true); chk('fresh 1 month', aiStale('2026-10-01', ref) === false); chk('no date not stale', aiStale('', ref) === false);
chk('age months', near(aiAgeMonths('2026-05-05', ref), 6, 0.02), aiAgeMonths('2026-05-05', ref));
// default custom project keeps the original model (no liquid split, no DC adjustment)
S.dc.tier = 'Tier III – Concurrently maintainable'; R = calcAll();
chk('custom no liquid', R.D.ai.liqHeat === 0 && R.D.ai.cdu === 0 && R.D.ai.convAdj === 0 && R.D.crah === RED('N+1', Math.ceil(R.D.heat / 150)));
// GB300 NVL72, 100 racks
const dcSave = JSON.parse(JSON.stringify(S.dc));
S.dc.platform = 'NVIDIA GB300 NVL72'; applyPlatform(S.dc.platform); S.dc.method = 'By rack count'; S.dc.racks = 100; R = calcAll();
let D = R.D, a = D.ai;
chk('GB300 fills density', S.dc.density === 138 && S.dc.liquid === 0.85 && S.dc.arch === AI_ARCH.ac.name, [S.dc.density, S.dc.liquid, S.dc.arch]);
chk('GB300 IT', D.it === 13800, D.it);
chk('liquid heat', near(a.liqHeat, 11730), a.liqHeat); chk('air heat', near(a.airHeat, D.heat - 11730), a.airHeat);
chk('CDU count N+1', a.cdu === 9, a.cdu);
chk('CRAH on residual air', D.crah === RED('N+1', Math.ceil(a.airHeat / 150)), D.crah);
chk('rack amps AC', near(a.rackAac, 138000 / (Math.sqrt(3) * 400 * 0.95)), a.rackAac); chk('rack amps DC', near(a.rackAdc, 172.5), a.rackAdc);
chk('pods', a.pods === 13 && a.podKw === 1104, [a.pods, a.podKw]);
chk('cooling split', near(D.cool, 11730 * COOL['Direct-to-chip liquid cooling'] + a.airHeat * COOL[S.dc.cooling]), D.cool);
const effAC = S.dc.upsEff * AI_ARCH.ac.dist * AI_ARCH.ac.psu, effDC = AI_ARCH.dc.rect * AI_ARCH.dc.dist * AI_ARCH.dc.psu, l54 = 13800 * AI_ARCH.ac.psu;
chk('conversion AC/DC', near(a.convAC, l54 / effAC - l54) && near(a.convDC, l54 / effDC - l54) && a.convSave > 0 && a.convAdj === 0, [a.convAC, a.convDC]);
chk('DC more efficient by 3–7%', effDC - effAC > 0.03 && effDC - effAC < 0.07, effDC - effAC);
const facAC = D.fac;
// switch to 800 V DC: facility load falls by the conversion-loss difference
S.dc.arch = AI_ARCH.dc.name; R = calcAll(); D = R.D; a = D.ai;
chk('DC facility reduction', near(facAC - D.fac, a.convSave * 1.015), facAC - D.fac);
chk('DC rack power text', /800 V DC/.test(D.rackPdu), D.rackPdu);
// training dynamics + grid-forming BESS
Object.assign(S.dc, { swing: 0.5, ramp: 2, bessRed: 0.5, bessMin: 15 }); R = calcAll(); a = R.D.ai;
chk('swing MW', near(a.swingMW, 6.9) && near(a.rampMWs, 3.45), [a.swingMW, a.rampMWs]);
chk('BESS', near(a.bessMW, 3.45) && near(a.bessMWh, 0.8625) && near(a.residMW, 3.45) && near(a.bessMVA, 3.45 / 0.9), [a.bessMW, a.bessMWh]);
chk('BESS range 40–60%', AI_DYN.bessReduction[0] === 0.4 && AI_DYN.bessReduction[1] === 0.6 && AI_DYN.mitigations.some(m => /capping/i.test(m)) && AI_DYN.mitigations.some(m => /checkpoint/i.test(m)));
// views, report and CSV carry the new results
view = 'data-centre'; render(); const dcHtml = document.querySelector('#main').innerHTML;
chk('view AI section', /AI \/ high-density power planning/.test(dcHtml) && /Grid-forming BESS/.test(dcHtml) && /Platform data as of/.test(dcHtml) && /Mitigations/.test(dcHtml));
const rep = reportDoc(), csv = csvSchedules();
chk('report AI rows', /AI \/ GPU platform/.test(rep) && /Grid-forming BESS/.test(rep) && /Conversion losses/.test(rep) && /not verified code/.test(rep));
chk('csv AI rows', /DATA CENTRE \/ AI POWER PLANNING/.test(csv) && /Grid-forming BESS power/.test(csv) && /Current per rack \(800 V DC\)/.test(csv));
// Kyber planned platform recommends 800 V DC
applyPlatform('NVIDIA Kyber / Rubin Ultra NVL576 (planned)'); chk('Kyber DC', S.dc.arch === AI_ARCH.dc.name && S.dc.density === 600);
// projects saved before these fields existed are upgraded
const old = upgrade({ v: 1, dc: { density: 12 } }); chk('upgrade old project', old.dc.density === 12 && old.dc.platform === AI_CUSTOM && old.dc.arch === AI_ARCH.ac.name && old.dc.liquid === 0);
S.dc = dcSave; R = calcAll();
out.aiDc = [D.it, Math.round(D.fac), Math.round(a.liqHeat), a.cdu, +a.rackAac.toFixed(1), +a.rackAdc.toFixed(1), Math.round(a.convAC), Math.round(a.convDC), +a.bessMW.toFixed(2), +a.bessMWh.toFixed(3)];
out.aiChecks = aiFail.length ? aiFail : 'all passed';

// ---- electrical drawing generator (rooms → luminaires, sockets, circuits, checks, DXF) ----
const dwFail = [], dchk = (name, ok, got) => { if (!ok) dwFail.push(name + (got !== undefined ? ' got ' + JSON.stringify(got) : '')); };
const dwSave = JSON.stringify({ dwg: S.dwg, rooms: S.dwgRooms, country: S.project.country });
const findChk = (W, item, st) => W.chk.filter(c => c.item === item && (!st || c.st === st));
Object.keys(SOCK_BY_COUNTRY).forEach(c => dchk('socket map country exists ' + c, COUNTRIES.some(x => x.name === c)));
S.project.country = 'United Kingdom'; R = calcAll(); let W = R.W;
const [off, meet, corr, wc, tea] = W.rooms;
dchk('UK socket system', W.sock.n === 'BS 1363 13 A (Type G)' && W.sock.ring === 1 && W.lightIn === 6);
dchk('lumen method office', off.nLum === 79 && near(off.Eavg, 79 * 4000 * 0.48 / 300) && off.Eavg >= 500, [off.nLum, off.Eavg]);
dchk('lumen method meeting/corridor/WC/tea', meet.nLum === 8 && corr.nLum === 6 && wc.nLum === 6 && tea.nLum === 5, [meet.nLum, corr.nLum, wc.nLum, tea.nLum]);
dchk('socket guidance', off.nSock === 30 && meet.nSock === 3 && corr.nSock === 1 && wc.nSock === 0 && tea.nSock === 4, W.rooms.map(o => o.nSock));
dchk('emergency counts', off.nEm === 2 && meet.nEm === 0 && corr.nEm === 2 && wc.nEm === 1 && tea.nEm === 1, W.rooms.map(o => o.nEm));
const LC = W.C.filter(c => c.kind === 'L'), PC = W.C.filter(c => c.kind === 'P');
dchk('lighting circuits within 80% of 6 A', LC.length === 3 && LC.every(c => c.w <= 0.8 * 6 * 230 * 0.9 + 1e-6) && LC.reduce((a, c) => a + c.pts, 0) === 79 + 8 + 6 + 6 + 5 + 6, LC.map(c => [c.pts, Math.round(c.w)]));
dchk('every luminaire circuited', W.rooms.every(o => o.lc.reduce((a, x) => a + x[1], 0) === o.nLum + o.nEm) && W.rooms.every(o => o.pc.reduce((a, x) => a + x[1], 0) === o.nSock));
dchk('rings ≤ 100 m²', PC.length === 4 && PC.every(c => c.area <= 100 + 1e-6 && c.cc.ring && c.cc.inr === 32 && c.cc.csa === 2.5), PC.map(c => [Math.round(c.area), c.pts]));
const Lr1 = 2 * 10 + 2 * 35 * (10 / 30);
dchk('ring VD and Zs (OSG method)', near(PC[0].len, Lr1) && near(PC[0].cc.vd, 18 * 6.5 * Lr1 / 1000 / 230 * 100) && near(PC[0].cc.zs, 0.35 + DD.R12TE[CSA.indexOf(2.5)] * Lr1 / 4 * 1.2 / 1000), [PC[0].len, PC[0].cc.vd, PC[0].cc.zs]);
dchk('lighting circuits use 6 A RCBO and pass', LC.every(c => c.cc.inr === 6 && /RCBO/.test(c.cc.dev.n) && /^PASS/.test(c.cc.status)), LC.map(c => c.cc.status));
dchk('example has no failures', W.nBad === 0, W.chk.filter(c => c.st === 'bad'));
dchk('ways check', findChk(W, 'Distribution board ways', 'ok').length === 1);
// parameter checks react to bad inputs
S.dwg.rcd = false; R = calcAll(); dchk('no RCD on sockets fails', findChk(R.W, 'Additional protection by 30 mA RCD', 'bad').length === PC.length);
S.dwg.rcd = true; S.dwgRooms[0].nLum = 40; R = calcAll(); dchk('too few luminaires fails lux', R.W.chk.some(c => c.scope === off.r.name && c.item === 'Maintained illuminance' && c.st === 'bad'));
S.dwgRooms[0].nLum = ''; S.dwgRooms[3].type = 'Bathroom / shower'; S.dwgRooms[3].nSock = 1; R = calcAll(); dchk('bathroom socket warns', findChk(R.W, 'Socket-outlets in bathroom', 'warn').length === 1);
S.dwgRooms[3].type = 'Toilets / WC'; S.dwgRooms[3].nSock = ''; S.dwgRooms[2].nEm = 0; R = calcAll(); dchk('missing emergency fails', R.W.chk.some(c => c.scope === corr.r.name && c.item === 'Emergency lighting' && c.st === 'bad'));
S.dwgRooms[2].nEm = ''; S.dwgRooms[1].L = 0; R = calcAll(); dchk('zero dimension fails', findChk(R.W, 'Room dimensions', 'bad').length === 1);
S.dwgRooms[1].L = 6; S.dwg.ways = 8; R = calcAll(); dchk('few ways warns', findChk(R.W, 'Distribution board ways', 'warn').length === 1, R.W.C.length);
S.dwg.ways = 24; S.dwg.uf = 1.2; R = calcAll(); dchk('UF out of range fails', findChk(R.W, 'Utilisation factor', 'bad').length === 1);
S.dwg.uf = 0.6; S.dwg.sockType = 'Schuko CEE 7/3 16 A (Type F)'; R = calcAll(); dchk('socket mismatch warns', findChk(R.W, 'Socket-outlet system matches region', 'warn').length === 1 && R.W.C.filter(c => c.kind === 'P').every(c => c.pts <= 8 && c.cc.inr === 16));
S.dwg.sockType = '';
// other regions
S.project.country = 'United States'; R = calcAll(); W = R.W;
dchk('US NEMA, 180 VA rule, 20 A', /NEMA/.test(W.sock.n) && W.ptsLim === 13 && W.lightIn === 20 && W.vaSock === 180 && W.C.filter(c => c.kind === 'P').every(c => c.pts <= 13 && c.cc.inr === 20), [W.sock.n, W.ptsLim]);
S.project.country = 'Germany'; R = calcAll(); dchk('Germany Schuko 8 per circuit', /Schuko/.test(R.W.sock.n) && R.W.ptsLim === 8 && R.W.lightIn === 10);
S.project.country = 'Australia'; R = calcAll(); dchk('Australia AS/NZS 3112', /3112/.test(R.W.sock.n));
// DXF and SVG for every country
let dxfBad = [];
for (const c of COUNTRIES) { S.project.country = c.name; R = calcAll(); const G = dwgGeom(R.W), dxf = toDXF(G), svg = toSVG(G); const lines = dxf.split('\r\n'); lines.pop();
  if (!/^0\r\nSECTION\r\n2\r\nHEADER/.test(dxf) || !dxf.endsWith('0\r\nEOF\r\n') || lines.length % 2 || /NaN|undefined|Infinity/.test(dxf + svg) || /[^\x00-\x7f]/.test(dxf) || lines.some((l, i) => i % 2 === 0 && !/^\s*\d+$/.test(l))) dxfBad.push(c.name); }
dchk('DXF valid for all countries', !dxfBad.length, dxfBad);
S.project.country = 'United Kingdom'; R = calcAll();
const G = dwgGeom(R.W), dxf = toDXF(G), cnt = t => (dxf.match(new RegExp('\\r\\n0\\r\\n' + t + '\\r\\n', 'g')) || []).length;
dchk('DXF layers and entities', DWG_LAYERS.every(([n]) => dxf.includes('\r\n2\r\n' + n + '\r\n')) && /AC1009/.test(dxf) && cnt('LINE') > 300 && cnt('CIRCLE') > 20 && cnt('TEXT') > 100, [cnt('LINE'), cnt('CIRCLE'), cnt('TEXT')]);
view = 'drawings'; render(); const dwHtml = document.querySelector('#main').innerHTML;
dchk('drawings view', /Parameter checks/.test(dwHtml) && /<svg/.test(dwHtml) && /Final circuit schedule/.test(dwHtml));
dchk('report + csv', /Lighting and small power – final circuits/.test(reportDoc()) && /FINAL CIRCUITS/.test(csvSchedules()) && /ROOM LOAD SHEET/.test(csvSchedules()));
dchk('old project upgrade', (() => { const o = upgrade({ v: 1, dc: {} }); return Array.isArray(o.dwgRooms) && o.dwgRooms.length === 0 && o.dwg.board === 'DB-L1'; })());
const sv = JSON.parse(dwSave); S.dwg = sv.dwg; S.dwgRooms = sv.rooms; S.project.country = sv.country; R = calcAll();
out.dwg = [R.W.rooms.map(o => [o.nLum, Math.round(o.Eavg), o.nSock, o.nEm]), R.W.C.map(c => [c.ref, c.pts, Math.round(c.w), c.cc.csa, c.cc.status]), [R.W.nOk, R.W.nWarn, R.W.nBad], dxf.length];
out.dwgChecks = dwFail.length ? dwFail : 'all passed';
for (const [k, v] of Object.entries(out)) console.log(k, '=', JSON.stringify(v));
