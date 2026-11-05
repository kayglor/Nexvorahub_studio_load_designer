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
for (const [k, v] of Object.entries(out)) console.log(k, '=', JSON.stringify(v));
