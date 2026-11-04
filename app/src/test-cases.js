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
for (const [k, v] of Object.entries(out)) console.log(k, '=', JSON.stringify(v));
