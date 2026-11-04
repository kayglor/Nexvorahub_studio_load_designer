const fs = require('fs');
let b = fs.readFileSync('build.js', 'utf8');
b = b.replace("  ['Student bedroom (cluster flat)', 2, 0, 'Typical 1.5–2.5 kVA per bedroom'],\n  ['', null, null, ''],",
  "  ['Student bedroom (cluster flat)', 2, 0, 'Typical 1.5–2.5 kVA per bedroom'],\n  ['Furnished flat – all-electric with appliances', 5, 0, 'Typical 4–6 kVA'],");
const head = b.slice(0, b.indexOf('// ---------------- 1. Project Info'));
const mid = b.slice(b.indexOf('// ---------------- 4. Residential ADMD'), b.indexOf('// ---------------- 6. Office Breakdown'));
const files = ['pi_section.js', 'data_section_v2.js', null, 'ns_part.js', 'new_sections2.js', 'external_section.js', 'code_calc_section.js', 'cable_section.js', 'lifts_pumps_section.js', 'systems_section.js', 'energy_section.js', 'datacentre_section.js', 'summary_app.js'];
const parts = [head, ...files.map(f => f ? fs.readFileSync(f, 'utf8') : mid)];
fs.writeFileSync('build_full.js', parts.join('\n'));
console.log('assembled');
