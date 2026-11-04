// Builds app/load-designer.html: injects design data extracted from the workbook generator into src/app.html.
// Usage: node build-app.js <path-to-designdata.json>
const fs = require('fs');
const path = require('path');
const data = fs.readFileSync(process.argv[2], 'utf8').replace(/</g, '\\u003c');
const tpl = fs.readFileSync(path.join(__dirname, 'app.html'), 'utf8');
const out = path.join(__dirname, '..', 'load-designer.html');
fs.writeFileSync(out, tpl.replace('/*DATA*/', () => data));
console.log('written', out, fs.statSync(out).size, 'bytes');
