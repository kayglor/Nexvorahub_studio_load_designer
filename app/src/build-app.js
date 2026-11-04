// Builds the app from src/app.html plus design data extracted from the workbook generator.
//   app/load-designer.html  – page content for the claude.ai Artifact
//   site/index.html         – complete standalone document for web hosting (Vercel)
// Usage: node build-app.js <path-to-designdata.json>
const fs = require('fs');
const path = require('path');
const data = fs.readFileSync(process.argv[2], 'utf8').replace(/</g, '\\u003c');
const tpl = fs.readFileSync(path.join(__dirname, 'app.html'), 'utf8');
const page = tpl.replace('/*DATA*/', () => data);

const out = path.join(__dirname, '..', 'load-designer.html');
fs.writeFileSync(out, page);
console.log('written', out, fs.statSync(out).size, 'bytes');

const reset = ':root{color-scheme:light;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}body{margin:0;font:14px/1.4 system-ui,sans-serif}img{max-width:100%}[hidden]{display:none!important}';
const site = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="description" content="Electrical load assessment and design report generator for 94 countries and regions.">
<style>${reset}</style>
</head>
<body>
${page}
</body>
</html>
`;
const siteDir = path.join(__dirname, '..', '..', 'site');
fs.mkdirSync(siteDir, { recursive: true });
const siteOut = path.join(siteDir, 'index.html');
fs.writeFileSync(siteOut, site);
console.log('written', siteOut, fs.statSync(siteOut).size, 'bytes');
