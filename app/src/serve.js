// Local preview server for load-designer.html: node serve.js  ->  http://localhost:5179/
const http = require('http'), fs = require('fs'), path = require('path');
const file = path.join(__dirname, '..', 'load-designer.html');
http.createServer((req, res) => {
  fs.readFile(file, (err, buf) => {
    if (err) { res.writeHead(500); res.end(String(err)); return; }
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' }); res.end(buf);
  });
}).listen(5179, () => console.log('serving on http://localhost:5179/'));
