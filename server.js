const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 3000;

const mimeTypes = {
    '.html': 'text/html',
    '.js': 'text/javascript',
    '.css': 'text/css',
    '.json': 'application/json',
    '.png': 'image/png',
    '.jpg': 'image/jpg',
    '.gif': 'image/gif',
    '.svg': 'image/svg+xml',
    '.xml': 'application/xml'
};

const server = http.createServer((req, res) => {
    console.log(`${req.method} ${req.url}`);
    
    // تجاهل favicon
    if (req.url === '/favicon.ico') {
        res.writeHead(204);
        res.end();
        return;
    }
    
    let filePath = '.' + req.url;
    if (filePath === './') {
        filePath = './taskpane.html';
    }
    
    // إنشاء مجلد assets وهمي للأيقونات
    if (req.url.startsWith('/assets/')) {
        res.writeHead(200, { 'Content-Type': 'image/png' });
        res.end();
        return;
    }
    
    const extname = String(path.extname(filePath)).toLowerCase();
    const contentType = mimeTypes[extname] || 'application/octet-stream';
    
    fs.readFile(filePath, (error, content) => {
        if (error) {
            if (error.code == 'ENOENT') {
                res.writeHead(404);
                res.end(`File not found: ${filePath}`);
            } else {
                res.writeHead(500);
                res.end(`Server error: ${error.code}`);
            }
        } else {
            res.writeHead(200, { 
                'Content-Type': contentType,
                'Access-Control-Allow-Origin': '*'
            });
            res.end(content, 'utf-8');
        }
    });
});

server.listen(PORT, () => {
    console.log(`✅ Server running at http://localhost:${PORT}/`);
    console.log(`📄 Taskpane: http://localhost:${PORT}/taskpane.html`);
    console.log('🛑 Press Ctrl+C to stop');
});
