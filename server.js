const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 3000;
const DATA_FILE = './data.json'; // حفظ كل البيانات هنا

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

// قراءة البيانات من الملف
function loadData() {
    try {
        if (fs.existsSync(DATA_FILE)) {
            return JSON.parse(fs.readFileSync(DATA_FILE, 'utf-8'));
        }
    } catch (e) {}
    return {
        usageStats: { totalTranslations: 0, totalWords: 0, totalInputTokens: 0, totalOutputTokens: 0, totalCost: 0 },
        cacheStats: { hits: 0, savedCost: 0 },
        translationCache: {},
        translationLog: []
    };
}

// حفظ البيانات في الملف
function saveData(data) {
    fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf-8');
}

// قراءة body من الـ request
function readBody(req) {
    return new Promise((resolve, reject) => {
        let body = '';
        req.on('data', chunk => body += chunk.toString());
        req.on('end', () => {
            try { resolve(JSON.parse(body)); }
            catch (e) { resolve({}); }
        });
        req.on('error', reject);
    });
}

function corsHeaders(res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

const server = http.createServer(async (req, res) => {
    console.log(`${req.method} ${req.url}`);
    corsHeaders(res);

    // Handle preflight
    if (req.method === 'OPTIONS') {
        res.writeHead(204); res.end(); return;
    }

    // ===== API ENDPOINTS =====

    // GET /api/stats → إرجاع كل البيانات
    if (req.method === 'GET' && req.url === '/api/stats') {
        const data = loadData();
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(data));
        return;
    }

    // POST /api/stats → تحديث الإحصائيات
    if (req.method === 'POST' && req.url === '/api/stats') {
        const body = await readBody(req);
        const data = loadData();
        if (body.usageStats) data.usageStats = body.usageStats;
        if (body.cacheStats) data.cacheStats = body.cacheStats;
        if (body.translationCache) data.translationCache = body.translationCache;
        saveData(data);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true }));
        return;
    }

    // POST /api/log → تسجيل ترجمة واحدة
    if (req.method === 'POST' && req.url === '/api/log') {
        const body = await readBody(req);
        const data = loadData();
        data.translationLog.push({ ...body, timestamp: new Date().toISOString() });
        // احتفظ بآخر 500 سجل بس
        if (data.translationLog.length > 500) data.translationLog = data.translationLog.slice(-500);
        saveData(data);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true }));
        return;
    }

    // POST /api/reset-stats → reset الإحصائيات
    if (req.method === 'POST' && req.url === '/api/reset-stats') {
        const data = loadData();
        data.usageStats = { totalTranslations: 0, totalWords: 0, totalInputTokens: 0, totalOutputTokens: 0, totalCost: 0 };
        saveData(data);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true }));
        return;
    }

    // POST /api/reset-cache → reset الكاش
    if (req.method === 'POST' && req.url === '/api/reset-cache') {
        const data = loadData();
        data.translationCache = {};
        data.cacheStats = { hits: 0, savedCost: 0 };
        saveData(data);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true }));
        return;
    }

    // POST /api/reset-log → مسح السجل
    if (req.method === 'POST' && req.url === '/api/reset-log') {
        const data = loadData();
        data.translationLog = [];
        saveData(data);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ ok: true }));
        return;
    }

    // ===== DASHBOARD =====
    if (req.url === '/dashboard' || req.url === '/dashboard.html') {
        res.writeHead(200, { 'Content-Type': 'text/html' });
        res.end(getDashboardHTML());
        return;
    }

    // تجاهل favicon
    if (req.url === '/favicon.ico') { res.writeHead(204); res.end(); return; }

    let filePath = '.' + req.url;
    if (filePath === './') filePath = './taskpane.html';

    if (req.url.startsWith('/assets/')) {
        res.writeHead(200, { 'Content-Type': 'image/png' }); res.end(); return;
    }

    const extname = String(path.extname(filePath)).toLowerCase();
    const contentType = mimeTypes[extname] || 'application/octet-stream';

    fs.readFile(filePath, (error, content) => {
        if (error) {
            res.writeHead(error.code === 'ENOENT' ? 404 : 500);
            res.end(error.code === 'ENOENT' ? `File not found: ${filePath}` : `Server error: ${error.code}`);
        } else {
            res.writeHead(200, { 'Content-Type': contentType, 'Access-Control-Allow-Origin': '*' });
            res.end(content, 'utf-8');
        }
    });
});

server.listen(PORT, () => {
    console.log(`✅ Server running at http://localhost:${PORT}/`);
    console.log(`📄 Taskpane:  http://localhost:${PORT}/taskpane.html`);
    console.log(`📊 Dashboard: http://localhost:${PORT}/dashboard`);
    console.log('🛑 Press Ctrl+C to stop');
});

// ===== DASHBOARD HTML =====
function getDashboardHTML() {
    return `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dar Al Marjaan - Dashboard</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Tahoma, sans-serif; background: #f0f2f5; color: #2d2d2d; }
.header {
    background: linear-gradient(135deg, #1a1a1a, #2d2d2d);
    color: #C9A961; padding: 20px 32px;
    display: flex; align-items: center; justify-content: space-between;
}
.header h1 { font-size: 20px; letter-spacing: 1px; }
.header small { color: #999; font-size: 12px; display: block; margin-top: 3px; }
.refresh-btn {
    background: #C9A961; color: #1a1a1a; border: none; padding: 8px 18px;
    border-radius: 6px; font-weight: 700; cursor: pointer; font-size: 13px;
}
.refresh-btn:hover { background: #b8954e; }
.container { max-width: 1100px; margin: 0 auto; padding: 24px; }
.stats-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin-bottom: 24px; }
.stat-card {
    background: white; border-radius: 12px; padding: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07); border-top: 4px solid #C9A961;
    text-align: center;
}
.stat-card .val { font-size: 28px; font-weight: 700; color: #2d2d2d; margin-bottom: 4px; }
.stat-card .val.gold { color: #C9A961; }
.stat-card .lbl { font-size: 11px; color: #999; text-transform: uppercase; letter-spacing: 0.5px; }
.section { background: white; border-radius: 12px; padding: 20px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); }
.section h2 { font-size: 14px; font-weight: 700; color: #2d2d2d; margin-bottom: 16px; display: flex; align-items: center; gap: 8px; border-bottom: 2px solid #f0f0f0; padding-bottom: 10px; }
.actions { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 16px; }
.btn-reset {
    background: #dc3545; color: white; border: none; padding: 8px 16px;
    border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer;
}
.btn-reset:hover { background: #c82333; }
.btn-info {
    background: #17a2b8; color: white; border: none; padding: 8px 16px;
    border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer;
}
table { width: 100%; border-collapse: collapse; font-size: 12px; }
th { background: #f8f9fa; padding: 10px 12px; text-align: right; font-weight: 700; color: #555; border-bottom: 2px solid #e8e8e8; }
td { padding: 10px 12px; border-bottom: 1px solid #f0f0f0; color: #444; vertical-align: top; }
tr:hover td { background: #fffdf5; }
.badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 10px; font-weight: 700; }
.badge-green { background: #d4edda; color: #155724; }
.badge-blue { background: #d1ecf1; color: #0c5460; }
.badge-gold { background: #fff3cd; color: #856404; }
.cache-item { padding: 6px 0; border-bottom: 1px solid #f5f5f5; font-size: 12px; display: flex; gap: 8px; align-items: center; }
.cache-src { color: #555; flex: 1; }
.cache-arrow { color: #C9A961; font-weight: 700; }
.cache-tgt { color: #2d2d2d; flex: 1; }
.empty-state { text-align: center; color: #bbb; padding: 30px; font-size: 13px; }
.last-update { font-size: 11px; color: #bbb; text-align: center; margin-top: 12px; }
</style>
</head>
<body>
<div class="header">
    <div>
        <h1>🌟 Dar Al Marjaan — Translation Dashboard</h1>
        <small id="lastUpdate">Loading...</small>
    </div>
    <button class="refresh-btn" onclick="loadAll()">🔄 Refresh</button>
</div>
<div class="container">

    <!-- Stats Cards -->
    <div class="stats-row" id="statsCards">
        <div class="stat-card"><div class="val" id="s-trans">—</div><div class="lbl">Total Translations</div></div>
        <div class="stat-card"><div class="val" id="s-words">—</div><div class="lbl">Total Words</div></div>
        <div class="stat-card"><div class="val" id="s-in">—</div><div class="lbl">Input Tokens</div></div>
        <div class="stat-card"><div class="val" id="s-out">—</div><div class="lbl">Output Tokens</div></div>
        <div class="stat-card"><div class="val gold" id="s-cost">—</div><div class="lbl">Total Cost (USD)</div></div>
        <div class="stat-card"><div class="val" id="s-hits">—</div><div class="lbl">Cache Hits</div></div>
        <div class="stat-card"><div class="val gold" id="s-saved">—</div><div class="lbl">Cost Saved by Cache</div></div>
        <div class="stat-card"><div class="val" id="s-pairs">—</div><div class="lbl">Saved Pairs</div></div>
    </div>

    <!-- Actions -->
    <div class="section">
        <h2>⚙️ Actions</h2>
        <div class="actions">
            <button class="btn-reset" onclick="resetStats()">🗑️ Reset Usage Stats</button>
            <button class="btn-reset" onclick="resetCache()">🗑️ Clear Translation Memory</button>
            <button class="btn-reset" onclick="resetLog()">🗑️ Clear Translation Log</button>
            <button class="btn-info" onclick="loadAll()">🔄 Refresh Data</button>
        </div>
    </div>

    <!-- Translation Log -->
    <div class="section">
        <h2>📋 Translation Log (Last 100)</h2>
        <div id="logTable"><div class="empty-state">Loading...</div></div>
    </div>

    <!-- Cache Pairs -->
    <div class="section">
        <h2>🧠 Saved Translation Pairs</h2>
        <div id="cacheList"><div class="empty-state">Loading...</div></div>
    </div>
</div>

<script>
async function loadAll() {
    try {
        const r = await fetch('/api/stats');
        const data = await r.json();
        renderStats(data);
        renderLog(data.translationLog || []);
        renderCache(data.translationCache || {});
        document.getElementById('lastUpdate').textContent = 'Last updated: ' + new Date().toLocaleString('ar-EG');
    } catch(e) {
        alert('Error loading data: ' + e.message);
    }
}

function renderStats(data) {
    const u = data.usageStats || {};
    const c = data.cacheStats || {};
    document.getElementById('s-trans').textContent = (u.totalTranslations||0).toLocaleString();
    document.getElementById('s-words').textContent = (u.totalWords||0).toLocaleString();
    document.getElementById('s-in').textContent = (u.totalInputTokens||0).toLocaleString();
    document.getElementById('s-out').textContent = (u.totalOutputTokens||0).toLocaleString();
    document.getElementById('s-cost').textContent = '$' + (u.totalCost||0).toFixed(4);
    document.getElementById('s-hits').textContent = (c.hits||0).toLocaleString();
    document.getElementById('s-saved').textContent = '$' + (c.savedCost||0).toFixed(4);
    document.getElementById('s-pairs').textContent = Object.keys(data.translationCache||{}).length.toLocaleString();
}

function renderLog(log) {
    if (!log.length) {
        document.getElementById('logTable').innerHTML = '<div class="empty-state">No translations logged yet</div>';
        return;
    }
    const rows = [...log].reverse().slice(0, 100).map(l => \`
        <tr>
            <td>\${new Date(l.timestamp).toLocaleString('ar-EG')}</td>
            <td><span class="badge badge-blue">\${l.sourceLang||'auto'}</span> → <span class="badge badge-blue">\${l.targetLang||'ar'}</span></td>
            <td><span class="badge badge-gold">\${l.model||'—'}</span></td>
            <td><span class="badge badge-green">\${l.mode||'general'}</span></td>
            <td>\${(l.words||0).toLocaleString()}</td>
            <td>\${(l.inputTokens||0).toLocaleString()}</td>
            <td>\${(l.outputTokens||0).toLocaleString()}</td>
            <td>$\${parseFloat(l.cost||0).toFixed(5)}</td>
        </tr>\`).join('');
    document.getElementById('logTable').innerHTML = \`
        <table>
            <thead><tr><th>التوقيت</th><th>اللغة</th><th>الموديل</th><th>النوع</th><th>كلمات</th><th>Input Tokens</th><th>Output Tokens</th><th>التكلفة</th></tr></thead>
            <tbody>\${rows}</tbody>
        </table>\`;
}

function renderCache(cache) {
    const entries = Object.entries(cache);
    if (!entries.length) {
        document.getElementById('cacheList').innerHTML = '<div class="empty-state">No saved pairs yet</div>';
        return;
    }
    const html = entries.map(([key, val]) => {
        const src = key.split('_').slice(2).join('_');
        return \`<div class="cache-item"><span class="cache-src">\${src}</span><span class="cache-arrow">→</span><span class="cache-tgt">\${val}</span></div>\`;
    }).join('');
    document.getElementById('cacheList').innerHTML = html;
}

async function resetStats() {
    if (!confirm('Reset all usage statistics?')) return;
    await fetch('/api/reset-stats', { method: 'POST' });
    loadAll();
}
async function resetCache() {
    if (!confirm('Clear all saved translation pairs?')) return;
    await fetch('/api/reset-cache', { method: 'POST' });
    loadAll();
}
async function resetLog() {
    if (!confirm('Clear the translation log?')) return;
    await fetch('/api/reset-log', { method: 'POST' });
    loadAll();
}

loadAll();
setInterval(loadAll, 30000); // auto-refresh every 30s
</script>
</body>
</html>`;
}
