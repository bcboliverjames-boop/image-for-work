const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 3001;
const LOG_DIR = path.join(__dirname, '..', 'logs');

// 确保日志目录存在
if (!fs.existsSync(LOG_DIR)) {
  fs.mkdirSync(LOG_DIR, { recursive: true });
}

const server = http.createServer((req, res) => {
  // 设置 CORS 头
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(200);
    res.end();
    return;
  }

  if (req.method === 'POST' && req.url === '/log') {
    let body = '';
    req.on('data', chunk => {
      body += chunk.toString();
    });
    req.on('end', () => {
      try {
        const logData = JSON.parse(body);
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const filename = `log-${timestamp}.json`;
        const filepath = path.join(LOG_DIR, filename);
        
        // 保存到文件
        fs.writeFileSync(filepath, JSON.stringify(logData, null, 2));
        
        // 打印到控制台
        const logs = logData.logs || [];
        console.log(`\n=== Received ${logs.length} logs ===`);
        logs.forEach(log => {
          const ts = log.timestamp ? log.timestamp.slice(11, 23) : '';
          console.log(`[${ts}] [${log.level}] [${log.module}] ${log.message}`, log.data || '');
        });
        console.log(`=== Saved to ${filename} ===\n`);
        
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true, filename: filename }));
      } catch (e) {
        console.log(`[ERROR] Failed to parse log: ${e.message}`);
        res.writeHead(400, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Invalid JSON' }));
      }
    });
  } else {
    res.writeHead(404);
    res.end('Not Found');
  }
});

server.listen(PORT, () => {
  console.log(`Log server running on http://localhost:${PORT}`);
  console.log(`Logs will be saved to: ${LOG_DIR}`);
});
