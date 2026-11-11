/* 零依赖静态服务器：开发预览 / 在线加载项托管 */
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const PORT = Number(process.env.PORT) || 3890;
const HOST = process.env.HOST || '127.0.0.1';

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.htm': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.xml': 'text/xml; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.ico': 'image/x-icon'
};

const INDEX = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<title>智能表格助手 · 本地预览</title>
<style>body{font-family:"Microsoft YaHei",sans-serif;background:#f6f7fb;color:#1c2233;margin:0;padding:48px;display:flex;justify-content:center}
.card{background:#fff;border-radius:16px;padding:28px 32px;box-shadow:0 8px 24px rgba(29,40,84,.08);max-width:560px}
h1{font-size:20px;margin:0 0 8px}p{color:#5b647d;line-height:1.8;font-size:14px}
a{display:inline-block;margin:8px 8px 0 0;padding:10px 16px;border-radius:10px;background:#4f6bff;color:#fff;text-decoration:none;font-size:13px}
a.secondary{background:#eef1ff;color:#4f6bff}code{background:#f2f4fa;padding:2px 6px;border-radius:5px}</style></head>
<body><div class="card"><h1>智能表格助手</h1>
<p>这是开发用的静态预览服务，只用于查看侧边栏界面。浏览器里没有真实表格宿主，会显示「未连接真实表格」：公式建议可正常试用，执行操作会被明确拦截，不会构造虚拟表格或写入任何内容。真实使用请安装到表格客户端。</p>
<a href="/addon/ui/taskpane.html">打开侧边栏预览</a>
<a class="secondary" href="/addon/ribbon.xml">查看 ribbon.xml</a>
<p style="margin-top:20px">要安装到本机表格客户端：运行 <code>npm run install:addon</code>，然后完全退出并重启客户端，在功能区的「智能表格助手」选项卡点击「打开侧边栏」。</p>
<p>在线调试时也可把 publish.xml 的 url 改为 <code>http://${HOST}:${PORT}/addon/</code>。</p>
</div></body></html>`;

const server = http.createServer((req, res) => {
  let urlPath;
  try { urlPath = decodeURIComponent(req.url.split('?')[0]); }
  catch (e) { urlPath = '/'; }
  if (urlPath === '/' || urlPath === '') {
    res.writeHead(200, { 'Content-Type': MIME['.html'], 'Access-Control-Allow-Origin': '*' });
    res.end(INDEX);
    return;
  }
  if (/\/$/.test(urlPath)) { urlPath += 'index.html'; }
  const filePath = path.resolve(ROOT, '.' + urlPath);
  if (filePath !== ROOT && !filePath.startsWith(ROOT + path.sep)) {
    res.writeHead(403); res.end('Forbidden'); return;
  }
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('404 Not Found: ' + urlPath);
      return;
    }
    res.writeHead(200, {
      'Content-Type': MIME[path.extname(filePath).toLowerCase()] || 'application/octet-stream',
      'Cache-Control': 'no-store',
      'Access-Control-Allow-Origin': '*'
    });
    res.end(data);
  });
});

server.listen(PORT, HOST, () => {
  console.log('智能表格助手预览服务已启动');
  console.log('  侧边栏预览： http://' + HOST + ':' + PORT + '/addon/ui/taskpane.html');
  console.log('  加载项目录： http://' + HOST + ':' + PORT + '/addon/');
  console.log('按 Ctrl+C 停止');
});
