import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { createReadStream } from 'node:fs';
import { extname, join, normalize } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(fileURLToPath(new URL('.', import.meta.url)), 'dist');
const port = Number(process.env.PORT || 3000);
const types = {'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.json':'application/json; charset=utf-8','.webmanifest':'application/manifest+json','.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg','.svg':'image/svg+xml','.woff2':'font/woff2','.ttf':'font/ttf'};

createServer(async (req,res)=>{
  try {
    const pathname = decodeURIComponent(new URL(req.url || '/', 'http://localhost').pathname);
    const safe = normalize(pathname).replace(/^([.][.][/\\])+/, '');
    let file = join(root, safe);
    let info;
    try { info = await stat(file); } catch {}
    if (!info?.isFile()) file = join(root, 'index.html');
    const ext = extname(file).toLowerCase();
    res.setHeader('Content-Type', types[ext] || 'application/octet-stream');
    res.setHeader('Cache-Control', 'no-cache');
    createReadStream(file).pipe(res);
  } catch {
    res.statusCode = 500;
    res.end('Internal Server Error');
  }
}).listen(port, '0.0.0.0', ()=>console.log(`Student Follow-up Demo listening on ${port}`));
