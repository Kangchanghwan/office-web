// cbEdit v1 확인: localhost 부모 페이지에서 편집기를 iframe 으로 열고 ready -> open -> (입력/dirty) -> export 를 형식마다 확인한다.
// 사용: node cs/gate/embed.mjs http://127.0.0.1:8080
import { chromium } from 'playwright';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const BASE = process.argv[2] || 'http://127.0.0.1:8080';
const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIX = path.join(HERE, 'fixtures');
const PARENT = 'http://localhost:9911';

const isZip = b => b[0] === 0x50 && b[1] === 0x4b;
const isOle = b => b.length > 8 && b.slice(0, 8).equals(Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]));
const isPdf = b => b.slice(0, 5).toString() === '%PDF-';
// type: 입력으로 dirty 를 시험할 형식
const CASES = [
  { file: 't.docx', ext: 'docx', check: isZip, type: 'word' },
  { file: 't.xlsx', ext: 'xlsx', check: isZip, type: 'cell' },
  { file: 't.pdf',  ext: 'pdf',  check: isPdf },
  { file: 's.doc',  ext: 'docx', check: isZip },
  { file: 't.pptx', ext: 'pptx', check: isZip },
  { file: 's.xls',  ext: 'xls',  check: isOle },
  { file: 's.ppt',  ext: 'pptx', check: isZip },
];

const PARENT_HTML = `<!doctype html><meta charset=utf-8><body style="margin:0">
<iframe id=f src="${BASE}/?embed=cs&theme=light" style="width:1280px;height:800px;border:0"></iframe>
<script>
window.events = []; window.pending = {}; let n = 0;
addEventListener('message', e => {
  if (e.origin !== ${JSON.stringify(new URL(BASE).origin)} || e.source !== document.getElementById('f').contentWindow) return;
  const d = e.data; if (!d || d.cbEdit !== 1) return;
  if (d.event) { window.events.push(d); return; }
  const p = window.pending[d.id]; if (p) { delete window.pending[d.id]; p(d); }
});
window.call = (method, args, transfer) => new Promise(res => {
  const id = 'r' + (++n); window.pending[id] = res;
  document.getElementById('f').contentWindow.postMessage({ cbEdit: 1, id, method, args }, ${JSON.stringify(new URL(BASE).origin)}, transfer || []);
});
</script>`;

const browser = await chromium.launch();
const results = [];
let failed = 0;
let ctx, page;
async function session() {
  if (ctx) await ctx.close();
  ctx = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  await ctx.route(PARENT + '/**', r => r.fulfill({ status: 200, contentType: 'text/html', body: PARENT_HTML }));
  page = await ctx.newPage();
  page.on('pageerror', e => console.log('  pageerror', String(e).slice(0, 120)));
  await page.goto(PARENT + '/p.html');
  await page.waitForFunction(() => window.events.some(e => e.event === 'ready'), null, { timeout: 120000 });
}
const call = (method, args) => page.evaluate(async ([m, a]) => {
  const r = await window.call(m, a);
  if (r.ok && r.result && r.result.bytes instanceof ArrayBuffer) r.result.bytes = Array.from(new Uint8Array(r.result.bytes));
  return r;
}, [method, args]);
async function openCase(c) {
  const buf = fs.readFileSync(path.join(FIX, c.file));
  return page.evaluate(async ([name, arr]) => {
    const ab = new Uint8Array(arr).buffer;
    return await window.call('open', { name, bytes: ab }, [ab]);
  }, [c.file, Array.from(buf)]);
}
async function editorFrame() {
  const end = Date.now() + 20000;
  while (Date.now() < end) {
    for (const fr of page.frames()) if (/\/web-apps\/apps\/\w+editor\/main\/index\.html/.test(fr.url())) return fr;
    await page.waitForTimeout(250);
  }
  return null;
}

await session();
const ready = await page.evaluate(() => window.events.find(e => e.event === 'ready'));
console.log('ready', JSON.stringify({ editor: ready.editor, version: ready.version, accepts: ready.accepts, chrome: ready.chrome }));
const unk = await call('nope', {});
console.log(unk.ok === false && unk.error.code === 'UNKNOWN_METHOD' ? 'PASS  UNKNOWN_METHOD' : 'FAIL  UNKNOWN_METHOD');
if (!(unk.ok === false && unk.error.code === 'UNKNOWN_METHOD')) failed++;
const noDoc = await call('export', {});
console.log(noDoc.ok === false && noDoc.error.code === 'NO_DOCUMENT' ? 'PASS  NO_DOCUMENT' : 'FAIL  NO_DOCUMENT ' + JSON.stringify(noDoc));
if (!(noDoc.ok === false && noDoc.error.code === 'NO_DOCUMENT')) failed++;

for (const c of CASES) {
  let ok = false, note = '';
  const t0 = Date.now();
  try {
    const o = await openCase(c);
    if (!o.ok) throw new Error('open: ' + JSON.stringify(o.error));
    const openMs = Date.now() - t0;
    const st = await call('state', {});
    if (!(st.ok && st.result.loaded && st.result.dirty === false && st.result.name === c.file)) throw new Error('state ' + JSON.stringify(st));
    let dirtyNote = '';
    if (c.type) {
      try {
        const fr = await editorFrame();
        await page.evaluate(() => window.events.length = 0);
        if (c.type === 'word') await fr.locator('#id_viewer_overlay, #id_target_cursor, canvas').first().click({ position: { x: 300, y: 200 }, timeout: 5000 });
        else await fr.locator('#id_viewer_overlay, canvas').first().click({ position: { x: 200, y: 100 }, timeout: 5000 });
        await page.keyboard.type('Z');
        await page.waitForFunction(() => window.events.some(e => e.event === 'dirty' && e.dirty === true), null, { timeout: 8000 });
        dirtyNote = ' dirty:OK';
      } catch (e) { dirtyNote = ' dirty:미확인(' + String(e).split('\n')[0].slice(0, 60) + ')'; }
    }
    const x = await call('export', {});
    if (!x.ok) throw new Error('export: ' + JSON.stringify(x.error));
    const bytes = Buffer.from(x.result.bytes);
    if (x.result.ext !== c.ext) throw new Error('ext ' + x.result.ext + ' != ' + c.ext);
    if (!c.check(bytes)) throw new Error('서명 불일치 ' + bytes.slice(0, 8).toString('hex'));
    if (dirtyNote.includes('OK')) {
      const st2 = await call('state', {});
      if (!st2.result.dirty) throw new Error('export 가 dirty 를 바꿈');
      await call('markSaved', {});
      const st3 = await call('state', {});
      if (st3.result.dirty) throw new Error('markSaved 후에도 dirty');
      dirtyNote += ' markSaved:OK';
    }
    ok = true;
    note = `open ${openMs}ms, pages ${o.result.pages ?? '-'}, ${bytes.length}B ext ${x.result.ext}${dirtyNote}`;
  } catch (e) {
    note = String(e).split('\n')[0].slice(0, 200);
  }
  if (!ok) { failed++; try { await session(); } catch (e) { note += ' (세션 복구 실패)'; } }
  results.push({ case: `${c.file} -> ${c.ext}`, ok, note });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${c.file} -> ${c.ext}  ${note}`);
}
const chromeEv = await page.evaluate(() => window.events.filter(e => e.event === 'chrome').slice(-1)[0] || null);
console.log('last chrome event', JSON.stringify(chromeEv));
await browser.close();
fs.writeFileSync(path.join(HERE, 'embed-report.json'), JSON.stringify(results, null, 2));
process.exit(failed ? 1 : 0);
