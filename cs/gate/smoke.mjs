// 배포 전 검사: out/ 을 띄워 놓고 형식마다 열기 → 내려받기 왕복을 확인한다.
// 사용: node cs/gate/smoke.mjs http://127.0.0.1:8080
import { chromium } from 'playwright';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const BASE = process.argv[2] || 'http://127.0.0.1:8080';
const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIX = path.join(HERE, 'fixtures');
const OUTDIR = path.join(HERE, 'out'); fs.mkdirSync(OUTDIR, { recursive: true });

// fmt: 내려받을 형식 코드 (asc_CDownloadOptions), check: 결과 바이트 검사
const CASES = [
  { file: 't.docx', fmt: 65,  ext: 'docx', check: b => isZip(b) && zipHas(b, '한글 본문 첫 문단') },
  { file: 't.xlsx', fmt: 257, ext: 'xlsx', check: b => isZip(b) },
  { file: 't.xlsx', fmt: 258, ext: 'xls',  check: b => isOle(b) },          // 10.0: XLS 저장
  { file: 't.pptx', fmt: 129, ext: 'pptx', check: b => isZip(b) && zipHas(b, '슬라이드 제목') },
  { file: 's.xls',  fmt: 257, ext: 'xlsx', check: b => isZip(b) },
  { file: 's.doc',  fmt: 65,  ext: 'docx', check: b => isZip(b) },
  { file: 's.ppt',  fmt: 129, ext: 'pptx', check: b => isZip(b) },
  { file: 't.pdf',  fmt: 513, ext: 'pdf',  check: b => b.slice(0, 5).toString() === '%PDF-' },
];

const isZip = b => b[0] === 0x50 && b[1] === 0x4b;
const isOle = b => b.slice(0, 8).equals(Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]));
// 압축된 zip 안의 글자는 직접 못 보므로 python 으로 풀어 본다
import { execFileSync } from 'node:child_process';
function zipHas(b, text) {
  const p = path.join(OUTDIR, 'tmp.zip'); fs.writeFileSync(p, b);
  const out = execFileSync('python3', ['-c',
    'import zipfile,sys;z=zipfile.ZipFile(sys.argv[1]);t=sys.argv[2];print(any(t in z.read(n).decode("utf-8","ignore") for n in z.namelist() if n.endswith(".xml")))', p, text]).toString().trim();
  return out === 'True';
}

const browser = await chromium.launch();
const results = [];
let failed = 0;
for (const c of CASES) {
  const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1440, height: 900 } });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(String(e)));
  const t0 = Date.now();
  let ok = false, note = '';
  try {
    await page.goto(BASE + '/');
    await page.setInputFiles('#upload-input', path.join(FIX, c.file));
    const frame = await waitFrame(page, 90000);
    await frame.waitForFunction(() => window.Asc && Asc.editor && (Asc.editor.isDocumentLoadComplete || Asc.editor.DocumentRenderer), null, { timeout: 120000 });
    await page.waitForTimeout(2000);
    const openMs = Date.now() - t0;
    const dl = page.waitForEvent('download', { timeout: 120000 });
    await frame.evaluate(fmt => Asc.editor.asc_DownloadAs(new Asc.asc_CDownloadOptions(fmt)), c.fmt);
    const d = await dl;
    const out = path.join(OUTDIR, `${c.file}.${c.ext}`);
    await d.saveAs(out);
    const bytes = fs.readFileSync(out);
    ok = c.check(bytes);
    note = `열기 ${openMs}ms, 결과 ${bytes.length}B`;
  } catch (e) {
    note = String(e).split('\n')[0];
  }
  if (errors.length) note += ` / 페이지 오류 ${errors.length}건: ${errors[0].slice(0, 120)}`;
  if (!ok) failed++;
  results.push({ case: `${c.file} -> ${c.ext}`, ok, note });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${c.file} -> ${c.ext}  ${note}`);
  await ctx.close();
}
await browser.close();
fs.writeFileSync(path.join(HERE, 'report.json'), JSON.stringify(results, null, 2));
process.exit(failed ? 1 : 0);

async function waitFrame(page, timeout) {
  const end = Date.now() + timeout;
  while (Date.now() < end) {
    const f = page.frames().find(fr => /\/web-apps\/apps\/\w+editor\/main\/index\.html/.test(fr.url()));
    if (f) return f;
    await page.waitForTimeout(250);
  }
  throw new Error('편집기 frame 이 뜨지 않음');
}
