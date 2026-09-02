// 复现 pptx→pdf "memory access out of bounds"：
// 用自编译 x2t.wasm（node）调 x2t_render_bin_to_pdf，传入不同大小的"渲染 bin"，
// 观察 >8MB（wasm 栈大小）时是否触发 wasm 陷阱。
// 渲染 bin 格式仿 sdkjs CDocumentRenderer 输出（MetafileToRenderer.cpp 重放）：
//   ctPageStart(202) ctPageWidth(200,4B) ctPageHeight(201,4B)
//   ctPathCommandStart(100) [ctPathCommandMoveTo(91,8B)|ctPathCommandLineTo(92,8B)|
//    ctPathCommandCurveTo(94,24B)]* ctPathCommandEnd(98) ctDrawPath(99,4B) ctPageEnd(203)
import { createRequire } from 'node:module';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
// opensource ships the built wasm in web/vendor (no core/ build tree)
const BIN = process.env.X2T_BIN
  || path.join(ROOT, 'web/vendor/sdkjs/common/wasm/x2t');
const require = createRequire(import.meta.url);
globalThis.__dirname = BIN;
process.chdir(BIN);

const x2t = require(path.join(BIN, 'x2t.js'));

function makeRenderBin(targetBytes) {
  // WriteDouble = int32 (val*100000)；页宽/高单位 cm？官方用 mm/100… 用 20cm/10cm
  const enc = (v) => { const b = Buffer.alloc(4); b.writeInt32LE(Math.round(v * 100000), 0); return b; };
  const parts = [];
  parts.push(Buffer.from([202]));       // ctPageStart
  parts.push(Buffer.from([200])); parts.push(enc(20)); // ctPageWidth
  parts.push(Buffer.from([201])); parts.push(enc(10)); // ctPageHeight
  parts.push(Buffer.from([100]));       // ctPathCommandStart
  let bytes = 0;
  let i = 0;
  while (bytes < targetBytes) {
    const t = i % 3;
    if (t === 0) { parts.push(Buffer.from([91])); parts.push(enc(1 + (i % 15)), enc(2 + (i % 7))); bytes += 9; }
    else if (t === 1) { parts.push(Buffer.from([92])); parts.push(enc(3 + (i % 11)), enc(4 + (i % 5))); bytes += 9; }
    else { parts.push(Buffer.from([94])); for (let k = 0; k < 6; k++) parts.push(enc(1 + ((i + k) % 13))); bytes += 25; }
    i++;
  }
  parts.push(Buffer.from([98]));        // ctPathCommandEnd
  parts.push(Buffer.from([99])); parts.push(enc(256)); // ctDrawPath fill
  parts.push(Buffer.from([203]));       // ctPageEnd
  return Buffer.concat(parts);
}

function writeFS(p, buf) {
  const stream = x2t.FS.open(p, 'w');
  x2t.FS.write(stream, buf, 0, buf.length, 0);
  x2t.FS.close(stream);
}

x2t.onRuntimeInitialized = function () {
  try { x2t.FS.mkdir('/working'); } catch (_) {}
  try { x2t.FS.mkdir('/working/fonts'); } catch (_) {}
  try { x2t.FS.mkdir('/working/pdf_tmp'); } catch (_) {}

  // 字体：放一个系统 ttf（跟浏览器 fetch vendor/pdf-fonts 等价）
  const fontSrc = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf';
  if (fs.existsSync(fontSrc)) writeFS('/working/fonts/DejaVuSans.ttf', fs.readFileSync(fontSrc));
  else console.log('WARN: no dejavu font, PDF 字形可能缺失（不影响本测试）');

  const sizesMB = [0.5, 2, 6, 7.5, 9, 16, 24];
  for (const mb of sizesMB) {
    const bin = makeRenderBin(mb * 1024 * 1024);
    let status = 'ok';
    try {
      // 新路径：bin 先写 FS（堆内存），只传路径——绕开 ccall 'array' 的 stackAlloc
      writeFS('/working/pdf_tmp/render.bin', bin);
      const ret = x2t.ccall('x2t_render_bin_to_pdf_file', 'number',
        ['string', 'string', 'string'],
        ['/working/pdf_tmp/render.bin', '/working/fonts/', '/working/pdf_tmp/out.pdf']);
      let pdfLen = 0;
      try { pdfLen = x2t.FS.stat('/working/pdf_tmp/out.pdf').size; } catch (_) {}
      status = `ret=${ret} pdf=${pdfLen}B`;
    } catch (e) {
      status = 'THROWN: ' + (e && e.message ? e.message.split('\n')[0] : e);
    }
    console.log(`bin=${mb}MB (${bin.length}B) → ${status}`);
    if (x2t.ABORT) { console.log('wasm ABORT=true，后续调用不可行，停止'); break; }
  }
  process.exit(0);
};
