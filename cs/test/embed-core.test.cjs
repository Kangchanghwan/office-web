// node --test cs/test/
const test = require('node:test');
const assert = require('node:assert/strict');
const C = require('../../web/cs-embed-core.js');

test('출처 허용 목록', () => {
  for (const o of ['https://clipboardapp.org', 'https://www.clipboardapp.org', 'https://x.vercel.app',
    'https://a-b.c.vercel.app', 'http://localhost:3000', 'http://127.0.0.1:8080', 'https://office.clipboardapp.org'])
    assert.equal(C.isAllowedOrigin(o, 'https://office.clipboardapp.org'), true, o);
  for (const o of ['http://clipboardapp.org', 'https://evil.com', 'https://clipboardapp.org.evil.com',
    'https://vercel.app', 'https://evilvercel.app', 'https://x.vercel.app.evil.com', 'http://localhost',
    'http://localhost:80@evil.com', 'https://localhost:3000', 'null', '', undefined, 'https://sub.clipboardapp.org'])
    assert.equal(C.isAllowedOrigin(o, 'https://office.clipboardapp.org'), false, String(o));
});

test('형식 규칙', () => {
  for (const [a, b] of [['docx','docx'],['xlsx','xlsx'],['pptx','pptx'],['xls','xls'],['pdf','pdf'],['doc','docx'],['ppt','pptx']])
    assert.equal(C.resolveExport(a).ext, b);
  assert.equal(C.resolveExport('doc', 'docx').ext, 'docx');
  assert.equal(C.resolveExport('docx', 'pdf').error.code, 'UNSUPPORTED');
  assert.equal(C.resolveExport('hwp').error.code, 'UNSUPPORTED');
  assert.equal(C.extOf('A.B.DOCX'), 'docx');
  assert.equal(C.isAccepted('hwp'), false);
});

test('메서드 분기', async () => {
  const host = { state: () => ({ loaded: false }), open: () => { throw { code: 'BUSY', message: 'b' }; },
    export: () => { throw new Error('x'); }, markSaved: () => undefined };
  assert.deepEqual(await C.dispatch(host, { cbEdit: 1, id: '1', method: 'state' }), { ok: true, result: { loaded: false } });
  assert.deepEqual(await C.dispatch(host, { cbEdit: 1, id: '2', method: 'markSaved' }), { ok: true, result: {} });
  assert.equal((await C.dispatch(host, { cbEdit: 1, id: '3', method: 'open', args: {} })).error.code, 'BUSY');
  assert.equal((await C.dispatch(host, { cbEdit: 1, id: '4', method: 'export' })).error.code, 'INTERNAL');
  assert.equal((await C.dispatch(host, { cbEdit: 1, id: '5', method: 'nope' })).error.code, 'UNKNOWN_METHOD');
  assert.equal((await C.dispatch(host, { cbEdit: 1, id: '6', method: 'setTheme' })).error.code, 'UNKNOWN_METHOD');
  assert.equal(C.dispatch(host, { id: '7', method: 'state' }), null);
  assert.equal(C.dispatch(host, { cbEdit: 1, method: 'state' }), null);
  assert.equal(C.dispatch(host, { cbEdit: 1, id: '8', method: 'toString' }).then ? 1 : 0, 1);
  assert.equal((await C.dispatch(host, { cbEdit: 1, id: '9', method: 'toString' })).error.code, 'UNKNOWN_METHOD');
});
