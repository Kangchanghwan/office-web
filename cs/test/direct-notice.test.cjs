// node --test cs/test/*.test.cjs
const test = require('node:test');
const assert = require('node:assert/strict');
const N = require('../../web/cs-direct-notice.js');

const mk = (search, top = true) => { const w = { location: { search } }; w.top = top ? w : {}; return w; };
const store = (v) => ({ getItem: () => v });

test('최상위 창이면 보인다', () => assert.equal(N.shouldShow(mk(''), store(null)), true));
test('iframe 안에서는 안 보인다', () => assert.equal(N.shouldShow(mk('', false), store(null)), false));
test('embed=cs 이면 안 보인다', () => assert.equal(N.shouldShow(mk('?embed=cs'), store(null)), false));
test('닫은 기록이 있으면 안 보인다', () => assert.equal(N.shouldShow(mk(''), store('1')), false));
test('저장소가 막혀도 보인다', () => assert.equal(N.shouldShow(mk(''), { getItem() { throw new Error('x'); } }), true));
test('문구·링크·높이', () => {
  assert.equal(N.TEXT, 'Clipboard Share로 돌아가기');
  assert.equal(N.HREF, 'https://www.clipboardapp.org/');
  assert.ok(N.HEIGHT <= 28);
});
