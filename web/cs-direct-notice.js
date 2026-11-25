// 주소창으로 직접 들어온 사람에게 Clipboard Share 로 돌아가는 링크를 보여 주는 얇은 한 줄(높이 28px).
// 최상위 창이고 embed=cs 가 아닐 때만 뜨고, 닫으면 localStorage 에 기억한다. 외부 리소스·리다이렉트 없음.
(function (root) {
  'use strict';
  var HEIGHT = 28, KEY = 'cs-direct-notice-dismissed';
  var TEXT = 'Clipboard Share로 돌아가기', HREF = 'https://www.clipboardapp.org/';

  function shouldShow(win, storage) {
    var top = false;
    try { top = win.top === win; } catch (e) { top = false; }
    if (!top) return false;
    if (new URLSearchParams(win.location.search).get('embed') === 'cs') return false;
    try { if (storage && storage.getItem(KEY) === '1') return false; } catch (e) { /* 저장소 막힘: 그냥 보여 줌 */ }
    return true;
  }

  function install(win, doc, storage) {
    if (storage === undefined) { try { storage = win.localStorage; } catch (e) { storage = null; } }
    if (!shouldShow(win, storage)) return null;
    var style = doc.createElement('style');
    style.textContent = 'body{padding-bottom:' + HEIGHT + 'px}' +
      '#cs-direct-notice{position:fixed;left:0;right:0;bottom:0;height:' + HEIGHT + 'px;z-index:2147483000;display:flex;align-items:center;justify-content:center;gap:10px;box-sizing:border-box;padding:0 12px;font:12px/1 system-ui,-apple-system,sans-serif;background:#eef2f7;color:#334;border-top:1px solid #d5dbe3;white-space:nowrap;overflow:hidden}' +
      '#cs-direct-notice a{color:#2b6cb0;text-decoration:underline}' +
      '#cs-direct-notice button{border:0;background:transparent;color:#667;font-size:16px;line-height:1;cursor:pointer;padding:0 4px}';
    var bar = doc.createElement('div');
    bar.id = 'cs-direct-notice';
    bar.setAttribute('role', 'note');
    var a = doc.createElement('a');
    a.href = HREF; a.target = '_blank'; a.rel = 'noopener'; a.textContent = TEXT;
    var x = doc.createElement('button');
    x.type = 'button'; x.setAttribute('aria-label', '닫기'); x.textContent = '×';
    function close() {
      try { if (storage) storage.setItem(KEY, '1'); } catch (e) { /* 무시 */ }
      bar.remove(); style.remove();
    }
    x.addEventListener('click', close);
    bar.appendChild(a); bar.appendChild(x);
    function mount() { doc.head.appendChild(style); doc.body.appendChild(bar); }
    if (doc.readyState === 'loading') doc.addEventListener('DOMContentLoaded', mount, { once: true }); else mount();
    return { bar: bar, style: style, close: close };
  }

  var api = { HEIGHT: HEIGHT, KEY: KEY, TEXT: TEXT, HREF: HREF, shouldShow: shouldShow, install: install };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else { root.CsDirectNotice = api; install(root, root.document); }
})(typeof window !== 'undefined' ? window : globalThis);
