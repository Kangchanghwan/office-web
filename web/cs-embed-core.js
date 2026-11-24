// [cs] cbEdit v1 순수 로직 (출처 검사, 메서드 분기, 형식 규칙). DOM 에 의존하지 않아 node 로 시험한다.
(function (root) {
    'use strict';

    var EXACT_ORIGINS = ['https://clipboardapp.org', 'https://www.clipboardapp.org'];
    var VERCEL_RE = /^https:\/\/(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+vercel\.app$/;
    var LOCAL_RE = /^http:\/\/(?:localhost|127\.0\.0\.1):\d+$/;

    // frame-ancestors 목록과 같아야 한다: 'self' https://clipboardapp.org https://www.clipboardapp.org
    // https://*.vercel.app http://localhost:* http://127.0.0.1:*   ('self' 는 selfOrigin 으로 넘긴다)
    function isAllowedOrigin(origin, selfOrigin) {
        if (typeof origin !== 'string' || !origin || origin === 'null') return false;
        if (selfOrigin && origin === selfOrigin) return true;
        return EXACT_ORIGINS.indexOf(origin) >= 0 || VERCEL_RE.test(origin) || LOCAL_RE.test(origin);
    }

    var ACCEPTS = ['docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt', 'pdf'];
    var SAVE_EXT = { docx: 'docx', xlsx: 'xlsx', pptx: 'pptx', xls: 'xls', pdf: 'pdf', doc: 'docx', ppt: 'pptx' };
    var MIME = {
        docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        xls: 'application/vnd.ms-excel',
        pdf: 'application/pdf'
    };

    function extOf(name) {
        var m = /\.([^./\\]+)$/.exec(String(name || ''));
        return m ? m[1].toLowerCase() : '';
    }
    function isAccepted(ext) { return ACCEPTS.indexOf(ext) >= 0; }

    // 원래 확장자 -> 저장할 확장자 (doc->docx, ppt->pptx). 지원 밖이면 null.
    function exportExt(openExt) {
        openExt = String(openExt || '').toLowerCase();
        return Object.prototype.hasOwnProperty.call(SAVE_EXT, openExt) ? SAVE_EXT[openExt] : null;
    }
    // format 생략 = 원래 형식. 지정했다면 그 형식이 원래 형식의 저장 결과와 같을 때만 허용.
    function resolveExport(openExt, format) {
        var ext = exportExt(openExt);
        if (!ext) return { error: { code: 'UNSUPPORTED', message: 'unsupported source ext: ' + openExt } };
        if (format !== undefined && format !== null && format !== '') {
            var f = String(format).toLowerCase();
            if (f !== ext && f !== String(openExt).toLowerCase()) {
                return { error: { code: 'UNSUPPORTED', message: 'format not supported: ' + format } };
            }
        }
        return { ext: ext, mime: MIME[ext] };
    }

    function err(code, message) { return { ok: false, error: { code: code, message: message || code } }; }

    // 요청 한 건 처리. host: { open, export, markSaved, state, setTheme } (각각 Promise 또는 값 반환)
    // 돌려주는 값: Promise<{ok:true,result}|{ok:false,error}>
    function dispatch(host, req) {
        if (!req || req.cbEdit !== 1 || typeof req.id !== 'string' || typeof req.method !== 'string') return null;
        var args = (req.args && typeof req.args === 'object') ? req.args : {};
        var known = ['open', 'export', 'markSaved', 'state', 'setTheme'];
        var fn = known.indexOf(req.method) >= 0 ? host[req.method] : null;
        if (!fn) return Promise.resolve(err('UNKNOWN_METHOD', req.method));
        return new Promise(function (resolve) { resolve(fn.call(host, args)); }).then(
            function (result) { return { ok: true, result: result === undefined ? {} : result }; },
            function (e) {
                if (e && typeof e.code === 'string') return err(e.code, e.message);
                return err('INTERNAL', String((e && e.message) || e));
            });
    }

    var api = {
        isAllowedOrigin: isAllowedOrigin, extOf: extOf, isAccepted: isAccepted, ACCEPTS: ACCEPTS,
        exportExt: exportExt, resolveExport: resolveExport, dispatch: dispatch, MIME: MIME
    };
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    root.CbEmbedCore = api;
})(typeof window !== 'undefined' ? window : globalThis);
