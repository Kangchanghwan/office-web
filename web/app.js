/*
 * app.js — standalone in-browser office editor.
 *
 * Opens, edits and saves Office documents entirely client-side using the
 * self-compiled x2t.wasm converter. No Document Server is required.
 *
 *   open : file bytes -> x2t (docx|xlsx|pptx -> bin) -> editor.openDocument({buffer})
 *   save : editor bin -> x2t (bin -> docx|xlsx|pptx) -> download / IndexedDB
 *   pdf  : raw PDF bytes -> pdfeditor (native format); save is two-stage —
 *          base PDF + compiled changes stream merged by x2t (m_bFromChanges)
 *
 * Files are persisted in IndexedDB (see store.js) so they survive reloads.
 */
(function () {
    'use strict';

    // Bump on every deploy: keys the Cache API entry for x2t.wasm, so users
    // re-download the converter exactly once per release.
    var BUILD_VERSION = '10.0.0-cs1';
    window.__OO_BUILD_VERSION = BUILD_VERSION;

    // ---------- Service Worker cache version guard ----------
    // The editor iframe's assets (notably offline-shim.js) are cached by the
    // Service Worker with a cache-first strategy keyed on g_buildVersion
    // (document_editor_service_worker.js). After a release, browsers that still
    // run the old SW keep serving stale files — a hard refresh bypasses only
    // the HTTP cache, not the SW cache. The entry pages (index/docx/...html)
    // are served no-cache so they are always fresh: remember the last seen
    // build version here, and when BUILD_VERSION changes delete the SW's
    // document_editor_* caches, then reload once (guarded against loops).
    function ensureFreshSwCache() {
        var KEY = '__oo_build_seen';
        var RELOAD_KEY = '__oo_build_reloaded';
        var seen = null;
        try { seen = localStorage.getItem(KEY); } catch (e) {}
        if (seen === BUILD_VERSION) return Promise.resolve(false);   // unchanged

        var deleteCaches = Promise.resolve();
        if (window.caches && caches.keys) {
            deleteCaches = caches.keys().then(function (keys) {
                return Promise.all(keys
                    .filter(function (k) { return k.indexOf('document_editor_') === 0; })
                    .map(function (k) { return caches.delete(k); }));
            }).catch(function () {});
        }
        return deleteCaches.then(function () {
            // The current navigation may still be controlled by the old SW (the
            // new one takes over after install/activate), so reload once to
            // guarantee the user runs the new code. RELOAD_KEY makes it a
            // one-shot per version. Anti-loop: only reload when the marker was
            // actually written — in private mode / full storage setItem throws,
            // and an unconditional reload would loop forever; without the
            // marker we just clear caches and let the next navigation pick up
            // the fresh files.
            var alreadyReloaded = null;
            try { alreadyReloaded = localStorage.getItem(RELOAD_KEY); } catch (e) {}
            try { localStorage.setItem(KEY, BUILD_VERSION); } catch (e) {}
            if (alreadyReloaded !== BUILD_VERSION) {
                var markerWritten = false;
                try { localStorage.setItem(RELOAD_KEY, BUILD_VERSION); markerWritten = true; } catch (e) {}
                if (markerWritten) {
                    window.location.reload();
                    return true;   // reload triggered: abort boot
                }
            }
            return false;
        }).catch(function () { return false; });
    }

    // ---------- config from URL ----------
    var PARAMS = {};
    new URLSearchParams(window.location.search).forEach(function (v, k) { PARAMS[k] = v; });

    // ---------- i18n ----------
    // The editors ship 40+ locale packs (vendor/web-apps/apps/*/main/locale/);
    // editorConfig.lang picks one. The app shell itself is translated for the
    // languages in SHELL_I18N below (others fall back to English).
    var LOCALES = ['ar', 'az', 'be', 'bg', 'ca', 'cs', 'da', 'de', 'el', 'en',
        'es', 'eu', 'fi', 'fr', 'gl', 'he', 'hr', 'hu', 'hy', 'id', 'it', 'ja',
        'ko', 'lo', 'lv', 'ms', 'nl', 'no', 'pl', 'pt', 'pt-pt', 'ro', 'ru',
        'si', 'sk', 'sl', 'sq', 'sr-cyrl', 'sr', 'sv', 'tr', 'uk', 'ur', 'vi',
        'zh', 'zh-tw'];
    function detectEditorLang() {
        var nav = (navigator.language || 'en').toLowerCase();
        if (nav.indexOf('zh') === 0) {
            return (nav.indexOf('tw') >= 0 || nav.indexOf('hk') >= 0 ||
                    nav.indexOf('mo') >= 0 || nav.indexOf('hant') >= 0) ? 'zh-tw' : 'zh';
        }
        if (nav.indexOf('pt-pt') === 0) return 'pt-pt';
        if (nav.indexOf('sr-cyrl') === 0) return 'sr-cyrl';
        var primary = nav.split('-')[0];
        for (var i = 0; i < LOCALES.length; i++) {
            if (LOCALES[i] === primary) return primary;
        }
        return 'en';
    }
    var UI_LANG = detectEditorLang();
    // Shell UI translations, keyed by the English source string. %s are
    // positional substitutions. Missing keys fall back to English.
    var SHELL_I18N = {
        'zh': {
            'Open, edit and save Office documents entirely in your browser.': '在浏览器中打开、编辑和保存 Office 文档。',
            '📂 Open a file': '📂 打开文件',
            '＋ Document': '＋ 文档',
            '＋ Spreadsheet': '＋ 表格',
            '＋ Presentation': '＋ 演示',
            'Delete all files': '删除全部文件',
            '← Files': '← 文件',
            '💾 Save': '💾 保存',
            'Back to file library': '返回文件库',
            'Save to file library (Ctrl+S)': '保存到文件库（Ctrl+S）',
            'Loading document …': '正在加载文档…',
            'Starting editor …': '正在启动编辑器…',
            'Converting document …': '正在转换文档…',
            'Saving …': '正在保存…',
            'Saved %s': '已保存 %s',
            'Editor not ready': '编辑器未就绪',
            'Open failed: %s': '打开失败：%s',
            'Save failed: %s': '保存失败：%s',
            'Save failed: No document data': '保存失败：没有文档数据',
            'Document render timeout. Please reload.': '文档渲染超时，请刷新页面。',
            'Editor error: %s': '编辑器错误：%s',
            'Download as %s is not supported': '不支持下载为 %s',
            'Generating %s …': '正在生成 %s…',
            'Exported %s': '已导出 %s',
            'Export failed: %s': '导出失败：%s',
            'Note: %s is a legacy format; saved as %s': '提示：%s 是旧格式，已保存为 %s',
            'Merge failed — downloaded the original PDF without the latest edits': '合并失败——已下载不含最新编辑的原始 PDF',
            'Saved the original PDF — this file could not be re-merged, so the latest edits are not included (%s)': '已保存原始 PDF——该文件无法重新合并，最新编辑未包含在内（%s）',
            'Name cannot be empty': '名称不能为空',
            'Renamed to %s': '已重命名为 %s',
            'Rename failed: %s': '重命名失败：%s',
            'Delete "%s"?': '删除"%s"？',
            'Rename file:': '重命名文件：',
            'Delete ALL files from the library? This cannot be undone.': '删除文件库中的所有文件？此操作无法撤销。',
            'New name:': '新名称：'
        }
    };
    function t(key) {
        var dict = SHELL_I18N[UI_LANG];
        var s = (dict && dict[key]) || key;
        for (var i = 1; i < arguments.length; i++) {
            s = s.replace('%s', String(arguments[i]));
        }
        return s;
    }
    // Translate the static shell markup (index.html) once the DOM is ready.
    function translateShell() {
        var dict = SHELL_I18N[UI_LANG];
        if (!dict) return;   // English (or untranslated language): keep markup as-is
        function setText(sel, key) {
            var el = document.querySelector(sel);
            if (el) el.textContent = dict[key] || el.textContent;
        }
        function setTitle(sel, key) {
            var el = document.querySelector(sel);
            if (el && dict[key]) el.setAttribute('title', dict[key]);
        }
        // [cs] subtitle is Korean static text in index.html
        setText('#upload-btn', '📂 Open a file');
        setText('[data-new="docx"]', '＋ Document');
        setText('[data-new="xlsx"]', '＋ Spreadsheet');
        setText('[data-new="pptx"]', '＋ Presentation');
        var clearBtn = document.querySelector('#clear-all-btn');
        if (clearBtn && dict['Delete all files']) clearBtn.textContent = '🗑 ' + dict['Delete all files'];
        setTitle('#clear-all-btn', 'Delete all files');
        setText('#back-home', '← Files'); setTitle('#back-home', 'Back to file library');
        setText('#save-btn', '💾 Save'); setTitle('#save-btn', 'Save to file library (Ctrl+S)');
    }

    // Determine the entry type. Priority: explicit global (?set by an embedder),
    // ?type= param, then the page filename (docx.html -> 'docx', etc.).
    function typeFromPath() {
        var m = /\/(docx|xlsx|pptx|pdf)\.html$/.exec(window.location.pathname);
        return m ? m[1] : '';
    }
    var PAGE_TYPE = window.__OO_PAGE_TYPE || PARAMS.type || typeFromPath();

    var statusEl = document.getElementById('status');
    var titlebarStatusEl = document.getElementById('titlebar-status');
    function setStatus(msg, isError) {
        // The floating chip is for the home screen; while a document is open
        // the same message is mirrored into the title bar.
        if (statusEl) {
            if (!msg) { statusEl.textContent = ''; statusEl.classList.remove('show'); }
            else {
                statusEl.textContent = msg;
                statusEl.classList.add('show');
                statusEl.classList.toggle('error', !!isError);
            }
        }
        if (titlebarStatusEl) {
            titlebarStatusEl.textContent = msg || '';
            titlebarStatusEl.classList.toggle('error', !!isError && !!msg);
        }
    }

    // ---------- x2t.wasm conversion ----------
    function fsWriteFile(mod, path, bytes) {
        var stream = mod.FS.open(path, 'w');
        mod.FS.write(stream, bytes, 0, bytes.length, 0);
        mod.FS.close(stream);
    }

    function resetWorking(mod) {
        try {
            var rmrf = function (p) {
                var st;
                try { st = mod.FS.stat(p); } catch (_) { return; }
                if (mod.FS.isDir(st.mode)) {
                    mod.FS.readdir(p).forEach(function (e) {
                        if (e === '.' || e === '..') return;
                        rmrf(p + '/' + e);
                    });
                    mod.FS.rmdir(p);
                } else {
                    mod.FS.unlink(p);
                }
            };
            rmrf('/working');
        } catch (_) {}
        mod.FS.mkdir('/working');
        mod.FS.mkdir('/working/fonts');
        mod.FS.mkdir('/working/themes');
    }

    function x2tConvert(mod, inputBytes, fromExt, toExt, workName) {
        // Only reset the working dir when converting TO bin (the open
        // direction). With m_bIsNoBase64=false, x2t extracts the OOXML's
        // images as files next to the bin (/working/media/<name>) and the bin
        // references them by relative path; converting FROM bin (save/download)
        // reads those same files back to embed the images into the output.
        // Wiping /working before a save would therefore strip every image from
        // the saved file. (The reference office.html simply never cleans
        // /working; here the open-direction reset also clears the previous
        // document's leftovers, which matters in the multi-document library.)
        if (toExt === 'bin') resetWorking(mod);
        var inName = workName + '.' + fromExt;
        var outName = workName + '.' + toExt;
        fsWriteFile(mod, '/working/' + inName, inputBytes);
        // m_bIsNoBase64 controls whether images are embedded into the bin or
        // referenced as external paths:
        //   - converting TO bin (open direction) must be false: embed the image
        //     bytes into the bin so the editor can render them directly. With
        //     true, the bin only holds references like `_offline_media/imageN.png`
        //     and the editor would fetch that (nonexistent) directory -> 404,
        //     no images.
        //   - converting FROM bin (save/download direction) also stays false:
        //     the bin already embeds the images and they are written back into
        //     word/media/ etc., producing a self-contained file.
        var noBase64 = false;
        var xml = '<?xml version="1.0" encoding="utf-8"?>' +
            '<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">' +
            '<m_sFontDir>/working/fonts/</m_sFontDir>' +
            '<m_sThemeDir>/working/themes</m_sThemeDir>' +
            '<m_sFileFrom>/working/' + inName + '</m_sFileFrom>' +
            '<m_sFileTo>/working/' + outName + '</m_sFileTo>' +
            '<m_bIsNoBase64>' + noBase64 + '</m_bIsNoBase64>' +
            '<m_nCsvTxtEncoding>46</m_nCsvTxtEncoding>' +
            '<m_nCsvDelimiter>4</m_nCsvDelimiter>' +
            '</TaskQueueDataConvert>';
        mod.FS.writeFile('/working/params.xml', xml);
        var ret = mod.ccall('main1', 'number', ['string'], ['/working/params.xml']);
        if (ret !== 0) throw new Error('x2t conversion failed ret=' + ret);
        return mod.FS.readFile('/working/' + outName);
    }

    // ---------- OOXML media extraction (kills offline _offline_media 404s) ----------
    // When opening a docx/xlsx/pptx offline, the editor asks for images by name
    // (e.g. image1.png) under `_offline_media/` — a directory that does not
    // exist -> 404. This extracts the */media/* images from the source OOXML
    // into data: URLs up front, exposed as window.__ooMediaMap (exact filename
    // -> dataURL) and window.__ooMediaMapByBase (extensionless basename ->
    // {ext: dataURL} variant table); the iframe offline-shims use them to remap
    // g_oDocumentUrls' `_offline_media/<name>` entries to data: URLs so the
    // editor gets its images locally without network requests.
    //
    // Why the basename index: x2t changes image extensions when converting
    // pptx->bin (source image1.jpeg, editor requests image1.jpg; SmartArt
    // png/svg are swapped too), and pptx->bin does NOT embed media into the
    // bin like docx does — pptx images rely entirely on this remap path.
    // Exact-name matching fails on the extension change, hence the fallback.
    // Minimal ZIP parsing (central directory only), no extra libraries.
    function extractOoxmlMedia(bytes) {
        var exact = {};   // exact filename -> dataURL
        var base = {};    // extensionless basename -> { ext: dataURL } variants
        try {
            var dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
            // find the EOCD (End Of Central Directory, signature 0x06054b50) from the tail
            var eocd = -1;
            for (var i = bytes.length - 22; i >= 0 && i >= bytes.length - 22 - 65536; i--) {
                if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
            }
            if (eocd < 0) return Promise.resolve({ exact: exact, base: base });
            var nEntries = dv.getUint16(eocd + 10, true);
            var off = dv.getUint32(eocd + 16, true);   // central directory offset
            var inflate = function (u8) {
                var ds = new DecompressionStream('deflate-raw');
                var stream = new Blob([u8]).stream().pipeThrough(ds);
                return new Response(stream).arrayBuffer().then(function (buf) {
                    return new Uint8Array(buf);
                });
            };
            var jobs = [];
            var decoder = new TextDecoder();
            for (var e = 0; e < nEntries; e++) {
                if (dv.getUint32(off, true) !== 0x02014b50) break;
                var method = dv.getUint16(off + 10, true);
                var cSize = dv.getUint32(off + 20, true);
                var nameLen = dv.getUint16(off + 28, true);
                var extraLen = dv.getUint16(off + 30, true);
                var cmtLen = dv.getUint16(off + 32, true);
                var lhOff = dv.getUint32(off + 42, true);
                var name = decoder.decode(bytes.subarray(off + 46, off + 46 + nameLen));
                off += 46 + nameLen + extraLen + cmtLen;
                // only images under */media/* (skip videos and other large media)
                var m = /(?:^|\/)media\/([^/]+)$/.exec(name);
                if (!m) continue;
                var fname = m[1];
                var ext = (fname.split('.').pop() || '').toLowerCase();
                if (ext === 'mp4' || ext === 'webm' || ext === 'mp3' || ext === 'wav') continue;
                // read the local file header for the real data offset
                if (dv.getUint32(lhOff, true) !== 0x04034b50) continue;
                var lhNameLen = dv.getUint16(lhOff + 26, true);
                var lhExtraLen = dv.getUint16(lhOff + 28, true);
                var dataStart = lhOff + 30 + lhNameLen + lhExtraLen;
                var raw = bytes.subarray(dataStart, dataStart + cSize);
                (function (fname, ext, method, raw) {
                    jobs.push(Promise.resolve().then(function () {
                        if (method === 8) return inflate(raw);        // deflate
                        if (method === 0) return raw;                 // stored
                        return null;
                    }).catch(function () { return null; }).then(function (data) {
                        if (!data) return;
                        var mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
                                     gif: 'image/gif', svg: 'image/svg+xml', bmp: 'image/bmp',
                                     webp: 'image/webp', emf: 'image/x-emf', wmf: 'image/x-wmf',
                                     tiff: 'image/tiff', tif: 'image/tiff' }[ext] || 'application/octet-stream';
                        var bin = '';
                        for (var j = 0; j < data.length; j += 0x8000) {
                            bin += String.fromCharCode.apply(null, data.subarray(j, j + 0x8000));
                        }
                        var dataUrl = 'data:' + mime + ';base64,' + btoa(bin);
                        exact[fname] = dataUrl;
                        // basename table: same name with different extensions (e.g.
                        // image1.png + image1.svg, Word's emf/png fallback pairs) keeps
                        // every variant keyed by extension; the remap prefers the variant
                        // matching the requested extension so the right image wins.
                        var b = fname.lastIndexOf('.') > 0 ? fname.slice(0, fname.lastIndexOf('.')) : fname;
                        if (!(b in base)) base[b] = {};
                        if (!(ext in base[b])) base[b][ext] = dataUrl;
                    }));
                })(fname, ext, method, raw);
            }
            return Promise.all(jobs).then(function () { return { exact: exact, base: base }; });
        } catch (e) { /* extraction failure must not block opening */ }
        return Promise.resolve({ exact: exact, base: base });
    }

    // ---------- PDF save: base PDF + compiled changes -> merged PDF ----------
    // The PDF editor's save is two-stage (isomorphic to the Document Server's
    // collaborative save):
    //   base    = the original PDF bytes as opened
    //             (shim __ooGetFileData -> DocumentRenderer.file.getFileBinary())
    //   changes = incremental serialization of the edits
    //             (shim __ooGetPdfChanges -> DocumentRenderer.Save(),
    //             "%PDF"+command stream; only a 26-byte header when empty)
    // Officially the server-side C++ CPdfFile::AddToPdfFromBinary merges
    // changes into base; here we reuse the same code from the self-compiled
    // x2t.wasm: main1 + <m_bFromChanges>true</m_bFromChanges>, with the
    // changes in a changes/ directory next to the source file (core
    // X2tConverter/src/lib/pdf_image.h fromCrossPlatform -> applyChangesPdf
    // -> applyCompiledChangesPdf convention).
    // With no edits (changes null or <= 26-byte header) the base is returned
    // directly, skipping wasm.
    // Fallback fonts for the PDF merge. Text inserted in PDF edit mode may
    // reference glyphs missing from the page's subset fonts; the recompiled
    // x2t (GetFallbackFontForChar) resolves them from this font dir and
    // embeds the needed glyphs. Fetched lazily (only when merging real
    // changes) and cached in memory. NotoSansSC covers BOTH Latin and CJK
    // (the primary fallback for newly typed text); DroidSansFallbackFull
    // (CJK) and DejaVuSans (Latin) are secondary fallbacks. The extra faces
    // cover scripts Noto Sans SC lacks so typed Arabic/Devanagari/Korean/
    // Thai text embeds correctly instead of dropping glyphs:
    //   NotoNaskhArabic (Arabic), LohitDevanagari (Devanagari/Hindi),
    //   NanumGothic (Korean), FreeSerif (Thai + many others).
    var PDF_MERGE_FONTS = [
        // [cs] Latin text typed into a PDF is set in Arial; without a matching
        // face in /working/fonts x2t 10.0 maps the glyphs with an offset
        // (L -> N). Liberation Sans is the open, metric-compatible Arial face
        // of the font catalog (decoded-fonts index: 062 regular, 059 bold).
        { url: 'vendor/decoded-fonts/062.ttf', name: 'LiberationSans-Regular.ttf' },
        { url: 'vendor/decoded-fonts/059.ttf', name: 'LiberationSans-Bold.ttf' },
        { url: 'vendor/pdf-fonts/NotoSansSC.ttf', name: 'NotoSansSC.ttf' },
        { url: 'vendor/pdf-fonts/DroidSansFallbackFull.ttf', name: 'DroidSansFallbackFull.ttf' },
        { url: 'vendor/pdf-fonts/DejaVuSans.ttf', name: 'DejaVuSans.ttf' },
        { url: 'vendor/pdf-fonts/NotoNaskhArabic.ttf', name: 'NotoNaskhArabic.ttf' },
        { url: 'vendor/pdf-fonts/LohitDevanagari.ttf', name: 'LohitDevanagari.ttf' },
        { url: 'vendor/pdf-fonts/NanumGothic.ttf', name: 'NanumGothic.ttf' },
        { url: 'vendor/pdf-fonts/FreeSerif.ttf', name: 'FreeSerif.ttf' }
    ];
    var __pdfMergeFontsPromise = null;
    function loadPdfMergeFonts() {
        if (__pdfMergeFontsPromise) return __pdfMergeFontsPromise;
        __pdfMergeFontsPromise = Promise.all(PDF_MERGE_FONTS.map(function (f) {
            return fetch(f.url).then(function (r) {
                if (!r.ok) return null;
                return r.arrayBuffer().then(function (buf) { return { name: f.name, data: new Uint8Array(buf) }; });
            }).catch(function () { return null; });
        })).then(function (list) { return list.filter(Boolean); });
        return __pdfMergeFontsPromise;
    }

    function mergePdfChanges(base, changes) {
        var EMPTY_CHANGES_LEN = 26;   // Save() with no edits: "%PDF"+length header only
        if (!changes || changes.length <= EMPTY_CHANGES_LEN) return Promise.resolve(base);
        return Promise.all([window.__x2tReadyPromise, loadPdfMergeFonts()]).then(function (ready) {
            var mod = ready[0];
            var fonts = ready[1];
            var FS = mod.FS;
            try { FS.mkdir('/working'); } catch (e) {}
            try { FS.mkdir('/working/changes'); } catch (e) {}
            try { FS.mkdir('/working/fonts'); } catch (e) {}
            try { FS.mkdir('/working/themes'); } catch (e) {}
            // clean leftovers from a previous round (save may fire back-to-back)
            ['/working/changes', '/working/fonts'].forEach(function (dir) {
                try {
                    FS.readdir(dir).filter(function (n) { return n !== '.' && n !== '..'; })
                        .forEach(function (e) { try { FS.unlink(dir + '/' + e); } catch (_) {} });
                } catch (_) {}
            });
            ['src.pdf', 'out.pdf', 'params.xml'].forEach(function (f) {
                try { FS.unlink('/working/' + f); } catch (_) {}
            });
            // supply fallback fonts so the merge can embed glyphs missing from
            // the page's subset fonts (inserted text)
            fonts.forEach(function (f) { fsWriteFile(mod, '/working/fonts/' + f.name, f.data); });
            fsWriteFile(mod, '/working/src.pdf', base);
            fsWriteFile(mod, '/working/changes/changes0.bin', changes);
            var xml = '<?xml version="1.0" encoding="utf-8"?>' +
                '<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">' +
                '<m_sFontDir>/working/fonts/</m_sFontDir>' +
                '<m_sThemeDir>/working/themes</m_sThemeDir>' +
                '<m_sFileFrom>/working/src.pdf</m_sFileFrom>' +
                '<m_sFileTo>/working/out.pdf</m_sFileTo>' +
                '<m_bIsNoBase64>true</m_bIsNoBase64>' +
                '<m_bFromChanges>true</m_bFromChanges>' +
                '</TaskQueueDataConvert>';
            FS.writeFile('/working/params.xml', xml);
            var ret = mod.ccall('main1', 'number', ['string'], ['/working/params.xml']);
            if (ret !== 0) throw new Error('PDF changes merge failed ret=' + ret);
            return FS.readFile('/working/out.pdf');
        });
    }

    // ---------- PDF export from the editors: render bin -> wasm CPdfFile -> PDF ----------
    // "Download As PDF" on docx/xlsx/pptx sends a *render bin* from the iframe
    // (sdkjs asc_nativeGetPDF output, the same drawing-command stream the
    // Document Server feeds to CPdfFile). Reuse the same C++ code from the
    // self-compiled x2t.wasm via the x2t_render_bin_to_pdf_file export.
    // From the render bin, extract the font names used by text (the
    // ctFontName=41 command: [41][len:2][utf16 chars]) so the PDF embeds the
    // same faces the page renders with, instead of a Helvetica fallback.
    function parseRenderBinFonts(bin) {
        var fonts = {};
        var names = [];
        // DataView to tolerate unaligned utf16 (ctFontName payload can start
        // at an odd offset; Uint16Array would throw)
        var dv = bin.buffer && bin.byteOffset !== undefined
            ? new DataView(bin.buffer, bin.byteOffset, bin.byteLength) : null;
        if (!dv) return names;
        var scanLen = Math.min(bin.length, 1024 * 1024);
        for (var i = 0; i < scanLen - 3; i++) {
            if (bin[i] === 0x29) {   // 41 = ctFontName
                var nChars = dv.getUint16(i + 1, true);
                var nBytes = nChars * 2;
                if (nBytes > 0 && nBytes < 400 && i + 3 + nBytes <= bin.length) {
                    var name = '';
                    for (var k = 0; k < nChars; k++) name += String.fromCharCode(dv.getUint16(i + 3 + k * 2, true));
                    if (/^[\x20-\x7E]+$/.test(name) && /[a-zA-Z]/.test(name) && !fonts[name]) {
                        fonts[name] = true;
                        names.push(name);
                    }
                    i += 2 + nBytes;   // skip what was just read
                }
            }
        }
        return names;
    }

    // Fetch the standard ttf/otf the render bin needs from vendor/decoded-fonts/
    // (the XOR-decoded font pool; name -> file mapping in fonts-index.json).
    // Missing names are skipped — the wasm side falls back to the faces staged
    // so far (see PDF_MERGE_FONTS for the always-supplied pool).
    var __fontsIndexPromise = null;
    function getFontsIndex() {
        if (!__fontsIndexPromise) {
            __fontsIndexPromise = fetch('vendor/decoded-fonts/fonts-index.json')
                .then(function (r) { return r.ok ? r.json() : {}; })
                .catch(function () { return {}; });
        }
        return __fontsIndexPromise;
    }

    function x2tRenderBinToPdf(mod, renderBin) {
        var FS = mod.FS;
        try { FS.mkdir('/working'); } catch (e) {}
        try { FS.mkdir('/working/fonts'); } catch (e) {}
        try { FS.mkdir('/working/pdf_tmp'); } catch (e) {}
        var fontsDir = '/working/fonts';
        var getFont = function (name) {
            return fetch('vendor/decoded-fonts/' + name)
                .then(function (r) { return r.ok ? r.arrayBuffer() : null; })
                .then(function (buf) { return buf ? new Uint8Array(buf) : null; })
                .catch(function () { return null; });
        };

        // 1) font names used by the render bin -> files from fonts-index.json
        return getFontsIndex().then(function (index) {
            var needed = parseRenderBinFonts(renderBin);
            var wanted = {};   // index file -> {outName, ext}
            needed.forEach(function (name) {
                var entries = index[name.toLowerCase()];
                if (!entries || !entries.length) return;
                // prefer the Regular-weight entry (full == family)
                var entry = entries[0];
                for (var j = 0; j < entries.length; j++) {
                    if (entries[j].full === name) { entry = entries[j]; break; }
                }
                if (wanted[entry.file]) return;
                var ext = entry.kind === 'otf' ? 'otf' : (entry.kind === 'ttc' ? 'ttc' : 'ttf');
                wanted[entry.file] = {
                    outName: name.replace(/[\s/\\]/g, '') + '.' + ext,
                    url: 'vendor/decoded-fonts/' + entry.file + '.' + ext
                };
            });
            var fetches = Object.keys(wanted).map(function (file) {
                return getFont(wanted[file].url.split('/').pop()).then(function (data) {
                    return { name: wanted[file].outName, data: data };
                });
            });
            return Promise.all(fetches);
        }).then(function (fetchedList) {
            // 2) stage the fonts: the decoded pool + the fixed merge pool (so
            // text whose family is not in the pool still embeds correctly)
            return loadPdfMergeFonts().then(function (mergeFonts) {
                fetchedList.forEach(function (f) {
                    if (f && f.data) fsWriteFile(mod, fontsDir + '/' + f.name, f.data);
                });
                mergeFonts.forEach(function (f) {
                    fsWriteFile(mod, fontsDir + '/' + f.name, f.data);
                });

                // 3) the render bin goes through the wasm FS, not ccall('array'):
                //    emscripten's 'array' argument stackAlloc's the whole bin on
                //    the 8MB wasm stack — complex decks' render bins exceed that
                //    and trap with "memory access out of bounds" (7.5MB passes,
                //    9MB crashes). Writing to FS puts it on the heap: no limit.
                FS.writeFile('/working/pdf_tmp/render.bin', renderBin);
                var ret = mod.ccall('x2t_render_bin_to_pdf_file', 'number',
                    ['string', 'string', 'string'],
                    ['/working/pdf_tmp/render.bin', fontsDir + '/', '/working/pdf_tmp/out.pdf']);
                if (ret !== 0) throw new Error('PDF conversion failed ret=' + ret);
                var pdf = FS.readFile('/working/pdf_tmp/out.pdf');
                try { FS.unlink('/working/pdf_tmp/render.bin'); } catch (e) {}
                try { FS.unlink('/working/pdf_tmp/out.pdf'); } catch (e) {}
                return pdf.slice();   // detach from the wasm heap view
            });
        });
    }

    // ---------- document type helpers ----------
    var CELL_EXTS = ['xls', 'xlsx', 'xlsm', 'ods', 'csv'];
    var SLIDE_EXTS = ['ppt', 'pptx', 'odp', 'ppsx'];
    var PDF_EXTS = ['pdf'];

    function detectDocumentType(ext) {
        ext = (ext || '').toLowerCase();
        if (CELL_EXTS.indexOf(ext) >= 0) return 'cell';
        if (SLIDE_EXTS.indexOf(ext) >= 0) return 'slide';
        if (PDF_EXTS.indexOf(ext) >= 0) return 'pdf';
        return 'word';
    }

    // bin header -> writable formats (measured against the compiled x2t)
    function detectBinType(bin) {
        if (!bin || bin.length < 4) return null;
        var head = String.fromCharCode(bin[0], bin[1], bin[2], bin[3]);
        if (head === 'DOCY') return ['docx', 'odt', 'rtf', 'txt'];
        if (head === 'XLSY') return ['xlsx', 'xls', 'ods'];
        if (head === 'PPTY') return ['pptx', 'odp'];
        return null;
    }
    var LEGACY_TO_NEW = { doc: 'docx', ppt: 'pptx', csv: 'xlsx' };

    // ---------- editor state ----------
    var docEditor = null;
    var pendingBin = null;
    var documentReady = false;
    var emptyDocGotReal = false;
    var currentFile = null;   // { id, name, ext }

    // ---------- autosave ----------
    // The editors report "has unsaved changes" through the same event that
    // drives their save-button state (asc_onDocumentCanSaveChanged). PDF is
    // additionally probed via the compiled changes stream length (26-byte
    // header == no edits), since its event path is less exercised offline.
    // While dirty, the document is persisted to the IndexedDB library every
    // AUTOSAVE_INTERVAL_MS, plus on tab-hide and before back-to-home, so a
    // crash/refresh loses at most one interval of work.
    var AUTOSAVE_INTERVAL_MS = 30000;
    var dirty = false;
    var editSeq = 0;               // bumped on every dirty signal (see doSave)
    var lastSavedChangesLen = 26;  // PDF changes-stream length at last save
    var autosaveTimer = null;

    function onEditorCanSaveChanged(canSave) {
        if (canSave) { dirty = true; editSeq++; }
        else dirty = false;
    }
    function editorDirtyNow() {
        if (!documentReady || !currentFile) return false;
        if (dirty) return true;
        if ((window.__ooDocExt || '').toLowerCase() === 'pdf') {
            var frame = document.querySelector('iframe[name="frameEditor"]');
            var w = frame && frame.contentWindow;
            if (w && typeof w.__ooGetPdfChanges === 'function') {
                try {
                    var ch = w.__ooGetPdfChanges();
                    if ((ch ? ch.length : 26) !== lastSavedChangesLen) return true;
                } catch (e) {}
            }
        }
        return false;
    }
    // After a successful persist, reset the editor's own modified flag too,
    // so its save button greys out exactly like the server-backed flow.
    function clearEditorModifiedFlag(frame) {
        try {
            var ed = frame.contentWindow.Asc && frame.contentWindow.Asc.editor;
            if (!ed) return;
            if (ed.SetUnchangedDocument) ed.SetUnchangedDocument();
            else if (ed.SetDocumentModified) ed.SetDocumentModified(false);
        } catch (e) {}
    }
    function autosaveTick() {
        if (saveInFlight || !editorDirtyNow()) return;
        saveDocument();
    }
    function startAutosave() {
        stopAutosave();
        autosaveTimer = setInterval(autosaveTick, AUTOSAVE_INTERVAL_MS);
    }
    function stopAutosave() {
        if (autosaveTimer) { clearInterval(autosaveTimer); autosaveTimer = null; }
        dirty = false;
        lastSavedChangesLen = 26;
    }
    function hookAutosave() {
        waitForEditorApi(function (ed) {
            try { ed.asc_registerCallback('asc_onDocumentCanSaveChanged', onEditorCanSaveChanged); } catch (e) {}
        });
        startAutosave();
    }
    // A persist succeeded. Clear the dirty flag unless the user kept editing
    // while the save was in flight (editSeq advanced), and sync the PDF
    // changes-stream baseline so the probe doesn't re-fire immediately.
    function markSaved(frame, seqAtStart, changesLen) {
        if (typeof changesLen === 'number') lastSavedChangesLen = changesLen;
        if (seqAtStart === editSeq) {
            dirty = false;
            if (frame) clearEditorModifiedFlag(frame);
        }
    }

    function waitDocumentRendered(docType) {
        var waited = 0;
        var timer = setInterval(function () {
            waited += 300;
            var ready = false;
            try {
                var frame = document.querySelector('iframe[name="frameEditor"]');
                var w = frame && frame.contentWindow;
                var ed = w && w.Asc && w.Asc.editor;
                if (!ed) return;
                if (docType === 'cell') {
                    var wb = ed.wbModel;
                    ready = !!(wb && wb.getWorksheetCount && wb.getWorksheetCount() > 0);
                } else if (docType === 'slide') {
                    var ldS = ed.WordControl && ed.WordControl.m_oLogicDocument;
                    ready = !!(ldS && ldS.Slides && ldS.Slides.length > 0);
                } else if (docType === 'pdf') {
                    // PDFEditorApi: the document model lives in
                    // DocumentRenderer.file; pages present => rendered.
                    var file = ed.DocumentRenderer && ed.DocumentRenderer.file;
                    ready = !!(file && file.pages && file.pages.length > 0);
                } else {
                    var ld = ed.WordControl && ed.WordControl.m_oLogicDocument;
                    ready = !!(ld && ld.Pages && ld.Pages.length > 0);
                }
            } catch (e) { /* iframe not ready */ }
            if (ready) {
                clearInterval(timer);
                documentReady = true;
                setStatus('');
                hookAutosave();
            } else if (waited > 60000) {
                clearInterval(timer);
                setStatus(t('Document render timeout. Please reload.'), true);
            }
        }, 300);
    }

    function createEditor(fileExt, title) {
        var docType = detectDocumentType(fileExt);
        var document = {
            fileType: fileExt,
            key: 'doc-' + Date.now(),
            title: title,
            url: '_offline_',
            permissions: { edit: true, download: true, print: true, copy: true }
        };
        if (docType === 'pdf') {
            // Explicit isForm: otherwise api.js first sends a checkParams
            // message making pdfeditor probe the "extended PDF" via the
            // downloadfile endpoint, which does not exist offline (wasted
            // round-trip). isForm=false puts &isForm=false on the iframe
            // URL and pdfeditor goes straight to startApp.
            document.isForm = false;
        }
        docEditor = new DocsAPI.DocEditor('editor', {
            type: 'desktop',
            width: '100%',
            height: '100%',
            documentType: docType,
            document: document,
            editorConfig: {
                mode: 'edit',
                lang: UI_LANG,
                canCoAuthoring: false,
                user: { id: 'local-user', name: 'Local User' },
                // [cs] spell-check dictionaries are not shipped; without them every
                // word is underlined, so start with spell checking off.
                customization: { about: false, feedback: false, autosave: false,
                                 features: { spellcheck: { mode: false, change: true } } }
            },
            events: {
                onAppReady: function () {},
                onDocumentReady: function () {
                    if (pendingBin && !emptyDocGotReal) {
                        var bin = pendingBin;
                        pendingBin = null;
                        setStatus(t('Loading document …'));
                        var buf = bin.slice();
                        docEditor.openDocument({ buffer: buf.buffer });
                    }
                },
                onError: function (e) { setStatus(t('Editor error: %s', JSON.stringify(e && e.data)), true); },
                onDownloadAs: function () {},
                onRequestClose: function () { window.close(); },
                onRequestSave: function () { saveDocument(); }
            }
        });
    }

    // ---------- open a file ----------
    function openFile(file) {
        // file: { id?, name, ext, data: Uint8Array }
        var ext = (file.ext || '').toLowerCase();
        currentFile = { id: file.id || null, name: file.name, ext: ext };
        document.title = currentFile.name;
        window.__ooDocExt = ext;
        hideHome();

        var docType = detectDocumentType(ext);
        if (docType === 'pdf') {
            // PDF skips the x2t bin round-trip: pdfeditor is a native-format
            // editor (sdkjs onEndLoadFile validates raw PDF bytes for
            // isPdfEditor(), not DOCY/XLSY/PPTY bin). The raw bytes go on
            // the bridge; the pdfeditor offline-shim's getEmpty() picks them
            // up during the one-shot load.
            // Note: opening does not wait for x2t.wasm either (it loads in
            // the background at boot; by the time a save merge needs it,
            // __x2tReadyPromise is ready or nearly so), saving the 36MB
            // wasm wait on the open path.
            window.__ooPendingDocBin = file.data;
            setStatus(t('Starting editor …'));
            createEditor(ext, currentFile.name);
            waitDocumentRendered(docType);
            if (window.OOEditorState && window.OOEditorState._start) {
                window.OOEditorState._start(ext, docType, currentFile.name);
            }
            return;
        }

        window.__x2tReadyPromise.then(function (mod) {
            setStatus(t('Converting document …'));
            // [cs] x2t 10.0 fails doc/ppt/xls -> bin directly (ret=88) for some
            // legacy files while legacy -> OOXML works; retry through OOXML.
            var bin;
            try {
                bin = x2tConvert(mod, file.data, ext, 'bin', 'open');
            } catch (e) {
                var via = { doc: 'docx', dot: 'docx', ppt: 'pptx', pps: 'pptx', xls: 'xlsx', xlt: 'xlsx' }[ext];
                if (!via) throw e;
                var ooxml = x2tConvert(mod, file.data, ext, via, 'legacy').slice();
                bin = x2tConvert(mod, ooxml, via, 'bin', 'open');
            }
            pendingBin = bin;
            window.__ooPendingDocBin = bin;
            // Extract the source OOXML's media (name -> data: URL) so the iframe
            // offline-shims can remap `_offline_media/<name>` to data: URLs and
            // kill the offline 404s. Awaited before createEditor: the editors
            // register image URLs very early during loading, so the maps must
            // be populated by then (failure/empty still opens fine — the bin
            // already embeds the image data for docx/xlsx).
            window.__ooMediaMap = {};
            window.__ooMediaMapByBase = {};
            return extractOoxmlMedia(file.data).then(function (media) {
                window.__ooMediaMap = media.exact;
                window.__ooMediaMapByBase = media.base;
            }).catch(function () {}).then(function () {
                setStatus(t('Starting editor …'));
                createEditor(ext, currentFile.name);
                waitDocumentRendered(docType);
                if (window.OOEditorState && window.OOEditorState._start) {
                    window.OOEditorState._start(ext, docType, currentFile.name);
                }
            });
        }).catch(function (e) {
            setStatus(t('Open failed: %s', (e && e.message || e)), true);
        });
    }

    function waitForEditorApi(cb) {
        var tries = 0;
        var timer = setInterval(function () {
            tries++;
            var frame = document.querySelector('iframe[name="frameEditor"]');
            var w = frame && frame.contentWindow;
            var ed = w && w.Asc && w.Asc.editor;
            if (ed) { clearInterval(timer); cb(ed); }
            else if (tries > 300) clearInterval(timer);
        }, 200);
    }

    // ---------- save ----------
    var saveInFlight = false;
    function saveDocument() {
        if (saveInFlight) return false;
        saveInFlight = true;
        doSave().then(function (ok) { saveInFlight = false; return ok; })
            .catch(function () { saveInFlight = false; });
        return true;
    }

    function doSave() {
        var frame = document.querySelector('iframe[name="frameEditor"]');
        if (!frame || !frame.contentWindow || typeof frame.contentWindow.__ooGetFileData !== 'function') {
            setStatus(t('Editor not ready'), true);
            return Promise.resolve(false);
        }
        setStatus(t('Saving …'));
        var openExt = (window.__ooDocExt || '').toLowerCase();
        if (openExt === 'pdf') {
            // Two-stage PDF save: bin = base PDF (__ooGetFileData), the
            // changes stream comes from __ooGetPdfChanges
            // (DocumentRenderer.Save()); x2t merges them into the final PDF.
            // With no edits the changes stream is a 26-byte header and
            // mergePdfChanges returns the base as-is.
            var base = frame.contentWindow.__ooGetFileData();
            if (!base || !base.length) {
                setStatus(t('Save failed: No document data'), true);
                return Promise.resolve(false);
            }
            var changes = frame.contentWindow.__ooGetPdfChanges
                ? frame.contentWindow.__ooGetPdfChanges() : null;
            var changesLen = changes ? changes.length : 26;
            var seqAtStart = editSeq;
            return mergePdfChanges(base, changes).then(function (merged) {
                return persistToLibrary(merged.slice(), 'pdf').then(function () {
                    markSaved(frame, seqAtStart, changesLen);
                    setStatus(t('Saved %s', new Date().toLocaleTimeString()));
                    setTimeout(function () { setStatus(''); }, 3000);
                    return true;
                });
            }).catch(function (e) {
                // The x2t merge could not re-process this particular PDF: its
                // C++ (xpdf) parser is stricter than the in-editor one, so for
                // some files it cannot re-open the base to apply the changes
                // (x2t returns 80 / AVS_FILEUTILS_ERROR_CONVERT and writes no
                // output). Fall back to persisting the untouched base PDF so
                // the document is never lost, and tell the user plainly that
                // the latest edits could not be merged into this file.
                return persistToLibrary(base.slice(), 'pdf').then(function () {
                    // Treat as saved for autosave purposes: the merge will keep
                    // failing for this file, so re-trying every interval would
                    // only spam the same warning. A manual save still re-tries.
                    markSaved(frame, seqAtStart, changesLen);
                    setStatus(t('Saved the original PDF — this file could not be re-merged, so the latest edits are not included (%s)', (e && e.message || e)), true);
                    setTimeout(function () { setStatus(''); }, 10000);
                    return true;
                });
            });
        }
        var seqAtStart2 = editSeq;
        return window.__x2tReadyPromise.then(function (mod) {
            var bin = frame.contentWindow.__ooGetFileData();
            if (!bin || !bin.length) throw new Error('No document data');

            var saveExt;
            var supported = detectBinType(bin);
            if (supported) {
                if (openExt && supported.indexOf(openExt) >= 0) {
                    saveExt = openExt;
                } else if (openExt && LEGACY_TO_NEW[openExt] && supported.indexOf(LEGACY_TO_NEW[openExt]) >= 0) {
                    saveExt = LEGACY_TO_NEW[openExt];
                    setStatus(t('Note: %s is a legacy format; saved as %s', openExt, saveExt));
                } else {
                    saveExt = supported[0];
                }
            } else {
                saveExt = openExt || 'docx';
            }
            var bytes = x2tConvert(mod, bin, 'bin', saveExt, 'save').slice();
            return persistToLibrary(bytes, saveExt);
        }).then(function () {
            markSaved(frame, seqAtStart2);
            setStatus(t('Saved %s', new Date().toLocaleTimeString()));
            setTimeout(function () { setStatus(''); }, 3000);
            return true;
        }).catch(function (e) {
            setStatus(t('Save failed: %s', (e && e.message || e)), true);
            return false;
        });
    }

    // Persist to the IndexedDB file library only (no browser download).
    // Saving keeps the document in the library; downloading is a separate
    // action available via File > Download As.
    function persistToLibrary(bytes, ext) {
        var base = currentFile && currentFile.name ? currentFile.name.replace(/\.[^.]+$/, '') : 'document';
        var name = base + '.' + ext;
        if (window.FileStore) {
            return window.FileStore.save({
                id: currentFile && currentFile.id,
                name: name,
                ext: ext,
                data: bytes
            }).then(function (rec) {
                currentFile.id = rec.id; currentFile.name = name; currentFile.ext = ext;
                document.title = name;
                var nameEl = document.getElementById('titlebar-name');
                if (nameEl) nameEl.textContent = name;
                refreshLibrary();
                return rec;
            });
        }
        return Promise.resolve();
    }

    // Trigger a browser download of the given bytes under the given filename.
    function triggerDownload(bytes, filename) {
        var blob = new Blob([bytes], { type: 'application/octet-stream' });
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 1000);
    }

    // ---------- offline download-as bridge ----------
    var OFFLINE_FORMAT_EXT = {
        65: 'docx', 67: 'odt', 68: 'rtf', 69: 'txt',
        257: 'xlsx', 258: 'xls', 259: 'ods', 260: 'csv',
        129: 'pptx', 131: 'odp',
        513: 'pdf'
    };
    var OFFLINE_FORMAT_NAMES = {
        65: 'DOCX', 66: 'DOC', 67: 'ODT', 68: 'RTF', 69: 'TXT', 70: 'HTML', 73: 'EPUB',
        257: 'XLSX', 258: 'XLS', 259: 'ODS', 260: 'CSV',
        129: 'PPTX', 130: 'PPT', 131: 'ODP',
        513: 'PDF', 515: 'DJVU', 516: 'XPS',
        1025: 'JPG', 1029: 'PNG'
    };

    function handleOfflineDownload(msg) {
        var ext = OFFLINE_FORMAT_EXT[msg.format];
        if (!ext) {
            var nm = OFFLINE_FORMAT_NAMES[msg.format] || ('format #' + msg.format);
            setStatus(t('Download as %s is not supported', nm), true);
            setTimeout(function () { setStatus(''); }, 4000);
            return;
        }
        setStatus(t('Generating %s …', ext));
        var bytesPromise;
        var pdfFallback = false;
        if (msg.format === 513 && msg.pdf) {
            // PDF editor "Download As PDF": base + compiled changes stream
            // merged by x2t (same path as saving). If the merge cannot
            // re-process this PDF (see doSave), download the untouched base
            // rather than failing outright.
            var pdfBase = new Uint8Array(msg.buffer);
            bytesPromise = mergePdfChanges(pdfBase,
                msg.changes ? new Uint8Array(msg.changes) : null)
                .then(function (merged) { return merged.slice(); })
                .catch(function () { pdfFallback = true; return pdfBase.slice(); });
        } else if (msg.format === 513) {
            // docx/xlsx/pptx "Download As PDF": the iframe sends the render
            // bin (sdkjs asc_nativeGetPDF output) -> wasm CPdfFile writes it.
            bytesPromise = window.__x2tReadyPromise
                .then(function (mod) { return x2tRenderBinToPdf(mod, new Uint8Array(msg.buffer)); });
        } else {
            bytesPromise = window.__x2tReadyPromise.then(function (mod) {
                var bin = new Uint8Array(msg.buffer);
                return x2tConvert(mod, bin, 'bin', ext, 'download').slice();
            });
        }
        bytesPromise.then(function (bytes) {
            // msg.title from asc_getDocumentName() already carries an extension
            // (e.g. "untitled.docx"); strip it before appending the target format,
            // otherwise Download As produces "untitled.docx.docx".
            var base = (msg.title || 'document').replace(/\.[^.]+$/, '');
            triggerDownload(bytes, base + '.' + ext);
            if (pdfFallback) {
                setStatus(t('Merge failed — downloaded the original PDF without the latest edits'), true);
                setTimeout(function () { setStatus(''); }, 8000);
            } else {
                setStatus(t('Exported %s', ext));
                setTimeout(function () { setStatus(''); }, 2500);
            }
        }).catch(function (e) {
            setStatus(t('Export failed: %s', (e && e.message || e)), true);
        });
    }

    // ---------- file library UI ----------
    function refreshLibrary() {
        var listEl = document.getElementById('file-list');
        var toolbarEl = document.getElementById('library-toolbar');
        if (!listEl || !window.FileStore) return;
        window.FileStore.list().then(function (files) {
            if (toolbarEl) {
                toolbarEl.style.display = files.length ? 'flex' : 'none';
                var countEl = toolbarEl.querySelector('.library-count');
                if (countEl) countEl.textContent = files.length + (files.length === 1 ? ' file' : ' files');
            }
            listEl.innerHTML = '';
            if (!files.length) {
                listEl.innerHTML = '<div class="empty-hint">No files yet. Open or create one.</div>';
                return;
            }
            files.forEach(function (f) {
                var item = document.createElement('div');
                item.className = 'file-item';
                item.setAttribute('data-id', f.id);
                var icon = { docx: '📄', xlsx: '📊', pptx: '📽', pdf: '📕' }[f.ext] || '📄';
                item.innerHTML =
                    '<span class="file-icon">' + icon + '</span>' +
                    '<span class="file-name"></span>' +
                    '<span class="file-meta">' + formatSize(f.size) + ' · ' + formatDate(f.updatedAt) + '</span>' +
                    '<button class="file-act file-rename" title="Rename">✏️</button>' +
                    '<button class="file-act file-download" title="Download">⬇</button>' +
                    '<button class="file-act file-del" title="Delete">×</button>';
                item.querySelector('.file-name').textContent = f.name;
                item.addEventListener('click', function (ev) {
                    if (ev.target.closest('.file-act')) return;
                    loadFromFileStore(f.id);
                });
                item.querySelector('.file-rename').addEventListener('click', function (ev) {
                    ev.stopPropagation();
                    renameFile(f);
                });
                item.querySelector('.file-download').addEventListener('click', function (ev) {
                    ev.stopPropagation();
                    downloadFile(f.id);
                });
                item.querySelector('.file-del').addEventListener('click', function (ev) {
                    ev.stopPropagation();
                    if (confirm(t('Delete "%s"?', f.name))) {
                        window.FileStore.remove(f.id).then(refreshLibrary);
                    }
                });
                listEl.appendChild(item);
            });
        });
    }

    function renameFile(f) {
        var input = prompt(t('Rename file:'), f.name);
        if (input === null) return;
        input = input.trim();
        if (!input) { setStatus(t('Name cannot be empty'), true); setTimeout(function () { setStatus(''); }, 3000); return; }
        if (input === f.name) return;
        // Keep the original extension: if the user typed a new one, accept it
        // only when it matches this editor's writable formats; otherwise the
        // old extension is re-appended.
        var dot = input.lastIndexOf('.');
        var base = dot > 0 ? input.slice(0, dot) : input;
        var ext = dot > 0 ? input.slice(dot + 1).toLowerCase() : f.ext;
        var writable = { docx: 1, xlsx: 1, pptx: 1, pdf: 1, doc: 1, xls: 1, ppt: 1, odt: 1, ods: 1, odp: 1, csv: 1, rtf: 1, txt: 1 };
        if (!writable[ext]) ext = f.ext;
        window.FileStore.rename(f.id, base + '.' + ext, ext).then(function () {
            setStatus(t('Renamed to %s', base + '.' + ext));
            setTimeout(function () { setStatus(''); }, 3000);
            refreshLibrary();
        }).catch(function (e) {
            setStatus('Rename failed: ' + (e && e.message || e), true);
        });
    }

    function downloadFile(id) {
        setStatus('Preparing download …');
        window.FileStore.get(id).then(function (rec) {
            if (!rec) throw new Error('File not found');
            triggerDownload(new Uint8Array(rec.data), rec.name);
            setStatus('');
        }).catch(function (e) {
            setStatus('Download failed: ' + (e && e.message || e), true);
        });
    }

    function clearLibrary() {
        if (!confirm(t('Delete ALL files from the library? This cannot be undone.'))) return;
        window.FileStore.clear().then(function () {
            setStatus('All files deleted');
            setTimeout(function () { setStatus(''); }, 3000);
            refreshLibrary();
        }).catch(function (e) {
            setStatus('Failed: ' + (e && e.message || e), true);
        });
    }
    function formatSize(n) {
        if (n == null) return '';
        if (n < 1024) return n + ' B';
        if (n < 1048576) return (n / 1024).toFixed(1) + ' KB';
        return (n / 1048576).toFixed(1) + ' MB';
    }
    function formatDate(ts) {
        if (!ts) return '';
        var d = new Date(ts);
        return d.toLocaleDateString() + ' ' + d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    }

    function loadFromFileStore(id) {
        window.FileStore.get(id).then(function (rec) {
            if (!rec) return;
            openFile({ id: rec.id, name: rec.name, ext: rec.ext, data: new Uint8Array(rec.data) });
        });
    }

    // ---------- home screen / upload ----------
    function showTitlebar(visible) {
        var bar = document.getElementById('titlebar');
        if (bar) bar.classList.toggle('show', visible);
    }
    function showHome() {
        var home = document.getElementById('home');
        if (home) home.style.display = '';
        var ed = document.getElementById('editor');
        if (ed) ed.style.display = 'none';
        showTitlebar(false);
        document.title = 'Clipboard Share 문서 편집기';
        refreshLibrary();
    }
    function hideHome() {
        var home = document.getElementById('home');
        if (home) home.style.display = 'none';
        var ed = document.getElementById('editor');
        if (ed) ed.style.display = '';
        showTitlebar(true);
        var nameEl = document.getElementById('titlebar-name');
        if (nameEl) nameEl.textContent = currentFile ? currentFile.name : '';
    }

    // Back-to-home: return to the file library without reloading the page.
    // The editor iframe is torn down, so any dirty document is persisted
    // first (best-effort; the iframe is still alive when this runs).
    function backToHome() {
        var finish = function () {
            stopAutosave();
            if (docEditor && typeof docEditor.destroyEditor === 'function') {
                try { docEditor.destroyEditor(); } catch (e) {}
            }
            docEditor = null;
            var ed = document.getElementById('editor');
            if (ed) ed.innerHTML = '';
            pendingBin = null;
            window.__ooPendingDocBin = null;
            window.__ooMediaMap = {};
            window.__ooMediaMapByBase = {};
            documentReady = false;
            emptyDocGotReal = false;
            currentFile = null;
            window.__ooDocExt = '';
            showHome();
        };
        if (editorDirtyNow()) {
            setStatus(t('Saving …'));
            doSave().then(finish, finish);
        } else {
            finish();
        }
    }

    function handleFiles(fileList) {
        if (!fileList || !fileList.length) return;
        var f = fileList[0];
        var name = f.name || 'untitled';
        var dot = name.lastIndexOf('.');
        var ext = dot > 0 ? name.slice(dot + 1).toLowerCase() : (PAGE_TYPE || 'docx');
        f.arrayBuffer().then(function (buf) {
            openFile({ name: name, ext: ext, data: new Uint8Array(buf) });
        });
    }

    function createNew(ext) {
        ext = ext || PAGE_TYPE || 'docx';
        var emptyUrl = { docx: 'assets/empty.docx', xlsx: 'assets/empty.xlsx', pptx: 'assets/empty.pptx', pdf: 'assets/empty.pdf' }[ext];
        if (!emptyUrl) return;
        fetch(emptyUrl).then(function (r) { return r.arrayBuffer(); }).then(function (buf) {
            var name = 'untitled.' + ext;
            openFile({ name: name, ext: ext, data: new Uint8Array(buf) });
        });
    }

    // ---------- x2t.wasm loading with progress ----------
    var X2T_DIR = '/vendor/sdkjs/common/wasm/x2t/';

    // Download the 36MB x2t.wasm with progress reporting.
    function downloadWasm(url) {
        return fetch(url).then(function (resp) {
            if (!resp.ok) throw new Error('wasm download failed: HTTP ' + resp.status);
            var total = parseInt(resp.headers.get('Content-Length'), 10) || 0;
            if (!resp.body) {
                setStatus('Downloading converter …');
                return resp.arrayBuffer().then(function (ab) { return gunzipIfNeeded(new Uint8Array(ab)); });
            }
            var reader = resp.body.getReader();
            var chunks = [], received = 0, lastUi = 0;
            function pump() {
                return reader.read().then(function (r) {
                    if (r.done) return;
                    chunks.push(r.value);
                    received += r.value.length;
                    var now = Date.now();
                    if (now - lastUi > 150) {
                        lastUi = now;
                        var pct = total ? ' (' + Math.floor(received * 100 / total) + '%)' : '';
                        setStatus('Downloading converter… ' + formatSize(received) +
                            (total ? ' / ' + formatSize(total) : '') + pct);
                    }
                    return pump();
                });
            }
            return pump().then(function () {
                var buf = new Uint8Array(received);
                var pos = 0;
                chunks.forEach(function (c) { buf.set(c, pos); pos += c.length; });
                return gunzipIfNeeded(buf);
            });
        });
    }

    // [cs] x2t.wasm (10.0, ~70MB) exceeds the 25MiB per-file limit of the
    // static host, so it ships gzip-compressed. Decompress in the browser unless
    // the host already removed the gzip layer (check the 1f 8b magic bytes).
    function gunzipIfNeeded(u8) {
        if (!(u8[0] === 0x1f && u8[1] === 0x8b)) return Promise.resolve(u8.buffer);
        setStatus('Unpacking converter …');
        var stream = new Blob([u8]).stream().pipeThrough(new DecompressionStream('gzip'));
        return new Response(stream).arrayBuffer();
    }

    // Fetch x2t.wasm, serving it from the Cache API when possible. The entry
    // is keyed by BUILD_VERSION, so a new deployment (which bumps the version)
    // downloads the new wasm exactly once and old entries are dropped.
    function fetchWasmWithProgress() {
        var url = X2T_DIR + 'x2t.wasm.gz';
        var cacheKey = X2T_DIR + 'x2t.wasm#v' + BUILD_VERSION;
        var cachePromise = (window.caches && caches.open)
            ? caches.open('oo-x2t-v1').catch(function () { return null; })
            : Promise.resolve(null);
        return cachePromise.then(function (cache) {
            if (!cache) return downloadWasm(url);
            return cache.match(cacheKey).then(function (cached) {
                if (cached) {
                    setStatus('Loading converter (cached) …');
                    return cached.arrayBuffer();
                }
                return downloadWasm(url).then(function (buf) {
                    return cache.put(cacheKey, new Response(buf)).catch(function () {}).then(function () {
                        // drop entries left over from older versions
                        return cache.keys().catch(function () { return []; }).then(function (keys) {
                            var keep = new Request(cacheKey).url;
                            return Promise.all(keys.map(function (k) {
                                return k.url !== keep ? cache.delete(k) : null;
                            }));
                        }).catch(function () {}).then(function () { return buf; });
                    });
                });
            });
        });
    }

    function loadX2t() {
        if (window.__x2tReadyPromise) return window.__x2tReadyPromise;
        window.__x2tReadyPromise = new Promise(function (resolve, reject) {
            fetchWasmWithProgress().then(function (bytes) {
                setStatus('Initializing converter …');
                // Set up the emscripten Module, then load x2t.js dynamically.
                // wasmBinary skips the glue's own fetch; locateFile keeps any
                // secondary lookup (e.g. the .mem file) on the right path.
                // onRuntimeInitialized fires once the wasm HEAP views exist.
                window.Module = {
                    wasmBinary: bytes,
                    locateFile: function (path) { return X2T_DIR + path; },
                    onRuntimeInitialized: function () { resolve(window.Module); }
                };
                var s = document.createElement('script');
                s.src = X2T_DIR + 'x2t.js';
                s.onerror = function () { reject(new Error('Failed to load x2t.js')); };
                document.head.appendChild(s);
            }).catch(function (e) {
                window.__x2tReadyPromise = null; // allow retry on next action
                setStatus('Converter load failed: ' + (e && e.message || e), true);
                reject(e);
            });
        });
        return window.__x2tReadyPromise;
    }

    // ---------- boot ----------
    function boot() {
        translateShell();
        // SW cache version guard first: on a version change it clears the
        // editor caches and reloads once (see ensureFreshSwCache). If it
        // triggered the reload, abort — the fresh page will boot normally.
        ensureFreshSwCache().then(function (reloading) {
            if (!reloading) bootMain();
        });
    }

    function bootMain() {
        loadX2t().catch(function () {});

        // message bridge from the editor iframe
        window.addEventListener('message', function (e) {
            var d = e.data;
            if (!d) return;
            if (d.type === 'oo-empty-doc-real') emptyDocGotReal = true;
            else if (d.type === 'oo-save-request') saveDocument();
            else if (d.type === 'oo-download-request') handleOfflineDownload(d);
            else if (d.type === 'oo-download-unsupported') {
                var nm = OFFLINE_FORMAT_NAMES[d.format] || ('format #' + d.format);
                setStatus('Offline export of ' + nm + ' is not supported', true);
                setTimeout(function () { setStatus(''); }, 4000);
            }
            else if (d.type === 'oo-print-unsupported') {
                setStatus('Printing is not supported offline', true);
                setTimeout(function () { setStatus(''); }, 4000);
            }
        });

        // Ctrl+S / Cmd+S
        window.addEventListener('keydown', function (e) {
            if ((e.ctrlKey || e.metaKey) && e.key === 's') {
                e.preventDefault();
                saveDocument();
            }
        });

        // home screen wiring
        var uploadInput = document.getElementById('upload-input');
        var uploadBtn = document.getElementById('upload-btn');
        var dropZone = document.getElementById('home');
        if (uploadBtn && uploadInput) {
            uploadBtn.addEventListener('click', function () { uploadInput.click(); });
            uploadInput.addEventListener('change', function () { handleFiles(uploadInput.files); });
        }
        // title bar buttons
        var backHomeBtn = document.getElementById('back-home');
        if (backHomeBtn) backHomeBtn.addEventListener('click', backToHome);
        var saveBtn = document.getElementById('save-btn');
        if (saveBtn) saveBtn.addEventListener('click', function () { saveDocument(); });

        // Autosave on tab-hide/close. visibilitychange->hidden fires on tab
        // switch while the page is still fully alive (the async IndexedDB
        // write completes normally); pagehide covers tab close/navigation
        // (best-effort: the browser usually lets a just-started same-origin
        // IndexedDB transaction finish). Either way the last dirty state
        // lands in the library even if the 30s timer hasn't fired yet.
        var saveIfDirty = function () {
            if (editorDirtyNow() && !saveInFlight) saveDocument();
        };
        document.addEventListener('visibilitychange', function () {
            if (document.visibilityState === 'hidden') saveIfDirty();
        });
        window.addEventListener('pagehide', saveIfDirty);

        // library toolbar
        var clearAllBtn = document.getElementById('clear-all-btn');
        if (clearAllBtn) clearAllBtn.addEventListener('click', clearLibrary);

        // per-type "New" buttons (data-new="docx|xlsx|pptx")
        var newActions = document.getElementById('new-actions');
        if (newActions) {
            newActions.addEventListener('click', function (e) {
                var btn = e.target.closest('[data-new]');
                if (btn) createNew(btn.getAttribute('data-new'));
            });
            // On a type-specific entry page, highlight that type's button
            if (PAGE_TYPE) {
                var focus = newActions.querySelector('[data-new="' + PAGE_TYPE + '"]');
                if (focus) focus.classList.add('primary');
            }
        }
        if (dropZone) {
            dropZone.addEventListener('dragover', function (e) { e.preventDefault(); dropZone.classList.add('dragover'); });
            dropZone.addEventListener('dragleave', function () { dropZone.classList.remove('dragover'); });
            dropZone.addEventListener('drop', function (e) {
                e.preventDefault();
                dropZone.classList.remove('dragover');
                handleFiles(e.dataTransfer.files);
            });
        }

        // ?open=<id> deep link into a stored file
        if (PARAMS.open && window.FileStore) {
            loadFromFileStore(PARAMS.open);
        } else if (PARAMS.autonew) {
            createNew(PAGE_TYPE || 'docx');
        } else {
            showHome();
        }
    }

    // expose a small API for embedding / testing
    window.OfficeApp = {
        openFile: openFile,
        saveDocument: saveDocument,
        createNew: createNew,
        version: BUILD_VERSION,
        // autosave internals (used by the e2e suite)
        isDirty: function () { return editorDirtyNow(); },
        autosaveNow: function () { autosaveTick(); return saveInFlight; }
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', boot);
    } else {
        boot();
    }
})();
