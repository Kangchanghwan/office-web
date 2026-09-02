#!/usr/bin/env python3
"""pptx -> 'Download as PDF' e2e: open hello.pptx, export PDF via format 513
(exercises the new render-bin path end to end in a real browser)."""
import base64, pathlib, sys
from playwright.sync_api import sync_playwright

BASE = "http://localhost:18093"
PPTX = "/home/coder/project/office-web/saved-docs/hello.pptx"

def editor_frame(page):
    for f in page.frames:
        if f.name == "frameEditor":
            return f
    return None

def wait_ready(page, timeout_ms=90000):
    page.wait_for_function("""() => {
        const fr = document.querySelector('iframe[name=frameEditor]');
        return fr && fr.contentWindow && fr.contentWindow.Asc && fr.contentWindow.Asc.editor;
    }""", timeout=timeout_ms)
    fr = editor_frame(page)
    fr.wait_for_function("() => window.Asc.editor.isDocumentLoadComplete === true", timeout=timeout_ms)

with sync_playwright() as p:
    b = p.chromium.launch()
    page = b.new_page()
    errors = []
    page.on("pageerror", lambda e: errors.append(str(e)))
    page.goto(BASE + "/pptx.html", wait_until="load")
    page.wait_for_function("() => window.OfficeApp && window.OfficeApp.openFile", timeout=30000)
    data = pathlib.Path(PPTX).read_bytes()
    page.evaluate("""(b64) => {
        const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
        window.OfficeApp.openFile({ name: 'hello.pptx', ext: 'pptx', data: bin });
    }""", base64.b64encode(data).decode())
    wait_ready(page)
    page.wait_for_timeout(2000)

    fr = editor_frame(page)
    has_getpdf = fr.evaluate("() => typeof window.Asc.editor.asc_nativeGetPDF === 'function'")
    print("PASS  asc_nativeGetPDF available" if has_getpdf else "FAIL  asc_nativeGetPDF missing")

    with page.expect_download(timeout=120000) as dl:
        fr.evaluate("window.Asc.editor.asc_DownloadAs({ fileType: 513 })")
    dl_path = "/tmp/pdftest/hello-exported.pdf"
    pathlib.Path("/tmp/pdftest").mkdir(exist_ok=True)
    dl.value.save_as(dl_path)
    pdf = pathlib.Path(dl_path).read_bytes()
    ok = pdf[:4] == b"%PDF"
    print(("PASS  " if ok else "FAIL  ") + f"pptx download-as PDF: {len(pdf)} bytes, head={pdf[:8]!r}")
    real = len(pdf) > 5000
    print(("PASS  " if real else "FAIL  ") + "PDF non-trivial (>5KB, real content)")
    print(("PASS  " if not errors else "FAIL  ") + f"no page errors ({len(errors)})")
    if errors: print(errors[:3])
    b.close()
    sys.exit(0 if ok and real and not errors else 1)
