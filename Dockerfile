# Standalone in-browser Office editor.
#
# Serves the web/ directory with a minimal Python static server. All document
# processing happens client-side (x2t.wasm); the server only delivers assets.
#
# Build:
#   docker build -t office-web .
# Run:
#   docker run -d -p 8080:8080 office-web
# Open:
#   http://localhost:8080/            (home / file library)
#   http://localhost:8080/docx.html   (document entry)
#   http://localhost:8080/xlsx.html   (spreadsheet entry)
#   http://localhost:8080/pptx.html   (presentation entry)
#   http://localhost:8080/pdf.html    (PDF entry)

FROM python:3.11-slim

WORKDIR /app

# The web frontend (editor + x2t.wasm + assets). web/vendor/decoded-fonts/ is
# gitignored (a build artifact), so it is absent from the checkout — regenerate
# it from the encrypted pool (web/vendor/fonts) during the build; the PDF
# export's font staging fetches it (see app.js x2tRenderBinToPdf).
COPY web/ /app/web/

# Minimal static server + build-time tooling
COPY server.py /app/server.py
COPY precompress.py /app/precompress.py
COPY scripts/extract-fonts.py /app/scripts/extract-fonts.py

# Build-time step: decode the font pool + pre-compress static assets to brotli
# so the server serves them with zero runtime CPU cost. Requires brotli.
RUN pip install --no-cache-dir brotli fonttools \
    && python /app/scripts/extract-fonts.py /app/web/vendor/fonts /app/build/fonts \
        --deploy /app \
    && python /app/precompress.py --root /app/web --min-size 4096 --quiet

ENV PORT=8080
EXPOSE 8080

CMD ["python", "/app/server.py", "--port", "8080", "--root", "/app/web"]
