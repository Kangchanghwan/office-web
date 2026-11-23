#!/usr/bin/env bash
# office.clipboardapp.org 배포본을 out/ 에 만든다. (GitHub Actions ubuntu-latest 기준, docker 필요)
#
#   web/ (앱, 오프라인 연결 코드)              이 저장소
#   web-apps + 실행 중 생성되는 sdkjs 파일      공식 Docker 이미지 (digest 고정)
#   sdkjs 코드                                 공식 소스 태그를 이어 붙여 빌드 (난독화 없음)
#   x2t                                        Kangchanghwan/onlyoffice-x2t-wasm 릴리스 (sha256 확인)
#   글꼴                                       ranuts/document 의 공개 라이선스 글꼴 묶음
set -euo pipefail

ROOT=$(cd "$(dirname "$0")/../.." && pwd)
OUT=${OUT:-$ROOT/out}
WORK=${WORK:-$ROOT/.cs-work}
U=$ROOT/cs/upstream.json
j() { python3 -c "import json,sys;d=json.load(open('$U'));print(eval('d'+sys.argv[1]))" "$1"; }

IMAGE=$(j "['web_apps']['image']")
SDKJS_TAG=$(j "['sdkjs']['tag']")
FORMS_TAG=$(j "['sdkjs']['forms_tag']")
X2T_URL=$(j "['x2t']['url']")
X2T_SHA=$(j "['x2t']['sha256']")
FONTS_REPO=$(j "['fonts']['repo']")
FONTS_COMMIT=$(j "['fonts']['commit']")

rm -rf "$OUT" "$WORK"; mkdir -p "$OUT" "$WORK"

echo "== 1. 공식 이미지에서 web-apps 와 생성 파일 꺼내기"
if [ -n "${DS_DIR:-}" ]; then
  # 로컬 시험용: 이미 꺼내 둔 폴더를 쓴다
  mkdir -p "$WORK/ds"; cp -R "$DS_DIR/web-apps" "$DS_DIR/sdkjs" "$WORK/ds/"
else
  docker pull -q "$IMAGE"
  docker run -d --name cs-ds -e JWT_ENABLED=false "$IMAGE" >/dev/null
  D=/var/www/onlyoffice/documentserver
  for i in $(seq 1 60); do
    docker exec cs-ds test -f $D/sdkjs/common/AllFonts.js && break; sleep 5
  done
  sleep 15
  mkdir -p "$WORK/ds"
  docker exec cs-ds tar -C $D -cf - web-apps sdkjs | tar -xf - -C "$WORK/ds"
  docker rm -f cs-ds >/dev/null
fi
chmod -R u+rwX "$WORK/ds"

echo "== 2. sdkjs $SDKJS_TAG 소스 빌드"
git clone -q --depth 1 --branch "$SDKJS_TAG" https://github.com/ONLYOFFICE/sdkjs.git "$WORK/sdkjs"
git clone -q --depth 1 --branch "$FORMS_TAG" https://github.com/ONLYOFFICE/sdkjs-forms.git "$WORK/sdkjs-forms"
(cd "$WORK/sdkjs/build" && PRODUCT_VERSION=10.0.0 BUILD_NUMBER=${SDKJS_TAG##*.} python3 build.py --addon "$WORK/sdkjs-forms" >/dev/null)

echo "== 3. 앱 복사"
cp -R "$ROOT/web/." "$OUT/"
SHIMS="$WORK/shims"; mkdir -p "$SHIMS"
for e in documenteditor spreadsheeteditor presentationeditor pdfeditor; do
  cp "$ROOT/web/vendor/web-apps/apps/$e/main/offline-shim.js" "$SHIMS/$e.js"
done
rm -rf "$OUT/vendor/web-apps" "$OUT/vendor/sdkjs" "$OUT/vendor/fonts" "$OUT/vendor/decoded-fonts"

echo "== 4. vendor 조립"
cp -R "$WORK/ds/web-apps" "$OUT/vendor/web-apps"
cp -R "$WORK/sdkjs/deploy/sdkjs" "$OUT/vendor/sdkjs"
python3 - "$WORK/ds/sdkjs" "$OUT/vendor/sdkjs" <<'EOF'
import os, shutil, sys
src, dst = sys.argv[1], sys.argv[2]; n = 0
for d, _, fs in os.walk(src):
    for f in fs:
        rel = os.path.relpath(os.path.join(d, f), src)
        if rel.endswith('sdk-all.bin'): continue          # 서버(x2t) 전용 스냅샷
        out = os.path.join(dst, rel)
        if not os.path.exists(out):
            os.makedirs(os.path.dirname(out), exist_ok=True); shutil.copy2(os.path.join(d, f), out); n += 1
print(f'  생성 파일 {n}개 복사')
EOF

echo "== 5. 공개 라이선스 글꼴 (ranuts $FONTS_COMMIT)"
git init -q "$WORK/fonts-src"
(cd "$WORK/fonts-src" && git remote add origin "$FONTS_REPO" \
  && git sparse-checkout set --no-cone /public/fonts/ /public/sdkjs/common/AllFonts.js '/public/sdkjs/common/Images/fonts_thumbnail*' /docs/font-licenses.md /docs/fonts.md >/dev/null \
  && git fetch -q --depth 1 --filter=blob:none origin "$FONTS_COMMIT" && git checkout -q FETCH_HEAD)
cp -R "$WORK/fonts-src/public/fonts" "$OUT/vendor/fonts"
cp "$WORK/fonts-src/public/sdkjs/common/AllFonts.js" "$OUT/vendor/sdkjs/common/AllFonts.js"
cp "$WORK/fonts-src/public/sdkjs/common/Images/"fonts_thumbnail* "$OUT/vendor/sdkjs/common/Images/"
mkdir -p "$OUT/licenses"
cp "$WORK/fonts-src/docs/font-licenses.md" "$OUT/licenses/font-licenses.md"
cp "$WORK/fonts-src/docs/fonts.md" "$OUT/licenses/fonts.md"

echo "== 6. 덜어내기 (도움말, Visio)"
find "$OUT/vendor/web-apps/apps" -type d -path '*/main/resources/help' -prune -exec rm -rf {} +
rm -rf "$OUT/vendor/web-apps/apps/visioeditor" "$OUT/vendor/sdkjs/visio"

echo "== 7. 오프라인 연결 코드와 패치"
for e in documenteditor spreadsheeteditor presentationeditor pdfeditor; do
  cp "$SHIMS/$e.js" "$OUT/vendor/web-apps/apps/$e/main/offline-shim.js"
done
python3 - "$OUT/vendor/web-apps/apps" <<'EOF'
import re, sys
base = sys.argv[1]
api = f'{base}/api/documents/api.js'; s = open(api).read()
s2 = re.sub(r"const ver = '/10\.0\.0-[0-9a-f]+';", "const ver = '/10.0.0-{{HASH_POSTFIX}}';", s)
assert s2 != s, 'api.js ver'; open(api, 'w').write(s2)
tag = '<script src="../../../vendor/requirejs/require.js"></script>'
for e in ['documenteditor', 'spreadsheeteditor', 'presentationeditor']:
    p = f'{base}/{e}/main/index.html'; s = open(p).read(); assert s.count(tag) == 1, p
    open(p, 'w').write(s.replace(tag, '<script src="offline-shim.js"></script>\n    ' + tag))
p = f'{base}/pdfeditor/main/index.html'; s = open(p).read(); old = '        function startApp() {\n'
assert s.count(old) == 1, p
open(p, 'w').write(s.replace(old, old + "            document.write('<script src=\"offline-shim.js\"><\\/script>');\n"))
print('  패치 5곳')
EOF

echo "== 7b. 공통 UI 파일의 빠진 의존성 선언 보충 (wasmtools 9.4 와 같은 수정)"
python3 - "$OUT/vendor/web-apps/apps/common/main/lib" <<'EOF2'
# Window/BaseView 를 상속하면서 define([]) 로 의존성을 비워 둔 파일은, 로드 순서에 따라
# "Cannot read properties of undefined (reading 'extend')" 로 편집기가 뜨지 않는다.
import os, re, sys
n = 0
for d, _, fs in os.walk(sys.argv[1]):
    for f in fs:
        if not f.endswith('.js'): continue
        p = os.path.join(d, f); s = open(p, encoding='utf-8').read()
        m = re.search(r'define\(\s*\[\s*\]\s*,\s*function', s)
        if not m: continue
        dep = None
        if re.search(r'Common\.UI\.Window\.extend\(', s): dep = 'common/main/lib/component/Window'
        elif re.search(r'Common\.UI\.BaseView\.extend\(', s): dep = 'common/main/lib/component/BaseView'
        if not dep: continue
        s = s[:m.start()] + "define(['" + dep + "'], function" + s[m.end():]
        open(p, 'w', encoding='utf-8').write(s); n += 1
print(f'  {n}개 파일')
EOF2

echo "== 8. sdkjs 줄이기 (공백·구문만, 이름은 그대로)"
for f in $(find "$OUT/vendor/sdkjs" -name 'sdk-all*.js'); do
  npx --yes esbuild@0.24.0 "$f" --minify-whitespace --minify-syntax --charset=utf8 --log-level=error --outfile="$f.tmp"
  mv "$f.tmp" "$f"
done

echo "== 9. x2t"
curl -sfL -o "$WORK/x2t.zip" "$X2T_URL"
echo "$X2T_SHA  $WORK/x2t.zip" | (command -v sha256sum >/dev/null && sha256sum -c - || shasum -a 256 -c -)
X2T_DIR="$OUT/vendor/sdkjs/common/wasm/x2t"; mkdir -p "$X2T_DIR"
unzip -oq "$WORK/x2t.zip" x2t.js x2t.wasm -d "$WORK/x2t"
cp "$WORK/x2t/x2t.js" "$X2T_DIR/x2t.js"
gzip -9 -n -c "$WORK/x2t/x2t.wasm" > "$X2T_DIR/x2t.wasm.gz"

echo "== 10. PDF 저장용 글꼴 색인"
python3 "$ROOT/scripts/extract-fonts.py" "$OUT/vendor/fonts" "$WORK/decoded" --deploy "$OUT" | tail -2

echo "== 11. 헤더와 빌드 정보"
cp "$ROOT/cs/build/_headers" "$OUT/_headers"
python3 - "$OUT/build-identity.json" <<EOF
import json, os, sys
json.dump({
  "source": "fork", "repo": "Kangchanghwan/office-web",
  "commit": os.environ.get("GITHUB_SHA", "local"), "ref": os.environ.get("GITHUB_REF_NAME", "local"),
  "onlyoffice": "10.0.0", "web_apps_image": "$IMAGE", "sdkjs": "$SDKJS_TAG",
  "x2t_sha256": "$X2T_SHA", "fonts_commit": "$FONTS_COMMIT"
}, open(sys.argv[1], "w"), indent=2)
EOF
echo "완료: $OUT"
