# 원본(wasmtools/office-web 80b7e9e)에서 바꾼 곳

| 파일 | 바꾼 내용 | 이유 |
|---|---|---|
| `.github/workflows/build.yaml` | 삭제 | 원본 이미지 빌드가 포크에서 돌지 않게 |
| `web/app.js` | x2t.wasm.gz 를 받아 브라우저에서 풀기 | 10.0 x2t(약 70MB)가 파일당 25MiB 한도를 넘음 |
| `web/app.js` | 내려받기 형식에 XLS(258) 추가 | 10.0 에서 XLS 저장 가능 |
| `web/app.js` | PDF 저장 글꼴에 Liberation Sans 추가 | 입력 글꼴(Arial)이 없으면 x2t 10.0 이 영문 글리프를 두 칸씩 밀어 씀 |
| `web/app.js` | BUILD_VERSION 10.0.0-cs1, 제목 한국어 | 캐시 갱신, 한국어 화면 |
| `web/app.js` | 옛날 형식(doc 등) 열기 실패 시 OOXML 을 거쳐 다시 변환 | x2t 10.0 에서 doc → bin 이 ret=88 로 실패, doc → docx 는 성공 |
| `web/app.js` | 맞춤법 검사 기본 끔 | 사전 파일을 싣지 않아 모든 단어에 빨간 줄이 생김 |
| `web/sw.js` (새 파일) | 캐시를 지우고 스스로 등록 해제하는 서비스 워커 | 예전 편집기가 등록한 /sw.js 가 남은 브라우저가 옛 화면을 보지 않게 |
| `web/index.html` | 한국어 문구, AGPL·소스·상표·글꼴 라이선스 안내, 한글 편집기 링크 | |
| `web/vendor/web-apps/apps/spreadsheeteditor/main/offline-shim.js` | XLS(258) 허용 | 위와 같음 |
| `web/app.js` | `?embed=cs` 모드와 cbEdit v1 창구(open/export/markSaved/state/setTheme, ready/chrome/dirty/saveRequest/error). 저장·내려받기·자동저장·홈 화면은 embed 에서 끔 | Clipboard Share 가 iframe 으로 쓰기 위함 |
| `web/cs-embed-core.js` (새 파일) | 출처 허용 목록, 메서드 분기, 형식 규칙(순수 로직) | 단위 시험(`cs/test/`) |
| `web/index.html` | cs-embed-core.js 로드, embed 일 때 홈·제목줄 숨김 | |
| `cs/build/_headers` | frame-ancestors 에 vercel.app, localhost, 127.0.0.1 추가 | cbEdit 규약의 부모 허용 목록 |
| `cs/test/`, `cs/gate/embed.mjs`, `.github/workflows/cs-deploy.yml` | cbEdit 단위 시험과 embed 게이트 | |
| `web/app.js` | BUILD_VERSION 10.0.0-cs2-embed | 캐시 갱신 |

vendor(web-apps, sdkjs, fonts)는 저장소에 있는 9.4 판을 쓰지 않고 `cs/build/build.sh` 가 매번 새로 만든다.

빌드 때 vendor 에 거는 패치(`cs/build/build.sh`): api.js 버전 경로, 편집기 4종 시작 HTML 에 offline-shim 로드, common/main/lib 의 빠진 RequireJS 의존성(Window/BaseView) 보충, sdkjs 공백·구문 줄이기.
