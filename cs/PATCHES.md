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
| `web/index.html` | 한국어 문구, AGPL·소스·상표·글꼴 라이선스 안내, 한글 편집기 링크 | |
| `web/vendor/web-apps/apps/spreadsheeteditor/main/offline-shim.js` | XLS(258) 허용 | 위와 같음 |

vendor(web-apps, sdkjs, fonts)는 저장소에 있는 9.4 판을 쓰지 않고 `cs/build/build.sh` 가 매번 새로 만든다.

빌드 때 vendor 에 거는 패치(`cs/build/build.sh`): api.js 버전 경로, 편집기 4종 시작 HTML 에 offline-shim 로드, common/main/lib 의 빠진 RequireJS 의존성(Window/BaseView) 보충, sdkjs 공백·구문 줄이기.
