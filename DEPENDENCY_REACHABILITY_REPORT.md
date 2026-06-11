# 프로젝트별 의존성 및 함수 도달성 분석

분석 기준:

- 분석일: 2026-06-12
- Python AST, 전체 저장소 이름 참조, import 관계, FastAPI 데코레이터, Qt signal 연결,
  `getattr` 동적 호출, HTTP handler override, 브라우저 확장 이벤트 진입점을 함께 확인했다.
- "미도달"은 현재 저장소의 표준 앱/서버/확장 진입점에서 호출되는 경로가 없다는 뜻이다.
- 진단 스크립트의 `main()`처럼 사용자가 직접 실행하는 별도 CLI 진입점은 미도달 함수로 보지 않았다.
- 향후 AI/API/UI 기능을 연결하려고 만든 명시적 미구현 스텁과 얇은 확장용 위임 메서드는
  삭제 대상인 미도달 함수에서 제외했다.

---

## 1. AI-grammary

### 프로젝트 경계

- 루트: `AI-grammary/`
- 데스크톱 앱 진입점: `main.py`, `client/main.py`
- 서버 진입점: `server/launch_server.py` -> `server/main.py`
- 브라우저 확장 진입점: `browser_extension/manifest.json` -> `content.js`
- 다른 두 프로젝트의 코드는 이 프로젝트의 호출 그래프에 포함하지 않았다.

### 주요 내부 의존성

```text
main.py / client/main.py
  -> client.ui.main_window.App
     -> client.core.analyzer.TextAnalyzer
        -> client.core.ai_client.AIClient
           -> client.core.local_server
     -> client.core.auth_api_client.AuthAPIClient
     -> client.input.realtime_text_monitor
        -> client.input.ai_grammary_text_reader
        -> client.input.browser_extension_bridge
        -> client.input.keyboard_monitor
     -> client.input.output_applier
     -> client.ui.result_panel

server/launch_server.py
  -> server/main.py FastAPI app
     -> server.ai_service
     -> server.auth
     -> server.database -> server.models
     -> server.schemas

browser_extension/content.js
  <-> http://127.0.0.1:8766
  <-> client.input.browser_extension_bridge
```

### 외부 의존성

- UI: `PyQt5`
- Windows 자동화: `pywin32`, `pywinauto`, `uiautomation`, `pynput`, `psutil`
- 클립보드/HTTP: `pyperclip`, `requests`
- AI: `openai`, `httpx`
- 서버/DB/인증: `fastapi`, `uvicorn`, `sqlalchemy`, `pydantic`, `python-dotenv`,
  `bcrypt`, `python-jose`, `psycopg2-binary`

### 미도달 함수: 높은 확신

- `client/core/analyzer.py:15` `TextAnalyzer.analyze_spelling`
- `client/core/auth_api_client.py:106` `AuthAPIClient.list_logs`
- `client/input/notepad_monitor.py:68` `get_active_notepad_text`
- `client/input/output_applier.py:160` `OutputApplier.remap_hwp_selection_segments`
- `client/input/output_applier.py:1492` `OutputApplier._apply_hwp_full_document_rtf_replacement`
- `client/input/output_applier.py:1641` `OutputApplier._build_hwp_selection_html`
- `client/input/output_applier.py:1701` `OutputApplier._set_html_clipboard`
- `client/input/output_applier.py:1740` `OutputApplier._apply_hwp_selection_runs_with_inline_style`
- `client/input/output_applier.py:1817` `OutputApplier._apply_hwp_saved_selection_rich_replacement`
- `client/input/output_applier.py:1838` `OutputApplier._apply_hwp_saved_selection_segmented_replacement`
- `client/input/output_applier.py:2424` `OutputApplier._hwpml2x_summary_has_mixed_shapes`
- `client/input/output_applier.py:2428` `OutputApplier._hwpml2x_summary_matches_text`
- `client/input/output_applier.py:2477` `OutputApplier._apply_hwp_style`
- `client/ui/main_window.py:716` `App.run_spell_check_sync`
- `client/ui/main_window.py:1244` `App._extract_corrected_text`
- `client/ui/main_window.py:1654` `App._current_title`
- `client/ui/main_window.py:1753` `App.sync_restored_login_settings`
- `browser_extension/content.js:125` `normalizeComparableText`

### 미도달 가능성: 중간 확신

- `client/input/ai_grammary_text_reader.py:700`
  `ActiveWordReader._first_visible_character_range`

현재 저장소에는 직접 참조가 없다. 다만 Word COM 자동화 계층은 런타임 객체 특성에 따라
간접 접근할 수 있으므로 삭제 전 실제 Word 시나리오 확인이 권장된다.

### 함수가 아닌 독립/고립 모듈

- `client/config.py`: 앱 진입점에서 import되지 않는다.
- `client/ui/auth_dialog.py`: 현재 UI 흐름에서 import되지 않는다.
- `client/input/hwp_com_diagnostic.py`: 독립 진단 CLI다.
- `client/input/hwp_textfile_diagnostic.py`: 독립 진단 CLI다.
- `server/checkdb.py`: 독립 DB 점검 스크립트다.

### 오탐에서 제외한 항목

- `browser_extension_bridge.py`의 `do_GET`, `do_POST`, `do_OPTIONS`, `log_message`:
  `BaseHTTPRequestHandler`가 이름으로 호출한다.
- `read_selection_info`, `reset_state`: `getattr` 기반 동적 호출이 존재한다.
- FastAPI 라우트 함수: 데코레이터가 외부 HTTP 진입점으로 등록한다.

### 향후 기능용 뼈대로 판단해 제외한 항목

이 프로젝트에서는 명시적인 `NotImplementedError`/`pass` 기반 AI 스텁이나,
향후 기능 연결만을 위한 것이 분명한 얇은 위임 메서드를 확인하지 못했다.

---

## 2. WA_yunseo

### 프로젝트 경계

- 루트: `WA_yunseo/`
- 데스크톱 앱 진입점: `main.py`, `client/main.py`
- 서버 진입점: `server/main.py`
- 브라우저 확장 진입점:
  `browser_extension/manifest.json` -> `background.js`, `popup.js`, `content.js`
- 다른 두 프로젝트의 코드는 이 프로젝트의 호출 그래프에 포함하지 않았다.

### 주요 내부 의존성

```text
main.py / client/main.py
  -> client.ui.dpi.configure_high_dpi
  -> client.ui.main_window.App
     -> client.ui.main_overlay
     -> client.ui.mini_overlay
     -> client.ui.result_panel
     -> client.ui.spelling_inspection_overlay
     -> client.core.analyzer -> client.core.ai_client
     -> client.core.auth_api_client
     -> client.input.drag_selection_monitor
     -> client.input.realtime_text_monitor
     -> client.input.output_applier
     -> client.input.global_hotkey

server/main.py FastAPI app
  -> server.ai_service -> server.ai_cache
  -> server.auth
  -> server.database -> server.models
  -> server.schemas

browser_extension/content.js
  -> chrome.runtime messaging
  -> browser_extension/background.js
  -> http://127.0.0.1:8766
  -> client.input.browser_extension_bridge
```

### 외부 의존성

- UI: `PyQt5`
- Windows 자동화: `pywin32`, `pywinauto`, `uiautomation`, `pynput`, `psutil`
- 클립보드/HTTP: `pyperclip`, `requests`
- AI: `openai`
- 서버/DB/인증: `fastapi`, `uvicorn`, `sqlalchemy`, `pydantic`, `python-dotenv`,
  `bcrypt`, `python-jose`, `psycopg2-binary`, `python-multipart`

### 미도달 함수: 높은 확신

- `client/input/drag_selection_monitor.py:286`
  `_has_large_word_non_document_surface`
- `client/input/notepad_monitor.py:68` `get_active_notepad_text`
- `client/ui/main_window.py:1152` `App._has_large_word_non_document_surface`
- `client/ui/main_window.py:1670` `App._schedule_word_focus_restore`
- `client/ui/main_window.py:3408` `App.sync_restored_login_settings`
- `browser_extension/content.js:154` `fallbackEditable`
- `browser_extension/content.js:386` `normalizeComparableText`
- `browser_extension/content.js:733` `fragmentHtml`
- `browser_extension/content.js:739` `styleSegmentsFromRange`
- `browser_extension/content.js:1423` `statusTextForElement`

### 미도달 가능성: 중간 확신

- `client/input/ai_grammary_text_reader.py:617`
  `ActiveWordReader._first_visible_character_range`

Word COM 계층의 런타임 간접 호출 가능성을 고려해 중간 확신으로 분류했다.

### 함수가 아닌 독립/고립 모듈

- `client/config.py`: 앱 진입점에서 import되지 않는다.
- `client/ui/auth_dialog.py`: 현재 UI 흐름에서 import되지 않는다.
- `client/input/hwp_com_diagnostic.py`: 독립 진단 CLI다.
- `client/input/hwp_textfile_diagnostic.py`: 독립 진단 CLI다.
- `server/checkdb.py`: 독립 DB 점검 스크립트다.

### 오탐에서 제외한 항목

- `browser_extension_bridge.py` HTTP handler override
- FastAPI 라우트 함수와 dependency 함수
- Qt signal에 연결된 메서드
- `poll`, `read_style_info`처럼 인터페이스 구현체에서 런타임 선택되는 메서드

### 향후 기능용 뼈대로 판단해 제외한 항목

다음 함수는 현재 호출되지 않지만 향후 AI/UI 기능 연결을 위한 구조로 보아 미도달
삭제 후보에서 제외했다.

- `client/ui/main_overlay.py:971` `MainOverlay.toggle_settings`
- `client/ui/main_overlay.py:1138` `MainOverlay.replace_summary_result`
- `client/ui/mini_overlay.py:1201` `MiniOverlay.tone_favorite_tones`
- `client/ui/mini_overlay.py:1210` `MiniOverlay.show_choice_prompt`
- `server/ai_service.py:115` `AIService.summarize_text`
- `server/ai_service.py:118` `AIService.evaluate_text`
- `server/ai_service.py:121` `AIService.recommend_title`
- `server/ai_service.py:124` `AIService.convert_tone`

서버의 네 AI 메서드는 현재 `NotImplementedError`를 발생시키지만 요약, 평가, 제목 추천,
문체 변경 기능을 붙이기 위한 명시적 확장 지점이다.

---

## 3. Whiting-assistant-team-project

### 프로젝트 경계

- Git 루트: `Whiting-assistant-team-project/`
- 실제 애플리케이션 루트:
  `Whiting-assistant-team-project/Writing-Assistant-AI-responsive/`
- 데스크톱 앱 진입점: 실제 앱 루트의 `main.py`, `client/main.py`
- 서버 진입점: 실제 앱 루트의 `server/main.py`
- 브라우저 확장 진입점:
  실제 앱 루트의 `browser_extension/manifest.json` -> `popup.js`, `content.js`
- 바깥 Git 루트의 `README.md` 외 코드는 실제 앱 호출 그래프에 섞지 않았다.

### 주요 내부 의존성

```text
main.py / client/main.py
  -> client.ui.main_window.App
     -> client.ui.main_overlay
     -> client.ui.mini_overlay
     -> client.ui.result_panel
     -> client.ui.spelling_inspection_overlay
     -> client.core.analyzer -> client.core.ai_client
     -> client.core.auth_api_client
     -> client.core.tone_presets
     -> client.input.drag_selection_monitor
     -> client.input.realtime_text_monitor
     -> client.input.output_applier
     -> client.input.browser_extension_bridge

server/main.py FastAPI app
  -> server.ai_service
  -> server.auth
  -> server.database -> server.models
  -> server.schemas

browser_extension/content.js
  <-> http://127.0.0.1:8766
  <-> client.input.browser_extension_bridge
```

특이점: 이 프로젝트의 `client/core/ai_client.py`는 로컬 FastAPI 호출보다
OpenAI SDK 직접 호출 의존성이 강하다.

### 외부 의존성

- UI: `PyQt5`
- Windows 자동화: `pywin32`, `pywinauto`, `uiautomation`, `pynput`, `psutil`
- 클립보드/HTTP: `pyperclip`, `requests`
- AI: `openai`
- 서버/DB/인증: `fastapi`, `uvicorn`, `sqlalchemy`, `pydantic`, `python-dotenv`,
  `bcrypt`, `python-jose`, `psycopg2-binary`, `python-multipart`

### 미도달 함수: 높은 확신

- `client/input/drag_selection_monitor.py:286`
  `_has_large_word_non_document_surface`
- `client/input/notepad_monitor.py:68` `get_active_notepad_text`
- `client/ui/main_window.py:1175` `App._has_large_word_non_document_surface`
- `client/ui/main_window.py:1709` `App._schedule_word_focus_restore`
- `client/ui/main_window.py:3537` `App.sync_restored_login_settings`
- `browser_extension/content.js:153` `fallbackEditable`
- `browser_extension/content.js:361` `normalizeComparableText`
- `browser_extension/content.js:479` `fragmentHtml`
- `browser_extension/content.js:485` `styleSegmentsFromRange`

### 미도달 가능성: 중간 확신

- `client/input/ai_grammary_text_reader.py:617`
  `ActiveWordReader._first_visible_character_range`

### 함수가 아닌 독립/고립 모듈

- `client/config.py`: 앱 진입점에서 import되지 않는다.
- `client/ui/auth_dialog.py`: 현재 UI 흐름에서 import되지 않는다.
- `client/input/hwp_com_diagnostic.py`: 독립 진단 CLI다.
- `client/input/hwp_textfile_diagnostic.py`: 독립 진단 CLI다.
- `server/checkdb.py`: 독립 DB 점검 스크립트다.

### 오탐에서 제외한 항목

- `client/core/ai_client.py:36` `AIClient.has_api_key`: `@property`이며
  `getattr(ai, "has_api_key", False)`로 사용된다.
- `client/ui/mini_overlay.py:433` `TonePrompt._submit`: Qt signal/버튼 콜백 경로다.
- `browser_extension_bridge.py` HTTP handler override
- FastAPI 라우트 함수와 Qt signal 연결 메서드

### 향후 기능용 뼈대로 판단해 제외한 항목

다음 함수는 설정, 요약 결과, 문체 즐겨찾기 및 선택 UI를 나중에 연결하기 위한 얇은
확장용 인터페이스로 보아 미도달 삭제 후보에서 제외했다.

- `client/ui/main_overlay.py:967` `MainOverlay.toggle_settings`
- `client/ui/main_overlay.py:1134` `MainOverlay.replace_summary_result`
- `client/ui/mini_overlay.py:1208` `MiniOverlay.tone_favorite_tones`
- `client/ui/mini_overlay.py:1217` `MiniOverlay.show_choice_prompt`

---

## 프로젝트 간 비교

세 프로젝트는 파일명과 구조가 매우 비슷하지만 호출 그래프는 독립적이다.

- `AI-grammary`: 오래된 HWP 적용 경로와 동기 교정 경로가 많이 남아 있어 미도달 함수가 가장 많다.
- `WA_yunseo`: 브라우저 확장과 DPI/오버레이 구조가 확장됐다. 미구현 AI 메서드와
  확장용 UI wrapper는 삭제 후보에서 제외했다.
- `Whiting-assistant-team-project`: `WA_yunseo`와 유사하지만 AI 호출 구조와 브라우저 확장 세부 구현이 다르다.

동일한 함수명이 여러 프로젝트에 있어도 다른 프로젝트의 참조를 근거로 도달 가능하다고 판단하지 않았다.

## 검증 결과

- 세 프로젝트의 모든 Python 파일은 AST 파싱에 성공했다.
- `python -m compileall` 구문 검사를 통과했다.
- 정적 분석 특성상 `eval`, 문자열 기반 메서드 이름, 외부 플러그인이 임의로 호출하는 공개 API는
  완전히 증명할 수 없다. 이 때문에 COM 관련 1개 메서드는 각 프로젝트에서 중간 확신으로 분리했다.
