# HWPX 서식 보존 작업 인수인계

## 현재 상태
- 프로젝트는 한글 HWP 문서의 HWPX 임시 복사 후 XML 기반 수정 흐름을 지원함.
- 현재 `client/input/output_applier.py`에서 HWP COM 객체를 통해 선택된 한글 문서를 HWPX로 저장하고, 수정 후 `InsertFile`로 다시 삽입하는 방식으로 동작함.
- `_tmp_hwp_word_like/hwpx_rebuilder.py`는 문서 선택 영역을 찾아 XML 텍스트 슬롯을 유지한 채 `replacement_text`를 적용하는 실험용 “Word-like” HWPX 재빌더 구현체임.
- `tests/test_hwpx_document.py`에 새로운 테스트를 추가하여 재빌더가 스타일 정보(`charPrIDRef`)를 보존하고 fragment 범위만 남기는지를 검증함.
- 현재 `python -m unittest tests.test_hwpx_document -v` 실행 결과 15개 테스트가 모두 통과함.

## 작업 내용 요약
1. HWPX 서식 보존 이슈 해결 방향 확인
   - HWPX 임시 복사 후 XML 열람
   - COM 객체로 한글 문서 인식/교정
   - 교정 결과와 원문 비교하여 diff 생성
   - 원문 XML 서식 정보와 diff를 활용해 원문 수정
   - 필요 시 `delete` 및 `insertfile`을 여러 번 수행
2. 코드 위치
   - `client/input/output_applier.py`: HWP COM save/export/InsertFile 흐름
   - `client/input/hwpx_document.py`: 기존 HWPX 텍스트 추출/교체/fragment 생성 구현
   - `_tmp_hwp_word_like/hwpx_rebuilder.py`: 실험적 HWPX 재빌더
   - `tests/test_hwpx_document.py`: HWPX 관련 유닛 테스트
3. 주요 변경 사항
   - 새로운 테스트 `test_create_rebuilt_hwpx_fragment_preserves_style_and_fragment_scope` 추가
   - `create_rebuilt_hwpx_fragment` 경로 검증 및 동작 확인

## 남은 작업 / Codex 인수인계
1. `output_applier.py`에서 HWPX fragment 삽입 로직의 다단계 `delete` + `insertfile` 지원 구현
   - 선택 영역이 여러 개의 분리된 교정 대상일 때 순차적으로 수정 적용 필요
2. 실제 한글 문서에서 `SaveAs(..., "HWPX", option)` 결과와 `GetTextFile`/selection text 간의 불일치 처리 강화
   - `selection`, `saveblock`, 기본 HWPX export 중 가장 적절한 것을 선택하도록 보강
3. 서식 보존 확인
   - `charPrIDRef`, `KeepCharshape`, `KeepParashape`, `KeepStyle` 옵션 유지 여부 점검
   - `InsertFile` 또는 `HAction.Execute("InsertFile")` 호출 시 동작 안정성 확인
4. 추가 테스트 필요
   - 복수 영역/복수 문단 교체 시 `fragment` 범위 유지 검증
   - Shift+Enter, soft line break, 내부 `<hp:lineBreak>` 보존 테스트
   - 실제 한글 `s.hwpx` 샘플 기반 통합 테스트

## 점검 방법
- `python -m unittest tests.test_hwpx_document -v`
- `client/input/runtime_hwp_insertfile_probe.py` 실행으로 실제 HWP `InsertFile` 프로세스 점검
- `client/input/output_applier.py`의 로그 경로 `.logs/hwp_hwpx_blocks`에서 생성된 HWPX 파일 확인

## 참고
- Windows 환경에서 PyWin32(`pythoncom`, `win32com.client`) 필요
- HWP COM 동작은 실제 한글 프로세스 상태에 따라 달라질 수 있음
- 현재 테스트는 HWPX XML 기반 변환 로직 위주로 안정화됨
