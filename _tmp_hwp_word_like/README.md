# 임시 실험: HWPX Word식 줄/슬롯 재구성

목적: 기존 `client/input/hwpx_document.py`를 바로 갈아엎지 않고, Word 방식에 가까운 구조를 HWPX에서 실험한다.

## 핵심 아이디어

- HWPX 선택 영역을 그대로 저장한다.
- XML의 기존 `<hp:t>` 텍스트 슬롯과 `<hp:lineBreak/>` 위치를 읽는다.
- 교정 결과가 한 줄로 와도 기존 줄 구조에 맞춰 다시 분배한다.
- 글자색/굵기/기울임 같은 서식은 XML의 기존 슬롯 구조를 유지해서 보존을 시도한다.
- COM으로 서식을 다시 칠하지 않는다.

## 앱에서 실험 활성화

Windows CMD:

```bat
set WA_HWP_EXPERIMENTAL_REBUILDER=1
cd "D:\Hong\26 Git\0528\Writing-Assistant"
py main.py
```

PowerShell:

```powershell
$env:WA_HWP_EXPERIMENTAL_REBUILDER="1"
cd "D:\Hong\26 Git\0528\Writing-Assistant"
py main.py
```

비활성화:

```bat
set WA_HWP_EXPERIMENTAL_REBUILDER=
```

## 되돌리기

1. 환경변수 `WA_HWP_EXPERIMENTAL_REBUILDER`를 끈다.
2. 필요하면 `_tmp_hwp_word_like` 폴더만 삭제한다.
3. 메인 코드 영향은 `output_applier.py`의 환경변수 분기뿐이다.

## 우선 테스트

1. 일반 문장 + Shift+Enter 줄 + 일반 문장 같이 선택
2. 색상/기울임/밑줄이 섞인 문장 선택
3. 전체 선택
4. 표 내부는 후순위
