# UnHelper

UnHelper는 Coupang Shipments의 반복 업무를 보조하는 Windows 데스크톱 앱입니다. 첫 릴리즈는 `RAW > Milkrun`의 텍스트 다운로드 흐름을 자동화합니다.

## v0.1.0 범위

`데이터 얻기`를 누르면 다음 순서로 동작합니다.

1. Coupang Shipments를 열고 사용자의 로그인이 끝날 때까지 제한 없이 대기합니다.
2. `입고 스케줄 > 밀크런 입고예약 목록`으로 이동합니다.
3. 조회 기간을 어제부터 오늘까지로 지정하고 `안산2` 센터를 선택합니다.
4. 조회 후 텍스트 다운로드를 요청하며, 사유는 `MM.DD-MM.DD` 형식으로 입력합니다.
5. 요청 완료 팝업을 닫고 텍스트 다운로드 내역으로 이동합니다.
6. 유형, 사유, 요청시각, 준비 상태가 모두 일치하는 가장 최신 행의 파일을 내려받습니다.

로그인과 추가 인증은 Chrome 창에서 직접 진행해야 합니다. 로그인 대기는 시간 제한이 없으며, 앱의 `작업 중지` 버튼이나 Chrome 창 닫기로 끝낼 수 있습니다.

다운로드 기본 경로는 `%USERPROFILE%\Downloads\UnHelper`이며 설정에서 변경할 수 있습니다. 실패 시 같은 폴더에 화면 PNG와 HTML 진단 파일을 남깁니다.

## 개발 실행

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe App.py
```

ChromeDriver는 `C:\Users\mrbin\Python\deadline\마감\chromedriver.exe`를 공용 원본으로 사용하고, 빌드할 때 UnHelper에 복사해 번들링합니다. 설치된 앱은 개발 PC의 절대 경로에 의존하지 않습니다.

## 테스트

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
$env:QT_QPA_PLATFORM='offscreen'
.\.venv\Scripts\python.exe App.py --smoke-test
```

실제 Shipments 로그인 이후 구간은 사내 인증이 필요하므로 사용자가 테스트 릴리즈에서 확인해야 합니다.

## 릴리즈 빌드

```powershell
.\.venv\Scripts\python.exe build_release.py un
.\.venv\Scripts\python.exe build_release.py un --installer
```

`release` 폴더에 다음 자산을 만듭니다.

- `UnHelper_manifest.json`
- `UnHelper_patch.zip`: 모든 버전에서 쓸 수 있는 전체 패치
- `UnHelper_delta_patch.zip`: 직전 버전과 파일이 달라질 때만 생성되는 델타 패치
- `UnHelper_Setup.exe`: `--installer` 빌드 시 생성되는 초기 설치본

앱 설정의 Beta 채널을 켜면 GitHub prerelease도 업데이트 대상으로 포함하고, 끄면 최신 정식 릴리즈 복구를 지원합니다. 최초 릴리즈에는 비교할 이전 매니페스트가 없으므로 델타 패치가 생기지 않는 것이 정상입니다.
