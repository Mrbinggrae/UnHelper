# UnHelper

UnHelper는 Coupang Shipments와 WMS의 반복 업무를 보조하는 Windows 데스크톱 앱입니다. `RAW > Milkrun` 다운로드, Excel RAW 시트 반영, 일별 입고 상품 상세와 SKU 무게 분류를 자동화합니다.

## RAW > Milkrun

`데이터 얻기`를 누르면 다음 순서로 동작합니다.

1. Coupang Shipments를 열고 사용자의 로그인이 끝날 때까지 제한 없이 대기합니다.
2. `입고 스케줄 > 밀크런 입고예약 목록`으로 이동합니다.
3. 조회 기간을 어제부터 오늘까지로 지정하고 `안산2` 센터를 선택합니다.
4. 조회 후 텍스트 다운로드를 요청하며, 사유는 `MM.DD-MM.DD` 형식으로 입력합니다.
5. 요청 완료 팝업을 닫고 텍스트 다운로드 내역으로 이동합니다.
6. 유형, 사유, 요청시각, 준비 상태가 모두 일치하는 가장 최신 행의 파일을 내려받습니다.
7. 다운로드 파일 첫 번째 시트의 C열 `입고일`을 확인해 어제 날짜 행을 제외합니다.
8. 설정에서 연결한 Excel의 `Raw_밀크런!C1:P1000` 값을 지운 뒤, 남은 오늘 데이터를 `C1`부터 값으로만 붙여넣고 저장합니다.
9. 같은 로그인 세션으로 `일별 입고 현황`에서 안산2와 오늘 날짜를 조회합니다.
10. `Raw_밀크런`의 P열 발주번호(`/` 앞 번호)를 카드와 연결해 상품 목록을 표시합니다.
11. 저장되지 않은 SKU를 WMS `재고관리 > 상품관리`에서 조회해 상품 무게를 가져옵니다.
12. `상품 무게(g) × (박스 수 ÷ 팔렛트 수) ÷ 1000`으로 1팔렛트 중량을 계산하고, 280kg 이상은 중량, 미만은 경량으로 표시합니다.

로그인과 추가 인증은 Chrome 창에서 직접 진행해야 합니다. 로그인 대기는 시간 제한이 없으며, 앱의 `작업 중지` 버튼이나 Chrome 창 닫기로 끝낼 수 있습니다.

다운로드 기본 경로는 `%USERPROFILE%\Downloads\UnHelper`이며 설정에서 변경할 수 있습니다. 각 실행은 전용 임시 다운로드 폴더를 사용하고, 완성된 지원 형식만 최종 폴더로 이동합니다. 자동화 실패 시 같은 폴더에 화면 PNG와 HTML 진단 파일을 남깁니다.

Milkrun 작업 전 설정의 `Milkrun 데이터를 반영할 Excel 파일`에서 파일명에 `입고스케줄관리`가 포함된 `.xlsx`, `.xlsm`, `.xlsb`, `.xls` 파일을 연결해야 합니다. 대상에는 `Raw_밀크런` 시트가 있어야 하며 읽기 전용이면 안 됩니다. 이 기능은 Windows용 Microsoft Excel 데스크톱 앱이 설치된 PC에서 동작합니다. 다운로드 원본은 Excel 형식과 CSV/TSV/TXT를 지원하고, `C1:P1000`을 넘는 데이터는 기존 값을 지우기 전에 중단합니다. 값 전용 대입을 사용하므로 대상 시트의 기존 서식은 유지됩니다.

설정의 WMS ID/Password는 저장되지 않은 SKU의 무게를 자동 조회할 때 사용합니다. Password는 Windows DPAPI로 현재 Windows 사용자만 복호화할 수 있게 보호합니다. 이미 저장된 SKU 무게는 다시 조회하지 않고 현재 행의 박스/팔렛트 비율로 분류를 다시 계산합니다. 표의 분류 버튼 하나를 눌러 경량·중량·고단을 순환 변경할 수 있으며, 수동 분류는 자동 계산보다 우선합니다.

저장 상품 목록은 설정에서 경량·중량·고단별로 확인할 수 있습니다. JSON 가져오기/내보내기에는 SKU ID, 상품명, 무게, 분류와 측정·계산 정보가 포함되며 WMS 계정은 포함되지 않습니다. 상품명은 실제 줄바꿈만 공백으로 바꾸고 `/`를 포함한 나머지 문자열을 그대로 저장하며 표에서도 자동 줄바꿈하지 않습니다. 저장 파일은 `%LOCALAPPDATA%\Mrbinggrae\UnHelper`에 유지되며, 파일이 손상되면 사용자 확인 후 원본을 백업하고 빈 목록으로 복구할 수 있습니다.

작업 오류 팝업에서는 상세 traceback과 최근 로그를 확인하고, 민감정보를 가린 신고 내용을 복사하거나 로그인된 브라우저에서 UnHelper GitHub 이슈 작성 화면을 열 수 있습니다. 설정의 `업데이트 내역` 버튼에서는 설치된 버전의 변경 기록을 확인할 수 있습니다.

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

실제 Shipments와 WMS 로그인 이후 구간은 사내 인증이 필요하므로 사용자가 릴리즈 환경에서 확인해야 합니다.

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
