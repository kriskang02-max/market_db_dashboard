# market_db 대시보드

**이 폴더(`Documents\market_db_dashboard`)가 대시보드·VBA·CSV 작업의 기준 디렉터리입니다.**  
다른 폴더에 복사본을 두지 말고 여기서만 수정하세요.

## 페이지

- **스프레드 분석** — `market_db.csv` 기준금리·라인2/라인3·스프레드(라인3−라인2)
- **펀드 비교분석** — **구간 시작일** 입력 시 **시작일~펀드별 최신일** 연환산 수익률·수탁고 변동만 표시. 비우면 1일·1주·…·YTD. 하단 시계열 차트.
- **Overview** — 제목 페이지(차트·임베드 없음).

### bond_db.xlsx → CSV

`BondDb_ExportToCsv.bas`를 통합문서 VBA로 가져온 뒤 **`ExportBondDbToCsv`**(현재 책) 또는 **`ExportBondDbFromDefaultPath`**(고정 경로 `bond_db.xlsx` 열기) 실행.

- 출력: `bond_db_ktb.csv`(국고채: 헤더 D10:M10, 데이터 D11:M400), `bond_db_msb.csv`(통안채: O10:X10, O11:X100), UTF-8 BOM
- 만기 축약 `25-4`, `24-12` 형식은 CSV에 선행 `'`를 붙여 Excel에서 텍스트로 열리게 함.
- 저장 폴더: 통합문서와 같은 경로(미저장 새 책은 `market_db_dashboard` 폴더로 폴백)

## 실행 방법

1. 이 폴더에서 로컬 서버 실행 (**Overview 저장 공유용**, CSV와 동일 폴더에 `overview_state.json` 기록):
   ```powershell
   cd C:\Users\infomax\Documents\market_db_dashboard
   python serve_dashboard.py
   ```
   (읽기만 필요하면 `python -m http.server 8000` 도 가능하나, Overview **저장**·ETF 개요 **메일 발송**은 `serve_dashboard.py` 필요.)
2. 브라우저: **http://localhost:8000/dashboard.html**  
   (`file://` 로 열어도 되며, CSV 경로는 스크립트에 이 폴더가 박혀 있습니다. Overview 공유 저장은 http + `serve_dashboard.py` 필요.)
3. Overview에서 저장 후 **`push_dashboard_csv.bat`**(또는 `overview_state.json` 포함 git push) → GitHub Pages·다른 PC에서 동일 내용 로드.

## 필요한 파일

- `dashboard.html` — 대시보드 (스프레드 + 펀드 탭)
- `index.html` — 웹 호스팅 시 루트 URL(`/`)에서 `dashboard.html`로 넘김
- `market_db.csv` — 시장 금리 (date, instrument, tenor, yield)
- `fund_db.csv` — 펀드 롱 CSV (VBA `FundDb_ManualCsvExport` 로보내기)
- `issues.csv` — (선택) 이슈 툴팁용
- `term_table_long.csv` — Term Structure 탭 (없으면 해당 탭만 오류)
- `bond_db.xlsx` — 채권 DB 원본 (선택); CSV는 `BondDb_ExportToCsv.bas`로 생성
- `bond_db_ktb.csv`, `bond_db_msb.csv` — 국고채·통안채 표 형태 CSV (`BondDb_ExportToCsv`)
- `overview_state.json` — Overview 메모·전일종가표·통화정책·국고채 발행 계획·주요종목민평 저장 (git push로 기기 간 공유)
- `serve_dashboard.py` — 로컬 http 서버 + Overview `overview_state.json` POST 저장 + ETF 개요 메일 발송 API

### ETF 개요 메일 발송

1. `serve_dashboard.py`로 대시보드를 연 뒤 ETF 개요 **발송** 클릭 → `chanhong.kang@shinhanamc.com` 으로 HTML 표 본문 발송.
2. **Outlook(권장, Windows)** — PowerShell에서 서버 실행 전:
   ```powershell
   $env:EMP_EMAIL_USE_OUTLOOK = "1"
   python serve_dashboard.py
   ```
   (`pip install pywin32` 필요, Outlook 로그인 상태)
3. **SMTP** — `EMP_SMTP_HOST`, `EMP_SMTP_PORT`(기본 587), `EMP_SMTP_USER`, `EMP_SMTP_PASS` 환경 변수 설정.

## 웹에 올리기 (정적 호스팅)

대시보드는 **정적** HTML·CSV·Plotly CDN을 쓰며, `http(s)://` 로 열었을 때 `market_db.csv` 등을 **같은 출처**에서 `fetch`합니다.

1. **올릴 파일** (최소): `index.html`, `dashboard.html`, `market_db.csv`, `fund_db.csv`, `issues.csv`(선택), `term_table_long.csv`(Term 사용 시)
2. **VBA·`.bas` 파일**은 웹 서버에 넣을 필요 없음 (로컬 Excel용).
3. **보안**: `market_db.csv` / `fund_db.csv`에 내부용 데이터가 있으면 **공개 GitHub 저장소**에 그대로 올리지 말고, 비공개 저장소·사내 호스팅·접근 제한(Netlify/Vercel 비밀번호 등)을 검토하세요.
4. **데이터 갱신**: 배포 후에도 CSV만 교체·다시 배포하면 됩니다. 대시보드는 주기적으로 CSV를 다시 불러옵니다.

### 예시 플랫폼

| 플랫폼 | 방법 |
|--------|------|
| **GitHub Pages** | 저장소 Settings → Pages → Branch `main` / 폴더 `/ (root)` 또는 `/docs`. 위 파일들을 커밋 후 `https://<user>.github.io/<repo>/` 접속 |
| **Netlify** | [app.netlify.com](https://app.netlify.com) → Add new site → 이 폴더 드래그 앤 드롭 또는 Git 연동 |
| **Vercel** | [vercel.com](https://vercel.com) → New Project → 폴더 업로드 또는 Git 연동 (Framework: Other) |
| **Cloudflare Pages** | Pages → Create project → 자산 업로드 또는 Git |

배포 후 주소는 **`…/dashboard.html`** 또는 루트 **`…/`** (`index.html` 리다이렉트)로 열면 됩니다.
