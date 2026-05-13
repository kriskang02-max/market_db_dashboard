# market_db 대시보드

**이 폴더(`Documents\market_db_dashboard`)가 대시보드·VBA·CSV 작업의 기준 디렉터리입니다.**  
다른 폴더에 복사본을 두지 말고 여기서만 수정하세요.

## 페이지

- **스프레드 분석** — `market_db.csv` 기준금리·라인2/라인3·스프레드(라인3−라인2)
- **펀드 비교분석** — **구간 시작일** 입력 시 **시작일~펀드별 최신일** 연환산 수익률·수탁고 변동만 표시. 비우면 1일·1주·…·YTD. 하단 시계열 차트.
- **Overview** — 제목 페이지(차트·임베드 없음).

## 실행 방법

1. 이 폴더에서 로컬 서버 실행:
   ```powershell
   cd C:\Users\infomax\Documents\market_db_dashboard
   python -m http.server 8000
   ```
2. 브라우저: **http://localhost:8000/dashboard.html**  
   (`file://` 로 열어도 되며, CSV 경로는 스크립트에 이 폴더가 박혀 있습니다.)

## 필요한 파일

- `dashboard.html` — 대시보드 (스프레드 + 펀드 탭)
- `index.html` — 웹 호스팅 시 루트 URL(`/`)에서 `dashboard.html`로 넘김
- `market_db.csv` — 시장 금리 (date, instrument, tenor, yield)
- `fund_db.csv` — 펀드 롱 CSV (VBA `FundDb_ManualCsvExport` 로보내기)
- `issues.csv` — (선택) 이슈 툴팁용
- `term_table_long.csv` — Term Structure 탭 (없으면 해당 탭만 오류)

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
