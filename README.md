# market_db 대시보드

market_db.csv 기준금리·라인2/라인3 금리 및 스프레드(라인3−라인2) 조회용 대시보드입니다.

## 실행 방법

1. 이 폴더에서 로컬 서버 실행:
   ```powershell
   python -m http.server 8000
   ```
2. 브라우저에서 열기: **http://localhost:8000/dashboard.html**

## 필요한 파일

- `dashboard.html` — 대시보드 페이지
- `market_db.csv` — 데이터 (date, instrument, tenor, yield)
