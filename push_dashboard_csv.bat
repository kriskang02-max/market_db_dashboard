@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
call :main
set "EC=%errorlevel%"
echo.
if not "%EC%"=="0" echo [종료 코드 %EC%]
pause
exit /b %EC%

:main
cd /d "%~dp0" || (
  echo 배치 파일 위치로 이동할 수 없습니다.
  exit /b 1
)
if not exist ".git" (
  echo Git 저장소가 아닙니다. push_dashboard_csv.bat 을 market_db_dashboard 폴더에 두세요.
  exit /b 1
)

git add market_db.csv fund_db.csv term_table_long.csv bond_db_ktb.csv bond_db_msb.csv overview_state.json ETF_Worksheet.xlsx
if errorlevel 1 (
  echo git add 실패 ^(파일 경로를 확인하세요^)
  exit /b 1
)

git diff --cached --quiet
if not errorlevel 1 (
  echo 커밋할 변경이 없습니다. ^(CSV·overview_state.json·ETF_Worksheet.xlsx 가 이전 커밋과 동일합니다^)
  set /p FORCE_PUSH=변경사항이 없어도 커밋/푸시하겠습니까? ^(Y/N^): 
  if /i not "!FORCE_PUSH!"=="Y" (
    echo 취소되었습니다.
    exit /b 0
  )
  git commit --allow-empty -m "chore: update dashboard CSV, overview_state, and ETF_Worksheet"
  if errorlevel 1 (
    echo git commit 실패
    exit /b 1
  )
  goto :push
)

git commit -m "chore: update dashboard CSV, overview_state, and ETF_Worksheet"
if errorlevel 1 (
  echo git commit 실패
  exit /b 1
)

:push
git push
if errorlevel 1 (
  echo git push 실패 ^(원격 저장소와 로그인을 확인하세요^)
  exit /b 1
)

echo 푸시 완료. 잠시 후 GitHub Pages에 반영됩니다.
echo.
call :start_dashboard_server
exit /b %errorlevel%

:start_dashboard_server
if not exist "serve_dashboard.py" (
  echo [경고] serve_dashboard.py 가 없어 로컬 서버를 시작하지 않습니다.
  exit /b 0
)

echo 로컬 대시보드 서버를 시작합니다 ^(ETF 개요 발송, Outlook^).
echo   http://127.0.0.1:8000/dashboard.html
echo 종료하려면 이 창에서 Ctrl+C 를 누르세요.
echo.

set "EMP_EMAIL_USE_OUTLOOK=1"
start "" "http://127.0.0.1:8000/dashboard.html"
python serve_dashboard.py
if errorlevel 1 (
  echo serve_dashboard.py 실행 실패 ^(Python 설치: python --version^)
  exit /b 1
)
exit /b 0
