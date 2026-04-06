@echo off
chcp 65001 >nul
setlocal
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

git add market_db.csv fund_db.csv term_table_long.csv
if errorlevel 1 (
  echo git add 실패 ^(파일 경로를 확인하세요^)
  exit /b 1
)

git diff --cached --quiet
if not errorlevel 1 (
  echo 커밋할 변경이 없습니다. ^(세 CSV가 이전 커밋과 동일합니다^)
  exit /b 0
)

git commit -m "chore: update market_db, fund_db, term_table_long CSV"
if errorlevel 1 (
  echo git commit 실패
  exit /b 1
)

git push
if errorlevel 1 (
  echo git push 실패 ^(원격 저장소와 로그인을 확인하세요^)
  exit /b 1
)

echo 푸시 완료. 잠시 후 GitHub Pages에 반영됩니다.
exit /b 0
