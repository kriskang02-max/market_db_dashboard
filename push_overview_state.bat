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
  echo Git 저장소가 아닙니다.
  exit /b 1
)

git add overview_state.json
if errorlevel 1 exit /b 1

git diff --cached --quiet
if not errorlevel 1 (
  echo overview_state.json 변경 없음.
  exit /b 0
)

git commit -m "chore: update overview_state.json"
if errorlevel 1 exit /b 1

git push
if errorlevel 1 exit /b 1

echo Overview 공유 데이터 푸시 완료.
exit /b 0
