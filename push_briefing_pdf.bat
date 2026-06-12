@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

set "REPORT_DIR=C:\Users\infomax\Documents\Report"
set "DASHBOARD_DIR=%~dp0"
if "%DASHBOARD_DIR:~-1%"=="\" set "DASHBOARD_DIR=%DASHBOARD_DIR:~0,-1%"

for /f %%i in ('powershell -NoProfile -Command "(Get-Date).ToString('yyyyMMdd')"') do set "TAG=%%i"
set "SRC=%REPORT_DIR%\bond_morning_briefing_%TAG%.pdf"
set "DST_DIR=%DASHBOARD_DIR%\reports"
set "DST=%DST_DIR%\bond_morning_briefing_%TAG%.pdf"

echo.
echo ========================================
echo   모닝브리핑 PDF 대시보드 업로드
echo ========================================
echo   날짜: %TAG%
echo   원본: %SRC%
echo   대상: %DST%
echo.

if not exist "%SRC%" (
  echo [오류] 당일 PDF가 없습니다.
  echo        먼저 run_all.bat 또는 run_all_and_upload.bat 으로 PDF를 생성하세요.
  exit /b 1
)

if not exist "%DASHBOARD_DIR%\.git" (
  echo [오류] Git 저장소가 아닙니다: %DASHBOARD_DIR%
  exit /b 1
)

if not exist "%DST_DIR%" mkdir "%DST_DIR%"
copy /Y "%SRC%" "%DST%" >nul
if errorlevel 1 (
  echo [오류] PDF 복사 실패
  exit /b 1
)

pushd "%DASHBOARD_DIR%" || exit /b 1
git add "reports\bond_morning_briefing_%TAG%.pdf"
if errorlevel 1 (
  echo [오류] git add 실패
  popd
  exit /b 1
)

git diff --cached --quiet
if not errorlevel 1 (
  echo 커밋할 변경이 없습니다. ^(이미 동일한 PDF가 업로드됨^)
  popd
  exit /b 0
)

git commit -m "chore: upload bond morning briefing %TAG%"
if errorlevel 1 (
  echo [오류] git commit 실패
  popd
  exit /b 1
)

git push
if errorlevel 1 (
  echo [오류] git push 실패
  popd
  exit /b 1
)

popd
echo.
echo 업로드 완료. 잠시 후 대시보드 Report 페이지에 반영됩니다.
exit /b 0
