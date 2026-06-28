#Requires -Version 5.1
try { chcp 65001 | Out-Null } catch {}
$OutputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

$Root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
Set-Location -LiteralPath $Root

function Write-Step([string]$Msg) { Write-Host $Msg }

Write-Host ""
Write-Host "========================================"
Write-Host "  대시보드 CSV / 상태 Git push"
Write-Host "========================================"
Write-Host "  폴더: $Root"
Write-Host ""

if (-not (Test-Path -LiteralPath (Join-Path $Root ".git") -PathType Container)) {
    Write-Host "[오류] Git 저장소가 아닙니다. market_db_dashboard 폴더에서 실행하세요."
    exit 1
}

$addFiles = @(
    "market_db.csv",
    "fund_db.csv",
    "term_table_long.csv",
    "bond_db_ktb.csv",
    "bond_db_msb.csv",
    "overview_state.json"
)
if (Test-Path -LiteralPath (Join-Path $Root "ETF_Worksheet.xlsx") -PathType Leaf) {
    $addFiles += "ETF_Worksheet.xlsx"
}

& git -C $Root add @addFiles
if ($LASTEXITCODE -ne 0) {
    Write-Host "[오류] git add 실패. 파일 경로를 확인하세요."
    exit 1
}

& git -C $Root diff --cached --quiet
$hasChanges = ($LASTEXITCODE -ne 0)

if (-not $hasChanges) {
    Write-Host "커밋할 변경이 없습니다. CSV와 overview_state가 이전 커밋과 동일합니다."
    $force = Read-Host "변경 없이도 커밋/푸시하겠습니까? (Y/N)"
    if ($force -notmatch '^[Yy]$') {
        Write-Host "취소되었습니다."
        exit 0
    }
    & git -C $Root commit --allow-empty -m "chore: update dashboard CSV, overview_state, and ETF_Worksheet"
} else {
    & git -C $Root commit -m "chore: update dashboard CSV, overview_state, and ETF_Worksheet"
}

if ($LASTEXITCODE -ne 0) {
    Write-Host "[오류] git commit 실패"
    exit 1
}

& git -C $Root push
if ($LASTEXITCODE -ne 0) {
    Write-Host "[오류] git push 실패. 원격 저장소와 로그인을 확인하세요."
    exit 1
}

Write-Host ""
Write-Host "푸시 완료. 잠시 후 GitHub Pages에 반영됩니다."
Write-Host ""

$serve = Join-Path $Root "serve_dashboard.py"
if (-not (Test-Path -LiteralPath $serve -PathType Leaf)) {
    Write-Host "[경고] serve_dashboard.py 가 없어 로컬 서버를 시작하지 않습니다."
    exit 0
}

Write-Host "로컬 대시보드 서버를 시작합니다."
Write-Host "  http://127.0.0.1:8000/dashboard.html"
Write-Host "종료하려면 이 창에서 Ctrl+C 를 누르세요."
Write-Host ""

$env:EMP_EMAIL_USE_OUTLOOK = "1"
Start-Process "http://127.0.0.1:8000/dashboard.html"
& python $serve
if ($LASTEXITCODE -ne 0) {
    Write-Host "[오류] serve_dashboard.py 실행 실패. Python 설치 확인: python --version"
    exit 1
}

exit 0
