# Windows 작업 스케줄러 전용 — 주간 자동 수집 실행 스크립트
# .env에서 FOOD_API_KEY를 읽어 환경변수로 설정한 뒤 main.py 실행, 결과를 logs/에 기록

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$envFile = Join-Path $PSScriptRoot ".env"
if (Test-Path $envFile) {
    Get-Content $envFile | ForEach-Object {
        if ($_ -match '^\s*([^#=]+)=(.*)$') {
            [Environment]::SetEnvironmentVariable($matches[1].Trim(), $matches[2].Trim(), 'Process')
        }
    }
}

$logDir = Join-Path $PSScriptRoot "logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$logFile = Join-Path $logDir ("weekly_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

& "C:\Users\thefuture_brand\AppData\Local\Programs\Python\Python311\python.exe" src/main.py *>> $logFile
