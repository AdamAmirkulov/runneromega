# run-server.ps1 — запуск приложения на сервере (без auto-reload)
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$py = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $py)) { $py = 'python' }

Write-Host "Запуск runneromega на http://0.0.0.0:8000 ..." -ForegroundColor Cyan
& $py -m uvicorn app:app --host 0.0.0.0 --port 8000
