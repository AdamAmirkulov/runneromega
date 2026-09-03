# run-server.ps1 - start the app on the server (no auto-reload)
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$py = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $py)) {
    Write-Host "No .venv - using system python. Run setup-venv.ps1 for an isolated env." -ForegroundColor Yellow
    $py = 'python'
}

Write-Host "Starting runneromega on http://0.0.0.0:8000 ..." -ForegroundColor Cyan
& $py -m uvicorn app:app --host 0.0.0.0 --port 8000
