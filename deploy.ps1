# deploy.ps1 - update the server to the latest version from GitHub.
# Run ON THE SERVER from the project folder:  .\deploy.ps1
#
# config.py, *.db, data/runs, uploads etc. are gitignored - git never
# touches them, local server data is preserved.
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

Write-Host "== Current version ==" -ForegroundColor Cyan
git log -1 --oneline

Write-Host "`n== Checking for updates ==" -ForegroundColor Cyan
git fetch origin

$behind = (git rev-list --count HEAD..origin/main).Trim()
if ($behind -eq '0') {
    Write-Host "Already up to date." -ForegroundColor Green
    exit 0
}
Write-Host "New commits: $behind" -ForegroundColor Yellow
git log --oneline HEAD..origin/main

$before = (git rev-parse HEAD).Trim()

git pull --ff-only origin main
if ($LASTEXITCODE -ne 0) {
    Write-Host "`nFast-forward not possible (history rewritten?) - hard reset to origin/main." -ForegroundColor Yellow
    Write-Host "Uncommitted changes to TRACKED files will be lost." -ForegroundColor Yellow
    $ans = Read-Host "Continue? (yes/no)"
    if ($ans -ne 'yes') { exit 1 }
    git reset --hard origin/main
}

$after = (git rev-parse HEAD).Trim()
$changed = git diff --name-only $before $after
if ($changed -match 'requirements\.txt') {
    Write-Host "`n== requirements.txt changed - installing deps ==" -ForegroundColor Cyan
    $py = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
    if (-not (Test-Path $py)) { $py = 'python' }
    & $py -m pip install -r requirements.txt
}

Write-Host "`n== Done. Server version: ==" -ForegroundColor Green
git log -1 --oneline
Write-Host "`nRestart the app to apply changes:" -ForegroundColor Yellow
Write-Host "  - if run manually: close the uvicorn terminal and run .\run-server.ps1" -ForegroundColor Yellow
Write-Host "  - if a service:    Restart-Service <name>" -ForegroundColor Yellow
