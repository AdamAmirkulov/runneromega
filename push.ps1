# push.ps1 - quick commit + push from the LOCAL machine.
#   .\push.ps1 "what I changed"
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$msg = $args -join ' '
if (-not $msg) { $msg = Read-Host "Commit message" }
if (-not $msg) { Write-Host "Empty message - abort." -ForegroundColor Red; exit 1 }

git add -A
git status --short
$staged = git diff --cached --name-only
if (-not $staged) { Write-Host "Nothing to commit." -ForegroundColor Yellow; exit 0 }

git commit -m $msg
git push origin main
Write-Host "`nPushed. On the server run: .\deploy.ps1" -ForegroundColor Green
