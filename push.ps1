# push.ps1 — быстрый коммит+пуш с ЛОКАЛЬНОЙ машины.
#   .\push.ps1 "что изменил"
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$msg = $args -join ' '
if (-not $msg) { $msg = Read-Host "Сообщение коммита" }
if (-not $msg) { Write-Host "Пустое сообщение — отмена." -ForegroundColor Red; exit 1 }

git add -A
git status --short
$staged = git diff --cached --name-only
if (-not $staged) { Write-Host "Нечего коммитить." -ForegroundColor Yellow; exit 0 }

git commit -m $msg
git push origin master
Write-Host "`nЗапушено. На сервере выполните: .\deploy.ps1" -ForegroundColor Green
