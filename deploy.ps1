# deploy.ps1 — обновить сервер до последней версии из GitHub.
# Запускать НА СЕРВЕРЕ из папки проекта:  .\deploy.ps1
#
# config.py, *.db, папки data/runs, uploads и т.п. в .gitignore —
# git их не трогает, локальные данные сервера сохраняются.
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

Write-Host "== Текущая версия ==" -ForegroundColor Cyan
git log -1 --oneline

Write-Host "`n== Проверяю обновления ==" -ForegroundColor Cyan
git fetch origin

$behind = (git rev-list --count HEAD..origin/master).Trim()
if ($behind -eq '0') {
    Write-Host "Уже актуально, обновлять нечего." -ForegroundColor Green
    exit 0
}
Write-Host "Новых коммитов: $behind" -ForegroundColor Yellow
git log --oneline HEAD..origin/master

$before = (git rev-parse HEAD).Trim()

git pull --ff-only origin master
if ($LASTEXITCODE -ne 0) {
    Write-Host "`nfast-forward невозможен (история переписана?) — жёсткий сброс на origin/master." -ForegroundColor Yellow
    Write-Host "Незакоммиченные изменения в отслеживаемых файлах будут потеряны." -ForegroundColor Yellow
    $ans = Read-Host "Продолжить? (yes/no)"
    if ($ans -ne 'yes') { exit 1 }
    git reset --hard origin/master
}

# Обновить зависимости, если менялся requirements.txt
$after = (git rev-parse HEAD).Trim()
$changed = git diff --name-only $before $after
if ($changed -match 'requirements\.txt') {
    Write-Host "`n== requirements.txt изменился — ставлю зависимости ==" -ForegroundColor Cyan
    $py = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
    if (-not (Test-Path $py)) { $py = 'python' }
    & $py -m pip install -r requirements.txt
}

Write-Host "`n== Готово. Версия на сервере: ==" -ForegroundColor Green
git log -1 --oneline
Write-Host "`nПерезапустите приложение, чтобы изменения вступили в силу:" -ForegroundColor Yellow
Write-Host "  - если запущено вручную: закрыть терминал uvicorn и запустить .\run-server.ps1" -ForegroundColor Yellow
Write-Host "  - если служба: Restart-Service <имя>" -ForegroundColor Yellow
