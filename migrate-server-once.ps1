# migrate-server-once.ps1
# РАЗОВАЯ миграция сервера на чистую ветку main.
# Запускать ОДИН РАЗ на сервере из папки проекта:
#     powershell -ExecutionPolicy Bypass -File .\migrate-server-once.ps1
#
# Что делает: бэкапит config.py / базы / ключи, переключает git на чистую
# main, чистит отслеживаемый мусор, возвращает данные, пересоздаёт .venv.
# Папки data/, uploads/, logs/ и сами базы не трогаются (они в .gitignore).

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$backup = Join-Path (Split-Path $PSScriptRoot -Parent) '_runneromega_backup'
Write-Host "== 1. Бэкап локальных данных сервера -> $backup ==" -ForegroundColor Cyan
New-Item -ItemType Directory -Force -Path $backup | Out-Null
foreach ($p in @('scripts\config.py','scripts\service_account.json')) {
    if (Test-Path $p) { Copy-Item $p $backup -Force; Write-Host "  сохранён $p" }
}
Get-ChildItem -Path . -Filter *.db      -ErrorAction SilentlyContinue | Copy-Item -Destination $backup -Force
Get-ChildItem -Path . -Filter *.sqlite  -ErrorAction SilentlyContinue | Copy-Item -Destination $backup -Force
if (Test-Path 'data\ais_oip.db') { Copy-Item 'data\ais_oip.db' $backup -Force }
git diff HEAD -- '*.py' '*.txt' | Out-File -Encoding utf8 (Join-Path $backup 'server_tracked_edits.patch')
Write-Host "  диф серверных правок кода -> $backup\server_tracked_edits.patch"

Write-Host "`n== 2. Переключение на чистую ветку main ==" -ForegroundColor Cyan
git fetch origin --prune
git checkout -f -B main origin/main
git reset --hard origin/main
git branch -D master 2>$null

Write-Host "`n== 3. Удаление отслеживаемого ранее мусора (данные не трогаются) ==" -ForegroundColor Cyan
git clean -fd            # без -x: файлы из .gitignore (config.py, .venv, *.db, data/) остаются

Write-Host "`n== 4. Возврат config.py / ключей / баз ==" -ForegroundColor Cyan
if (Test-Path (Join-Path $backup 'config.py'))            { Copy-Item (Join-Path $backup 'config.py')            'scripts\' -Force }
if (Test-Path (Join-Path $backup 'service_account.json')) { Copy-Item (Join-Path $backup 'service_account.json') 'scripts\' -Force }
Get-ChildItem -Path $backup -Filter *.db     -ErrorAction SilentlyContinue | Copy-Item -Destination . -Force
Get-ChildItem -Path $backup -Filter *.sqlite -ErrorAction SilentlyContinue | Copy-Item -Destination . -Force
if (Test-Path (Join-Path $backup 'ais_oip.db')) { Copy-Item (Join-Path $backup 'ais_oip.db') 'data\' -Force }

Write-Host "`n== 5. Проверка ==" -ForegroundColor Cyan
git log -1 --oneline
"config.py:        " + (Test-Path 'scripts\config.py')
"users.db:         " + (Test-Path 'users.db')
"my_database.sqlite:" + (Test-Path 'my_database.sqlite')
$dirty = (git status --porcelain | Measure-Object).Count
"незакоммич. (должно быть 0): $dirty"

Write-Host "`n== 6. Пересоздание .venv ==" -ForegroundColor Cyan
if (Test-Path .venv) { Remove-Item -Recurse -Force .venv }
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip --quiet
.\.venv\Scripts\python.exe -m pip install -r requirements.txt

Write-Host "`n== ГОТОВО ==" -ForegroundColor Green
Write-Host "Запусти приложение:  powershell -ExecutionPolicy Bypass -File .\run-server.ps1" -ForegroundColor Yellow
Write-Host "Бэкап данных лежит в: $backup  (удали после проверки, что всё работает)" -ForegroundColor Yellow
