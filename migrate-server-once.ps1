# migrate-server-once.ps1
# ONE-TIME server migration to the clean 'main' branch.
# Run ONCE on the server from the project folder:
#     powershell -ExecutionPolicy Bypass -File .\migrate-server-once.ps1
#
# Backs up config.py / databases / keys, switches git to the clean main,
# cleans previously-tracked junk, restores the data, rebuilds .venv.
# data/, uploads/, logs/ and the DB files are left untouched (gitignored).

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$backup = Join-Path (Split-Path $PSScriptRoot -Parent) '_runneromega_backup'
Write-Host "== 1. Backup local server data -> $backup ==" -ForegroundColor Cyan
New-Item -ItemType Directory -Force -Path $backup | Out-Null
foreach ($p in @('scripts\config.py','scripts\service_account.json')) {
    if (Test-Path $p) { Copy-Item $p $backup -Force; Write-Host "  saved $p" }
}
Get-ChildItem -Path . -Filter *.db     -ErrorAction SilentlyContinue | Copy-Item -Destination $backup -Force
Get-ChildItem -Path . -Filter *.sqlite -ErrorAction SilentlyContinue | Copy-Item -Destination $backup -Force
if (Test-Path 'data\ais_oip.db') { Copy-Item 'data\ais_oip.db' $backup -Force }
git diff HEAD -- '*.py' '*.txt' | Out-File -Encoding utf8 (Join-Path $backup 'server_tracked_edits.patch')
Write-Host "  server code edits diff -> $backup\server_tracked_edits.patch"

Write-Host "`n== 2. Switch to clean branch main ==" -ForegroundColor Cyan
git fetch origin --prune
git checkout -f -B main origin/main
git reset --hard origin/main
git branch -D master 2>$null

Write-Host "`n== 3. Remove previously-tracked junk (data is kept) ==" -ForegroundColor Cyan
git clean -fd    # no -x: gitignored files (config.py, .venv, *.db, data/) stay

Write-Host "`n== 4. Restore config.py / keys / databases ==" -ForegroundColor Cyan
if (Test-Path (Join-Path $backup 'config.py'))            { Copy-Item (Join-Path $backup 'config.py')            'scripts\' -Force }
if (Test-Path (Join-Path $backup 'service_account.json')) { Copy-Item (Join-Path $backup 'service_account.json') 'scripts\' -Force }
Get-ChildItem -Path $backup -Filter *.db     -ErrorAction SilentlyContinue | Copy-Item -Destination . -Force
Get-ChildItem -Path $backup -Filter *.sqlite -ErrorAction SilentlyContinue | Copy-Item -Destination . -Force
if (Test-Path (Join-Path $backup 'ais_oip.db')) { Copy-Item (Join-Path $backup 'ais_oip.db') 'data\' -Force }

Write-Host "`n== 5. Check ==" -ForegroundColor Cyan
git log -1 --oneline
Write-Host ("config.py present:         " + (Test-Path 'scripts\config.py'))
Write-Host ("users.db present:          " + (Test-Path 'users.db'))
Write-Host ("my_database.sqlite present:" + (Test-Path 'my_database.sqlite'))
$dirty = (git status --porcelain | Measure-Object).Count
Write-Host ("uncommitted (expect 0):    " + $dirty)

Write-Host "`n== 6. Rebuild .venv ==" -ForegroundColor Cyan
if (Test-Path .venv) { Remove-Item -Recurse -Force .venv }
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip --quiet
.\.venv\Scripts\python.exe -m pip install -r requirements.txt

Write-Host "`n== DONE ==" -ForegroundColor Green
Write-Host "Start the app:  powershell -ExecutionPolicy Bypass -File .\run-server.ps1" -ForegroundColor Yellow
Write-Host "Backup is at:   $backup  (delete it once everything works)" -ForegroundColor Yellow
