# setup-venv.ps1 - create/rebuild an isolated .venv for the app.
# Keeps the app dependencies away from the shared system Python.
# All packages in requirements.txt ship prebuilt wheels for CPython 3.14,
# so no C/Rust compiler is needed.
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

if (Test-Path .venv) {
    Write-Host "Removing existing .venv ..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force .venv
}

Write-Host "Creating .venv (Python 3.14) ..." -ForegroundColor Cyan
py -3.14 -m venv .venv

$py = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
& $py -m pip install --upgrade pip
& $py -m pip install -r requirements.txt

Write-Host "`nDone. Start the app:  .\run-server.ps1" -ForegroundColor Green
& $py -c "import fastapi, uvicorn, sqlalchemy, openpyxl, pandas, docx, selenium, pdfplumber, gspread, pyodbc; print('imports OK')"
