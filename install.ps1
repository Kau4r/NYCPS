# NYCPS environment setup (Windows PowerShell)
# Run from this folder:  .\install.ps1

$ErrorActionPreference = 'Stop'

if (-not (Test-Path .\.venv)) {
    Write-Host "Creating virtual environment (.venv)..."
    python -m venv .venv
}

Write-Host "Activating virtual environment..."
. .\.venv\Scripts\Activate.ps1

Write-Host "Upgrading pip..."
python -m pip install --upgrade pip

Write-Host "Installing requirements..."
pip install -r requirements.txt

Write-Host "Registering Jupyter kernel..."
python -m ipykernel install --user --name nycps --display-name "Python (NYCPS)"

Write-Host ""
Write-Host "Done. To use:"
Write-Host "  .\.venv\Scripts\Activate.ps1"
Write-Host "  jupyter notebook NYC_BA_PROJECT.ipynb"
