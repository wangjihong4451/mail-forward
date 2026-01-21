$ErrorActionPreference = "Stop"

if (!(Test-Path ".\.env")) {
  Write-Host "Missing .env. Please copy config.example.env to .env and fill it first." -ForegroundColor Yellow
  exit 1
}

python -m pip install -r requirements.txt
python .\forwarder.py

