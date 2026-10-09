# Gera o ControleFaltas.exe portátil (um único arquivo, sem instalar).
# Uso: clique com botão direito > "Executar com PowerShell" ou rode:
#   powershell -ExecutionPolicy Bypass -File build_exe.ps1

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

Write-Host "1/3 Instalando dependências..."
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

Write-Host "2/3 Gerando o .exe (pode levar alguns minutos)..."
python -m PyInstaller --noconfirm --onefile --windowed `
  --name "ControleFaltas" `
  --exclude-module matplotlib --exclude-module scipy --exclude-module pandas `
  app.py

Write-Host "3/3 Organizando..."
$exe = Join-Path $PSScriptRoot "dist\ControleFaltas.exe"
if (Test-Path -LiteralPath $exe) {
  Copy-Item -LiteralPath $exe -Destination (Join-Path $PSScriptRoot "ControleFaltas.exe") -Force
  Write-Host ""
  Write-Host "PRONTO! Use: $PSScriptRoot\ControleFaltas.exe"
  Write-Host "Deixe o .exe na MESMA PASTA da planilha .xlsx."
} else {
  Write-Host "FALHOU: exe não encontrado em dist."
  exit 1
}
