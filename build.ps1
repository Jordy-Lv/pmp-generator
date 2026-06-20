# Construye el ejecutable de Windows en tu propia maquina (Windows).
# Para el binario de Mac hace falta un Mac o usar el CI (build.yml).
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
python -m pip install -r requirements.txt -r requirements-build.txt
pyinstaller --clean --noconfirm pmp.spec
Write-Host ""
Write-Host "OK. El ejecutable esta en dist\PMP.exe"
Write-Host "Para distribuir, comprime dist\PMP.exe y pasalo. Doble clic para abrir."
