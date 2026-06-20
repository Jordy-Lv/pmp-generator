#!/bin/bash
# Construye el ejecutable de macOS en tu propia máquina (Mac).
# Para el .exe de Windows hace falta una máquina Windows o usar el CI (build.yml).
set -e
cd "$(dirname "$0")"
python3 -m pip install -r requirements.txt -r requirements-build.txt
pyinstaller --clean --noconfirm pmp.spec
chmod +x dist/pmp
cp scripts/Lanzar_PMP.command dist/
chmod +x dist/Lanzar_PMP.command
echo ""
echo "✓ Listo. Para distribuir, comprime dist/pmp + dist/Lanzar_PMP.command"
echo "  La compañera abre con doble clic en 'Lanzar_PMP.command'."
