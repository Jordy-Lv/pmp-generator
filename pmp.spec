# -*- mode: python ; coding: utf-8 -*-
"""
Receta de PyInstaller para empaquetar el PMP Generator como un único ejecutable
autocontenido (incluye Python y todas las dependencias).

  Windows → dist/PMP.exe        (doble clic abre la consola)
  macOS   → dist/pmp            (se abre con scripts/Lanzar_PMP.command)

El mismo spec sirve en ambos sistemas; PyInstaller NO cross-compila, así que el
.exe se construye en Windows y el binario de Mac en macOS (ver build.yml / CI).
"""
import sys
from PyInstaller.utils.hooks import collect_all

# openpyxl, prompt_toolkit y questionary cargan datos/plantillas por ruta en
# tiempo de ejecución; `collect_all` los empaqueta para que el binario no falle
# al guardar el Excel ni al dibujar los menús.
_datas, _binaries, _hidden = [], [], []
for _pkg in ("openpyxl", "et_xmlfile", "prompt_toolkit", "questionary", "rich"):
    d, b, h = collect_all(_pkg)
    _datas += d
    _binaries += b
    _hidden += h

a = Analysis(
    ["pmp_generator.py"],
    pathex=[],
    binaries=_binaries,
    datas=_datas,
    hiddenimports=_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter"],          # la GUI se eliminó; excluirla aligera el binario
    noarchive=False,
)

pyz = PYZ(a.pure)

_name = "PMP" if sys.platform.startswith("win") else "pmp"

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=True,                  # CLI interactivo: necesita consola con TTY real
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,              # usa la arquitectura nativa del runner (arm64 / x86_64)
    codesign_identity=None,
    entitlements_file=None,
)
