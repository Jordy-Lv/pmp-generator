#!/bin/bash
# Doble clic en este archivo abre la Terminal y lanza el PMP Generator.
# Debe estar en la misma carpeta que el binario «pmp».
cd "$(dirname "$0")" || exit 1
./pmp
