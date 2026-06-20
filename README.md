# PMP Generator — Célula 3

Herramienta que automatiza la **rotación semanal de clientes PMP** del equipo de
la Célula 3 (Mateo → Heiner → Estefania → Mateo).

La fecha es **automática**: detecta la última semana registrada en el Control y
genera la **siguiente** (solo hay que confirmar). A partir de los dos Excel reales:

- **`Control_Gestion_PMP.xlsx`** (pestaña `2026`): **copia el bloque semanal tal
  cual** debajo del último y cambia solo lo necesario: las fechas a la semana
  siguiente, el encargado de cada cliente de Célula 3 (giro
  `Heiner→Mateo→Estefania`), deja **en blanco** el consultor de los clientes que no
  son de Célula 3, y pone **"FESTIVO"** únicamente en los días que de verdad lo son
  (festivos de Colombia, calculados). Los clientes se ubican **por nombre**, así
  aguanta que la tabla-resumen de Célula 3 se haya movido de columna.
- **`Matriz Unificada de Recursos.xlsx`** (pestaña `ROTACION DISPO CELULA 3 `):
  detecta hasta qué viernes está cubierta y **extiende** las semanas N2/N3
  viernes-viernes que falten. Si no está la Matriz, la disponibilidad se calcula
  por rotación.
- **Cuadro resumen** (`PMP_Semana.xlsx`): Excel formateado para compartir al equipo.
  Es **opcional** (se pregunta al final).

El Control y el cuadro resumen salen de **una sola rotación**, así que coinciden.
**Se trabaja sobre un único archivo:** el Control y la Matriz se **sobrescriben in
situ** (no se generan copias nuevas cada semana). El guardado es **atómico** (se
escribe a un temporal y se renombra), así un fallo a mitad nunca deja el archivo a
medias. El cuadro resumen también usa un nombre fijo (`PMP_Semana.xlsx`) que se
sobrescribe. La herramienta detecta sola la última semana del Control y genera la
siguiente, encadenando sobre el mismo archivo.

> **Para la persona que solo va a usar la herramienta:** ver **[GUIA_USO.md](GUIA_USO.md)**
> (instalación por doble clic, sin tocar nada técnico).

---

## Uso (resumen)

Al abrir la herramienta aparece un menú con flechas:

1. **🌙 Consultar disponibilidad nocturna** — N2/N3 de una semana.
2. **📋 Generar semana siguiente** — flujo completo: detecta y propone la fecha,
   vista previa de la rotación, confirmación, y actualización del Control + Matriz
   (y, opcionalmente, el cuadro resumen).
3. **⚙ Configurar rutas de archivos** — auto-detección o ruta manual.

La herramienta **auto-detecta** los dos Excel en `~/Downloads`, `~/Desktop` y
`~/Documents` (y un nivel de subcarpetas). Si hay un único candidato lo usa; si
hay varios, pregunta; si no hay ninguno, pide la ruta. Las rutas se recuerdan en
`~/.pmp_celula3.json`.

---

## Instalación para usuarios finales (recomendado)

No requiere instalar Python ni nada: se descarga un ejecutable y se abre con
**doble clic**. Los pasos detallados están en **[GUIA_USO.md](GUIA_USO.md)**. En
corto:

1. Descargar el `.zip` del sistema correspondiente desde la **Release** del
   repositorio:
   - Windows → `PMP-Windows.zip`
   - Mac con Apple Silicon (M1/M2/M3…) → `PMP-macOS-AppleSilicon.zip`
   - Mac con Intel → `PMP-macOS-Intel.zip`
2. Descomprimir.
3. Abrir: en Windows `PMP.exe`; en Mac `Lanzar PMP.command`.
   (La primera vez, el sistema pide una confirmación de seguridad — ver la guía.)

---

## Construir el ejecutable

Los ejecutables se generan solos en **GitHub Actions** al publicar un tag de
versión (`git tag v1.0.0 && git push --tags`): el workflow
[`.github/workflows/build.yml`](.github/workflows/build.yml) construye Windows,
macOS Apple Silicon y macOS Intel, y los adjunta a la Release.

Para construir a mano en tu propia máquina:

```bash
./build.sh      # macOS  → dist/pmp + scripts/Lanzar_PMP.command
```
```powershell
.\build.ps1     # Windows → dist\PMP.exe
```

(PyInstaller no cross-compila: cada sistema produce solo su propio ejecutable;
por eso el CI cubre los tres.)

---

## Modo desarrollador (ejecutar el script directamente)

Requiere Python 3.8+ con `openpyxl`, `rich` y `questionary`. Si falta alguno, el
script intenta instalarlo al arrancar.

```bash
python3 -m pip install -r requirements.txt
python3 pmp_generator.py
python3 pmp_generator.py --self-test   # pruebas rápidas
```

---

## Configuración (`~/.pmp_celula3.json`)

| Clave | Significado |
|-------|-------------|
| `ruta_pmp` / `ruta_matriz` | Rutas a los dos Excel (auto-detectadas). |
| `rotacion_pmp` | Orden de rotación de los 3 ingenieros. |
| `clientes_celula3` | Clientes que rota la Célula 3 (se buscan por nombre al leer). |
| `rotacion_n2` / `fecha_base_n2` | Rotación N2 y fecha base para el cálculo de respaldo. |
| `rotacion_n3` | Orden de respaldo de escalamiento N3. |
| `horario_tarde` | Clientes que se atienden en jornada de tarde. |
| `clientes_largos` | Clientes marcados como "PMP largo". |
