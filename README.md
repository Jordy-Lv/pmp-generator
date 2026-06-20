# PMP Generator — Célula 3

Herramienta de terminal que automatiza la **rotación semanal de clientes PMP**
del equipo de la Célula 3 (Mateo → Heiner → Estefania → Mateo).

Lee dos Excel reales y genera el cuadro de la semana siguiente, ya formateado:

- **`Control_Gestion_PMP.xlsx`** — pestaña `2026`: clientes asignados por ingeniero.
- **`Matriz Unificada de Recursos.xlsx`** — pestaña `ROTACION DISPO CELULA 3 `:
  disponibilidad nocturna N2 / escalamiento N3. Si no está, se calcula por rotación.

## Uso

```bash
pmp
```

(El comando `pmp` vive en `~/.local/bin` y usa el entorno `~/.venvs/pmp`.)

También se puede lanzar directamente:

```bash
~/.venvs/pmp/bin/python ~/Proyectos/tools/pmp-generator/pmp_generator.py
```

### GUI

Para abrir la interfaz gráfica:

```bash
pmp --gui
```

O directamente:

```bash
~/.venvs/pmp/bin/python ~/Proyectos/tools/pmp-generator/pmp_generator.py --gui
```

Si se abre el script sin una terminal interactiva, intentará abrir la GUI automáticamente.

En la GUI, **Actualizar copias** crea copias de los archivos reales:

- En `Control_Gestion_PMP`, copia el último bloque semanal debajo, conserva el formato, actualiza fechas y rota los consultores de Célula 3.
- En `Matriz Unificada de Recursos`, asegura la fila de disponibilidad viernes-viernes correspondiente; si ya existe, no rehace la tabla.
- La fecha sugerida se calcula desde el último bloque encontrado en el Control PMP seleccionado.

### Menú

1. **🌙 Consultar disponibilidad nocturna** — N2/N3 de una semana.
2. **📋 Generar cuadro PMP semanal** — flujo completo con vista previa.
3. **⚙ Configurar rutas de archivos** — auto-detección o ruta manual.

## Qué hace especial a esta versión

- **Menús con flechas** (`questionary`): selección, checkbox de ausentes, confirmaciones.
- **Salida formateada** (`rich`): paneles, tablas y barras de progreso.
- **Auto-detección de los Excel** en `~/Downloads`, `~/Desktop`, `~/Documents`
  (y un nivel de subcarpetas). Si hay un único candidato lo usa; si hay varios,
  pregunta; si no hay ninguno, pide la ruta con autocompletado.
- **Vista previa** de la rotación antes de escribir nada, con confirmación.
- **Guarda el Excel junto al PMP fuente** y ofrece **abrirlo** al terminar.
- **Config persistente** en `~/.pmp_celula3.json` (rutas y parámetros de rotación).

## Requisitos

Python 3 con `openpyxl`, `rich` y `questionary` (ver `requirements.txt`).
Si falta alguno, el script intenta instalarlo automáticamente al arrancar.

```bash
python3 -m pip install -r requirements.txt
```

## Configuración (`~/.pmp_celula3.json`)

| Clave | Significado |
|-------|-------------|
| `ruta_pmp` / `ruta_matriz` | Rutas a los dos Excel (auto-detectadas). |
| `rotacion_pmp` | Orden de rotación de los 3 ingenieros. |
| `rotacion_n2` / `fecha_base_n2` | Rotación N2 y fecha base para el cálculo de respaldo. |
| `horario_tarde` | Clientes que se atienden en jornada de tarde. |
| `clientes_largos` | Clientes marcados como “PMP largo”. |
