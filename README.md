# PMP Generator — Célula 3

Herramienta de terminal que **arma sola el PMP de la semana siguiente** para el
equipo de la Célula 3 (giro **Mateo → Heiner → Estefania → Mateo**). Lee el Excel
del **Control de Gestión PMP**, rota a los encargados, marca los festivos y deja
todo listo para subir a SharePoint. Es **100 % offline y determinista**: no usa
internet, ni claves de API, ni servicios externos.

> **¿Solo vas a usar la herramienta (no instalarla)?** Salta a
> **[GUIA_USO.md](GUIA_USO.md)**: descargar, doble clic y listo.

---

## Índice — elige tu caso

| Quiero… | Ve a |
|---|---|
| **Solo usarla** (no soy técnico) | [GUIA_USO.md](GUIA_USO.md) o la **[Opción A](#opción-a--ejecutable-de-doble-clic-sin-python)** |
| **Instalarla en mi PC/Mac** (tengo o puedo instalar Python) | [Opción B](#opción-b--con-python-script--comando-pmp) |
| **Dejarla configurada y funcionando** | [3. Puesta en marcha](#3-puesta-en-marcha-dejarla-funcionando) |
| **Repartirla a mi equipo** (generar los ejecutables) | [5. Distribuir al equipo](#5-distribuir-al-equipo-publicar-una-versión) |
| **Resolver un problema** | [7. Problemas frecuentes](#7-problemas-frecuentes) |

---

## 1. ¿Qué hace?

Al abrir aparece un menú con flechas (↑ ↓ + Enter):

1. **📋 Generar semana siguiente** — el flujo principal:
   - Detecta **sola** la última semana del Control y propone la **siguiente** (no
     escribes fechas, solo confirmas).
   - Te pregunta si **alguien está ausente** y, si lo hay, **quién cubre** sus PMP.
   - Muestra una **vista previa** y, al confirmar, actualiza el **Control** y
     (opcional) genera un **cuadro resumen** para compartir.
2. **⚙ Configurar** — el archivo del Control, los consultores (rotación PMP de
   Célula 3) y los clientes de Célula 3 (con su horario y si es "PMP largo").

**Sobre los archivos:** trabaja **sobre un único archivo** — el Control se
**sobrescribe in situ** (no crea una copia nueva cada semana). El guardado es
**atómico** (escribe a un temporal y lo renombra), así un fallo a mitad nunca deja
el archivo dañado. El cuadro resumen, si lo pides, se guarda como `PMP_Semana.xlsx`
(nombre fijo, se reescribe).

---

## 2. Instalación

### Opción A — Ejecutable de doble clic (sin Python)

La forma recomendada para quien **solo va a usar** la herramienta. No requiere
instalar nada.

1. Entra a la **[página de Releases del repositorio](https://github.com/Jordy-Lv/pmp-generator/releases)**.
2. Descarga el `.zip` de tu sistema:
   - **Windows** → `PMP-Windows.zip`
   - **Mac con chip Apple** (M1/M2/M3/M4) → `PMP-macOS-AppleSilicon.zip`
   - **Mac con Intel** → `PMP-macOS-Intel.zip`
   - *(¿No sabes cuál? Menú Apple () → «Acerca de este Mac». Si dice «Chip Apple»
     es Apple Silicon; si dice «Procesador Intel», el de Intel.)*
3. Descomprime el `.zip` (doble clic).
4. Ábrelo:
   - **Windows**: doble clic en `PMP.exe`. La primera vez sale «Windows protegió tu
     PC» → **Más información** → **Ejecutar de todas formas**.
   - **Mac**: doble clic en **`Lanzar PMP.command`**. La primera vez, **clic
     derecho → Abrir → Abrir** (porque la app no está firmada). Mantén juntos
     `pmp` y `Lanzar PMP.command` en la misma carpeta.

> Si todavía **no hay ninguna Release publicada**, primero hay que crear una: ver
> **[5. Distribuir al equipo](#5-distribuir-al-equipo-publicar-una-versión)**.

### Opción B — Con Python (script + comando `pmp`)

Para instalarla en tu máquina y abrirla escribiendo `pmp` en la terminal. Sirve en
**macOS y Linux** (y en Windows con pequeñas variaciones).

**Requisito:** Python 3.8 o superior. Compruébalo:

```bash
python3 --version          # debe imprimir 3.8 o mayor
```

Si no lo tienes: macOS → `brew install python` · Windows → [python.org](https://www.python.org/downloads/) (marca «Add Python to PATH») · Linux → `sudo apt install python3 python3-venv`.

**Pasos:**

```bash
# 1. Clonar el repositorio (queda una carpeta pmp-generator)
git clone https://github.com/Jordy-Lv/pmp-generator.git
cd pmp-generator

# 2. Crear un entorno virtual dedicado (aísla las librerías)
python3 -m venv ~/.venvs/pmp

# 3. Instalar las 3 dependencias dentro de ese entorno
~/.venvs/pmp/bin/pip install -r requirements.txt

# 4. Comprobar que todo quedó bien (debe imprimir "self-test ok")
~/.venvs/pmp/bin/python pmp_generator.py --self-test
```

Con eso ya puedes ejecutarla con:

```bash
~/.venvs/pmp/bin/python pmp_generator.py
```

**(Opcional) Crear el comando corto `pmp`** para abrirla escribiendo solo `pmp`.
Crea un pequeño lanzador en `~/.local/bin` (ajusta la ruta del repo si lo clonaste
en otro sitio):

```bash
mkdir -p ~/.local/bin
cat > ~/.local/bin/pmp <<EOF
#!/bin/sh
exec "\$HOME/.venvs/pmp/bin/python" "$(pwd)/pmp_generator.py" "\$@"
EOF
chmod +x ~/.local/bin/pmp
```

Si al escribir `pmp` dice «command not found», añade `~/.local/bin` al PATH:

```bash
echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.zshrc   # bash → ~/.bashrc
```

Cierra y vuelve a abrir la terminal. Ahora abres la herramienta con solo:

```bash
pmp
```

> En **Windows** sin el comando `pmp`, usa directamente:
> `py -m venv %USERPROFILE%\.venvs\pmp` → `%USERPROFILE%\.venvs\pmp\Scripts\pip install -r requirements.txt` → `%USERPROFILE%\.venvs\pmp\Scripts\python pmp_generator.py`.

---

## 3. Puesta en marcha (dejarla funcionando)

1. **Coloca el Excel del Control** en una de estas carpetas (o una subcarpeta
   directa): `~/Downloads`, `~/Desktop` o `~/Documents`. Debe llamarse algo con
   **`Control_Gestion_PMP`** y tener la pestaña **`2026`**.
2. **Abre la herramienta** (`pmp`, o el ejecutable). Verás un panel **«Detección de
   archivos»** que confirma que encontró el Control. Si hay varios candidatos te
   deja elegir; si no encuentra, entra a **⚙ Configurar → Archivo del Control** y
   selecciónalo a mano (queda recordado).
3. **Revisa consultores y clientes** en **⚙ Configurar**:
   - **Consultores → rotación PMP**: confirma que el orden del **giro** sea el
     correcto (la pantalla muestra `A → B → C → A`). Ajusta con añadir/quitar/
     reordenar.
   - **Clientes de Célula 3**: revisa la lista, el **horario** (mañana/tarde) y los
     marcados como **«PMP largo»**.
   - Los nombres deben coincidir **exactamente** con los del Excel (mayúsculas,
     tildes, paréntesis). Todo esto se guarda en `~/.pmp_celula3.json`.
4. **Haz una semana de prueba**: **📋 Generar semana siguiente** → revisa la **vista
   previa** con calma → si algo no cuadra, cancela (no escribe nada). Cuando esté
   bien, confirma.

> La configuración (ruta del Control, rotación, clientes) vive en
> `~/.pmp_celula3.json`, que es **propio de cada máquina**. Ver
> [6. Configuración](#6-configuración-referencia).

---

## 4. Uso semana a semana

1. Descarga/actualiza el Control desde SharePoint (o usa el que ya tienes).
2. Abre la herramienta → **📋 Generar semana siguiente**.
3. Confirma la fecha propuesta (la siguiente a la última del Control).
4. Marca ausentes si los hay y elige **quién los cubre**.
5. Revisa la **vista previa** y confirma.
6. La herramienta actualiza el Control; al terminar te ofrece **abrir el Control**
   para que lo revises antes de **subirlo a SharePoint**.

---

## 5. Distribuir al equipo (publicar una versión)

Los ejecutables se generan **solos** en GitHub Actions. Pasos:

1. **(Importante) Deja la configuración «de fábrica» correcta.** El ejecutable
   arranca con los valores escritos en el código (`DEFAULT_CFG` en
   [`pmp_generator.py`](pmp_generator.py)), **no** con tu `~/.pmp_celula3.json`. Si
   personalizaste rotación o clientes desde el menú, conviene reflejar esos valores
   en `DEFAULT_CFG` antes de publicar, o pasarle a cada persona tu
   `~/.pmp_celula3.json`.
2. **Crea y sube un tag de versión** (esto dispara la construcción):

   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```

3. Espera unos minutos. En la pestaña **Actions** del repo verás el workflow
   [`build`](.github/workflows/build.yml) compilando los tres sistemas y corriendo
   `--self-test` dentro de cada binario.
4. Al terminar, los `.zip` quedan adjuntos en
   **[Releases](https://github.com/Jordy-Lv/pmp-generator/releases)**.
5. **Comparte** a cada persona el `.zip` de su sistema (o el enlace a la Release).
   A partir de ahí ellos siguen la **[Opción A](#opción-a--ejecutable-de-doble-clic-sin-python)**.

**Construir a mano** (sin GitHub, solo para tu propio sistema):

```bash
./build.sh      # macOS  → dist/pmp + Lanzar PMP.command
.\build.ps1     # Windows → dist\PMP.exe
```

> PyInstaller **no** compila para otros sistemas: cada máquina solo produce su
> propio ejecutable. Por eso el CI cubre Windows, Mac Intel y Mac Apple Silicon.

---

## 6. Configuración (referencia)

Se guarda en `~/.pmp_celula3.json` y se edita cómodamente desde **⚙ Configurar**
(no hace falta tocar el archivo a mano).

| Clave | Significado |
|-------|-------------|
| `ruta_pmp` | Ruta al Excel del Control (auto-detectada). |
| `rotacion_pmp` | Orden del giro de los consultores de Célula 3. |
| `clientes_celula3` | Clientes que rota la Célula 3 (se buscan por nombre al leer). |
| `horario_tarde` | Clientes que se atienden en jornada de tarde. |
| `clientes_largos` | Clientes marcados como "PMP largo". |

---

## 7. Problemas frecuentes

- **«No encuentra el Control»** → ⚙ Configurar → Archivo del Control → selecciónalo a mano.
  Verifica que estén en Downloads/Desktop/Documents y con el nombre esperado.
- **«Ya existe un bloque» o «no es la semana siguiente»** → la fecha debe ser el
  lunes justo posterior a la última semana del Control. Revisa que estés usando el
  Control más reciente.
- **Mac: «no se puede abrir, desarrollador no identificado»** → clic derecho →
  **Abrir** → **Abrir** (solo la primera vez).
- **Windows: «Windows protegió tu PC»** → **Más información** → **Ejecutar de todas
  formas**.
- **`pmp: command not found`** → falta añadir `~/.local/bin` al PATH (ver
  [Opción B](#opción-b--con-python-script--comando-pmp)).
- **`python3: command not found`** → instala Python 3.8+ (ver Opción B).
- **Un consultor o cliente nuevo no se rota** → revisa que su nombre en ⚙ Configurar
  sea **idéntico** al del Excel (mayúsculas, tildes, paréntesis).
- **Cualquier otra cosa** → cierra y vuelve a abrir; si persiste, guarda una captura
  de lo que aparece.

---

## 8. Para desarrolladores

- **Un solo archivo:** toda la lógica está en [`pmp_generator.py`](pmp_generator.py).
- **Dependencias:** `openpyxl`, `rich`, `questionary` (ver
  [`requirements.txt`](requirements.txt)). Si faltan, el script intenta instalarlas
  al arrancar.
- **Pruebas rápidas** (sin tocar ningún Excel real):

  ```bash
  python3 pmp_generator.py --self-test
  ```

- **Empaquetado:** [`pmp.spec`](pmp.spec) (PyInstaller),
  [`.github/workflows/build.yml`](.github/workflows/build.yml) (CI),
  `build.sh` / `build.ps1` (build local).
- **La rotación es determinista:** la regla única vive en `siguiente_disponible()`,
  usada tanto por la vista previa como por la escritura del Control, de modo que
  ambos coinciden siempre.
