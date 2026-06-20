# PMP Generator — Célula 3

Herramienta que automatiza la **rotación semanal de clientes PMP** del equipo de
la Célula 3 (Mateo → Heiner → Estefania → Mateo).

A partir de los dos Excel reales, genera de una sola vez la semana siguiente:

- **`Control_Gestion_PMP.xlsx`** (pestaña `2026`): añade el bloque de la semana
  siguiente, conservando el formato real y rotando los consultores.
- **`Matriz Unificada de Recursos.xlsx`** (pestaña `ROTACION DISPO CELULA 3 `):
  asegura la fila de disponibilidad nocturna N2/N3 viernes-viernes. Si no está la
  Matriz, la disponibilidad se calcula por rotación.
- **Cuadro resumen** (`PMP_Semana_AAAAMMDD.xlsx`): un Excel formateado y listo
  para compartir al equipo.

Los tres salen de **una sola rotación**, así que siempre coinciden. Todo se
escribe en **copias** junto a los archivos fuente; los originales no se tocan.

> **Para la persona que solo va a usar la herramienta:** ver **[GUIA_USO.md](GUIA_USO.md)**
> (instalación por doble clic, sin tocar nada técnico).

---

## Uso (resumen)

Al abrir la herramienta aparece un menú con flechas:

1. **🌙 Consultar disponibilidad nocturna** — N2/N3 de una semana.
2. **📋 Generar semana siguiente** — flujo completo: vista previa de la rotación,
   confirmación, y generación del Control actualizado + Matriz + cuadro resumen.
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

## Asistente IA (opcional · DeepSeek)

La herramienta funciona **100 % sin IA**: la rotación y las fechas son siempre
código determinista (`siguiente_disponible`, `siguiente_lunes…`), exacto y sin
internet. La IA es una capa **opcional** (módulo [`pmp_ia.py`](pmp_ia.py)) que
solo asiste en lo difícil de forma determinista:

1. **Leer Excel desordenados** — si la lectura exacta no encuentra la tabla
   (cambió la pestaña, las columnas o el formato), DeepSeek interpreta la
   estructura como último recurso.
2. **Clasificar clientes nuevos** — sugiere horario (mañana/tarde) y si es "PMP
   largo" para clientes que aún no están en la config.
3. **Redactar el aviso para el equipo** — una nota lista para Teams/correo.
4. **Revisar anomalías** — avisa de cosas raras antes de generar (alguien con
   demasiados clientes, un cliente que desaparece, etc.).

**Privacidad — anonimización.** Antes de enviar nada a DeepSeek, los nombres
reales de consultores y clientes se reemplazan por códigos (`Consultor A`,
`Cliente 1`) y se reconstruyen al recibir la respuesta. Los nombres reales no
salen del equipo. **Única excepción:** la clasificación de clientes envía solo el
nombre del cliente aislado (sin consultores ni asignaciones), porque clasificarlo
lo requiere; se puede apagar con `ia_clasificar_clientes = false`.

**Robustez.** Sin API key, sin internet o ante cualquier error, todas las
funciones de IA se omiten en silencio y el flujo determinista continúa igual.

**Key embebida (nadie la configura).** La key de DeepSeek viaja **dentro del
ejecutable** mediante el archivo `pmp_key.py` (NO versionado, ver `.gitignore`),
así la compañera abre la app y la IA ya funciona. Cómo se rellena:

- **Build local** (`build.sh`/`build.ps1`): se usa el `pmp_key.py` que tengas en
  el equipo de desarrollo.
- **CI**: el workflow lo genera desde el *secret* `DEEPSEEK_API_KEY`
  (GitHub → Settings → Secrets and variables → Actions). Sin el secret, el binario
  se construye igual pero sin IA.

Quien quiera puede **sobrescribir** la key o desactivar la IA desde el menú
**⚙ Configurar rutas → Asistente IA** (se guarda en `~/.pmp_celula3.json`).
Modelo por defecto: `deepseek-v4-flash`, sin "thinking" (respuestas directas;
céntimos al mes para este volumen).

> **Seguridad:** una key embebida es extraíble del binario por quien lo tenga.
> Para uso interno suele bastar; conviene ponerle a la cuenta DeepSeek un
> **límite de gasto** y rotar la key si se filtra (basta regenerar `pmp_key.py` /
> el secret y reconstruir).

## Configuración (`~/.pmp_celula3.json`)

| Clave | Significado |
|-------|-------------|
| `ruta_pmp` / `ruta_matriz` | Rutas a los dos Excel (auto-detectadas). |
| `rotacion_pmp` | Orden de rotación de los 3 ingenieros. |
| `rotacion_n2` / `fecha_base_n2` | Rotación N2 y fecha base para el cálculo de respaldo. |
| `rotacion_n3` | Orden de respaldo de escalamiento N3. |
| `horario_tarde` | Clientes que se atienden en jornada de tarde. |
| `clientes_largos` | Clientes marcados como "PMP largo". |
| `deepseek_api_key` | Clave de DeepSeek (vacía = IA desactivada). |
| `deepseek_model` | Modelo a usar (por defecto `deepseek-v4-flash`). |
| `ia_habilitada` | Activa/desactiva la IA aunque haya key. |
| `ia_clasificar_clientes` | Permite enviar el nombre del cliente para clasificarlo. |
