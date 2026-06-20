#!/usr/bin/env python3
"""
PMP Generator — Célula 3  ·  v3.0
Herramienta de terminal para la rotación semanal de clientes PMP
(Mateo → Heiner → Estefania) del equipo de la Célula 3.

UI rica con  rich  +  questionary.  Lectura/escritura con  openpyxl.
Config persistente en  ~/.pmp_celula3.json

Lanzar:   pmp           (comando instalado en ~/.local/bin)
   o:     ~/.venvs/pmp/bin/python pmp_generator.py
"""
from __future__ import annotations

# ══════════════════════════════════════════════════════════════════════════════
#  0 · BOOTSTRAP DE DEPENDENCIAS
#     Si falta rich / questionary / openpyxl, se muestra un aviso, se instalan
#     automáticamente en el intérprete actual y el script se relanza una vez.
# ══════════════════════════════════════════════════════════════════════════════
import copy, importlib.util, json, os, subprocess, sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

_REQUERIDAS = ("rich", "questionary", "openpyxl")


def _faltantes() -> List[str]:
    return [m for m in _REQUERIDAS if importlib.util.find_spec(m) is None]


def _bootstrap_dependencias() -> None:
    # En el ejecutable empaquetado (PyInstaller) las dependencias ya van dentro:
    # no hay pip que invocar ni intérprete que relanzar, así que se omite.
    if getattr(sys, "frozen", False):
        return
    falt = _faltantes()
    if not falt:
        return
    print("\n  ⚠  Faltan librerías necesarias: " + ", ".join(falt))
    print("  ⏳  Instalándolas automáticamente con pip…\n")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", *falt])
    except Exception:
        print("\n  ❌  No se pudieron instalar automáticamente.")
        print("      Instálalas a mano y vuelve a ejecutar:\n")
        print(f"        {sys.executable} -m pip install {' '.join(falt)}\n")
        sys.exit(1)
    # Relanzar una sola vez para que los import del módulo encuentren todo.
    if os.environ.get("PMP_REEXEC") != "1":
        os.environ["PMP_REEXEC"] = "1"
        os.execv(sys.executable, [sys.executable, *sys.argv])
    # Ya se reintentó: si aún faltan, pip "instaló" pero este intérprete no las
    # ve (otro site-packages / PEP 668). Abortar con mensaje claro, no ImportError.
    if _faltantes():
        print("\n  ❌  Las librerías se instalaron pero no son visibles para este Python.")
        print("      Instálalas a mano en este intérprete y vuelve a ejecutar:\n")
        print(f"        {sys.executable} -m pip install {' '.join(falt)}\n")
        sys.exit(1)


_bootstrap_dependencias()

# A partir de aquí las dependencias están garantizadas.
import questionary
from questionary import Style as QStyle
from rich import box
from rich.console import Console, Group
from rich.panel import Panel
from rich.table import Table
from rich.text import Text

console = Console()

# Capa OPCIONAL de IA (DeepSeek), en su propio módulo. Si no está disponible o
# falla al importar, la herramienta funciona igual sin IA.
try:
    import pmp_ia
except Exception:
    pmp_ia = None


def _ia_on(cfg: dict) -> bool:
    """True solo si el módulo de IA cargó y hay key + IA habilitada en la config."""
    return pmp_ia is not None and pmp_ia.ia_activa(cfg)


# Estilo de los prompts de questionary (acento cian, coherente con el banner).
QS = QStyle([
    ("qmark",       "fg:#36c2ce bold"),
    ("question",    "bold"),
    ("pointer",     "fg:#36c2ce bold"),
    ("highlighted", "fg:#36c2ce bold"),
    ("selected",    "fg:#22a565"),
    ("answer",      "fg:#36c2ce bold"),
    ("instruction", "fg:#888888"),
])


class Volver(Exception):
    """El usuario canceló (Ctrl+C) — volver al menú principal."""


# ══════════════════════════════════════════════════════════════════════════════
#  1 · CONFIG PERSISTENTE  (~/.pmp_celula3.json)   ·  formato sin cambios
# ══════════════════════════════════════════════════════════════════════════════
CONFIG_PATH = Path.home() / ".pmp_celula3.json"
DEFAULT_CFG: dict = {
    "ruta_pmp":        "",
    "ruta_matriz":     "",
    "rotacion_pmp":    ["Mateo Florez", "Heiner Diaz", "Estefania Sanabria"],
    "rotacion_n2":     ["Estefania Sanabria", "Heiner Diaz", "Mateo Florez"],
    "fecha_base_n2":   "2026-06-12",
    "rotacion_n3":     ["Carlos Barrera", "Santiago Amaya", "Adriano Carreño"],
    "horario_tarde":   ["La Riviera (APP)", "La Riviera (DB)", "HOMI Oracle", "HOMI SQLServer / MySQL"],
    "clientes_largos": ["HOMI Oracle", "HOMI SQLServer / MySQL"],
    # ── IA opcional (DeepSeek). Vacía = desactivada; la configura el responsable. ──
    "deepseek_api_key":        "",
    "deepseek_model":          "deepseek-v4-flash",
    "ia_habilitada":           True,
    "ia_clasificar_clientes":  True,   # envía solo el nombre del cliente (ver pmp_ia.py)
}


def cargar_cfg() -> dict:
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                return {**DEFAULT_CFG, **json.load(f)}
        except Exception:
            pass
    return DEFAULT_CFG.copy()


def guardar_cfg(cfg: dict) -> None:
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


FORMATOS_FECHA = ("%d/%m/%Y", "%Y-%m-%d", "%d%m%Y", "%d-%m-%Y")


def parse_fecha(raw: str) -> date:
    for fmt in FORMATOS_FECHA:
        try:
            return datetime.strptime(raw.strip(), fmt).date()
        except ValueError:
            continue
    raise ValueError("Usa una fecha tipo DD/MM/AAAA")


# ══════════════════════════════════════════════════════════════════════════════
#  2 · LÓGICA DE NEGOCIO   ·  sin cambios funcionales
# ══════════════════════════════════════════════════════════════════════════════
def siguiente_lunes() -> date:
    hoy = date.today()
    dias = (7 - hoy.weekday()) % 7 or 7          # días hasta el próximo lunes
    return hoy + timedelta(days=dias)


def leer_dispo(lunes: date, ruta: str) -> Tuple[Optional[str], Optional[str]]:
    """Pestaña 'ROTACION DISPO CELULA 3 ' (con espacio final) de la Matriz."""
    from openpyxl import load_workbook
    wb = load_workbook(ruta, data_only=True)
    ws = wb["ROTACION DISPO CELULA 3 "]
    for row in ws.iter_rows(values_only=True):
        if len(row) < 5:
            continue                       # fila sin las 5 columnas esperadas
        ini, fin, n2, n3 = row[1], row[2], row[3], row[4]
        if hasattr(ini, "date") and hasattr(fin, "date"):
            if ini.date() <= lunes <= fin.date():
                # Preservar None en celdas vacías (str(None) daría el literal "None",
                # que es truthy y rompería el fallback / mostraría "None" en pantalla).
                return (str(n2) if n2 is not None else None,
                        str(n3) if n3 is not None else None)
    return None, None


def calcular_dispo_fallback(lunes: date, cfg: dict) -> str:
    """Fallback por cálculo cuando no hay Matriz: rota desde una fecha base.
    Tolera un config corrupto (lista vacía o fecha mal escrita) cayendo a los
    valores por defecto en vez de reventar con un traceback."""
    orden = cfg.get("rotacion_n2") or DEFAULT_CFG["rotacion_n2"]
    try:
        base = date.fromisoformat(str(cfg.get("fecha_base_n2", "")))
    except (ValueError, TypeError):
        base = date.fromisoformat(DEFAULT_CFG["fecha_base_n2"])
    idx = ((lunes - base).days // 7) % len(orden)
    return orden[idx]


def leer_asignaciones(lunes: date, ruta: str, rotacion: List[str]) -> Dict[str, List[str]]:
    """Pestaña '2026' del PMP: clientes por ingeniero en la semana dada."""
    from openpyxl import load_workbook
    asig: Dict[str, List[str]] = {ing: [] for ing in rotacion}
    set_rotacion = set(rotacion)
    ignorar = {"", "-", "FESTIVO", "HORA DE ENVIO", *set_rotacion}
    wb = load_workbook(ruta, data_only=True)
    ws = wb["2026"]
    header_row = None
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if len(row) > 3 and hasattr(row[3], "date") and row[3].date() == lunes:
            header_row = i
            break
    if not header_row:
        return asig
    for i in range(header_row + 1, header_row + 40):
        col_d = ws.cell(i, 4).value
        col_b = ws.cell(i, 2).value
        if hasattr(col_d, "date"):
            break
        if col_b in set_rotacion:
            # ponytail: la hoja pone FESTIVO/vacío en lunes; D:K cubre la semana sin mapear toda la plantilla.
            for celda in range(4, 12):
                valor = ws.cell(i, celda).value
                if hasattr(valor, "date"):
                    continue
                cliente = str(valor).strip() if valor is not None else ""
                if cliente and cliente not in ignorar and cliente not in asig[col_b]:
                    asig[col_b].append(cliente)
    return asig


def _hoja_probable(wb):
    """La hoja con la tabla semanal: '2026' si existe; si no, la que más fechas
    tenga (usado solo en el fallback de IA, cuando el formato cambió)."""
    if "2026" in wb.sheetnames:
        return wb["2026"]
    mejor, score = None, -1
    for ws in wb.worksheets:
        n = sum(1 for row in ws.iter_rows(max_row=min(ws.max_row, 60))
                for c in row if hasattr(c.value, "date"))
        if n > score:
            mejor, score = ws, n
    return mejor


def leer_asignaciones_resiliente(lunes_act: date, ruta_pmp: str, cfg: dict):
    """Lee las asignaciones. Si la lectura exacta no encuentra nada Y la IA está
    activa, pide a DeepSeek que interprete la estructura del Excel (fallback que
    no envía nombres reales). Devuelve (asignaciones, nota|None)."""
    try:
        asig = leer_asignaciones(lunes_act, ruta_pmp, cfg["rotacion_pmp"])
    except Exception:
        asig = {ing: [] for ing in cfg["rotacion_pmp"]}
    if any(asig.values()) or not _ia_on(cfg):
        return asig, None
    try:
        from openpyxl import load_workbook
        ws = _hoja_probable(load_workbook(ruta_pmp, data_only=True))
        mapa = pmp_ia.mapear_estructura(cfg, ws) if ws is not None else None
        if not mapa:
            return asig, None
        asig2 = pmp_ia.leer_asignaciones_con_mapeo(ws, mapa, cfg["rotacion_pmp"])
    except Exception:
        return asig, None
    if any(asig2.values()):
        return asig2, "estructura interpretada por IA — revisa la vista previa"
    return asig, None


def clasificar_clientes_nuevos(cfg: dict, clientes: List[str]) -> List[str]:
    """Para clientes aún no clasificados, pide a la IA su horario / si es PMP largo
    y, si el usuario acepta, los añade a la config. Devuelve los nombres añadidos.
    Privacidad: clasificar_cliente envía solo el nombre del cliente (ver pmp_ia)."""
    if not _ia_on(cfg):
        return []
    conocidos = set(cfg.get("horario_tarde", [])) | set(cfg.get("clientes_largos", []))
    nuevos = [c for c in dict.fromkeys(clientes) if c and c not in conocidos]
    if not nuevos or not confirmar(
            f"¿Clasificar {len(nuevos)} cliente(s) nuevo(s) con IA (horario / PMP largo)?",
            default=False):
        return []
    with console.status("[cyan]Consultando IA…", spinner="dots"):
        sugerencias = {c: pmp_ia.clasificar_cliente(cfg, c) for c in nuevos}
    cambios: List[str] = []
    for c, s in sugerencias.items():
        if not s:
            continue
        etiquetas = ([f"horario {s['horario']}"] +
                     (["PMP largo"] if s["pmp_largo"] else []))
        if confirmar(f"  «{c}» → {', '.join(etiquetas)}.  ¿Añadir a la config?", default=True):
            if s["horario"] == "tarde" and c not in cfg["horario_tarde"]:
                cfg["horario_tarde"].append(c)
            if s["pmp_largo"] and c not in cfg["clientes_largos"]:
                cfg["clientes_largos"].append(c)
            cambios.append(c)
    if cambios:
        guardar_cfg(cfg)
    return cambios


def siguiente_disponible(nombre: str, rotacion: List[str], ausentes: List[str]) -> str:
    """Siguiente consultor en la rotación que no esté ausente.

    Es la ÚNICA regla de rotación del proyecto: la usan tanto `rotar()` (que mueve
    las carteras en memoria para el cuadro resumen) como `actualizar_control_pmp()`
    (que renombra los consultores dentro del Control real). Así ambos artefactos
    rotan exactamente igual y nunca divergen. Sin ausentes equivale a "el siguiente".
    """
    if nombre not in rotacion:
        return nombre
    n = len(rotacion)
    sig = (rotacion.index(nombre) + 1) % n
    intentos = 0
    while rotacion[sig] in ausentes and intentos < n:
        sig = (sig + 1) % n
        intentos += 1
    return rotacion[sig]


def rotar(actuales: Dict[str, List[str]], ausentes: List[str], rotacion: List[str]) -> Dict[str, List[str]]:
    """Mateo→Heiner→Estefania→Mateo. Si el destino está ausente, salta al siguiente."""
    nuevas: Dict[str, List[str]] = {ing: [] for ing in rotacion}
    for ing in rotacion:
        clientes = actuales.get(ing, [])
        if clientes:
            nuevas[siguiente_disponible(ing, rotacion, ausentes)].extend(clientes)
    return nuevas


# Clasificación de horario — extraída como helper para que el PREVIEW y el Excel
# muestren exactamente lo mismo (mismo cálculo, sin divergencias).
def horario_de_cliente(cliente: str, horario_tarde: set) -> str:
    return "Tarde (12:00–17:00)" if cliente in horario_tarde else "Mañana (7:00–12:00)"


def horario_predominante(clientes: List[str], cfg: dict) -> str:
    tarde = set(cfg["horario_tarde"])
    horarios = [horario_de_cliente(c, tarde) for c in clientes]
    return max(set(horarios), key=horarios.count) if horarios else "Mañana (7:00–12:00)"


def generar_excel(lunes: date, asig: Dict[str, List[str]], n2: str, n3: Optional[str],
                  ausentes: List[str], cfg: dict, dir_destino: Path) -> str:
    """Genera el Excel formateado. Lo guarda en `dir_destino` (junto al PMP fuente)."""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    rotacion      = cfg["rotacion_pmp"]
    largos        = set(cfg["clientes_largos"])
    dias          = [lunes + timedelta(i) for i in range(5)]
    dias_nom      = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]

    wb = Workbook(); ws = wb.active; ws.title = "Asignación PMP"

    def fill(h):       return PatternFill("solid", fgColor=h)
    def fnt(color="000000", bold=False, size=10): return Font(color=color, bold=bold, size=size)
    def aln(h="center", v="center"): return Alignment(horizontal=h, vertical=v, wrap_text=True)
    t  = Side(border_style="thin", color="BBBBBB")
    br = Border(left=t, right=t, top=t, bottom=t)
    N  = 2 + len(dias)

    def hdr(cell, txt, bg, fg="FFFFFF", sz=10, bold=True):
        cell.value = txt; cell.fill = fill(bg)
        cell.font  = fnt(fg, bold, sz); cell.alignment = aln(); cell.border = br

    ws.merge_cells(f"A1:{get_column_letter(N)}1")
    hdr(ws["A1"], f"📋  PMP SEMANAL — {lunes.strftime('%d/%m/%Y')} al {(lunes+timedelta(4)).strftime('%d/%m/%Y')}",
        "C00000", sz=12)
    ws.row_dimensions[1].height = 30

    hdr(ws["A2"], "Consultor", "C00000"); hdr(ws["B2"], "Horario", "C00000")
    for i, (d_, nom) in enumerate(zip(dias, dias_nom)):
        hdr(ws[f"{get_column_letter(3+i)}2"], f"{nom}\n{d_.strftime('%d/%m/%Y')}", "E74C3C")
    ws.row_dimensions[2].height = 38

    fila = 3
    colores_fila = ["FDEDEC", "FFFFFF"]
    for idx, ing in enumerate(rotacion):
        ausente  = ing in ausentes
        clientes = asig.get(ing, [])
        cf       = "D5D8DC" if ausente else colores_fila[idx % 2]
        ws.row_dimensions[fila].height = max(32, 18 * max(len(clientes), 1))

        horario_display = horario_predominante(clientes, cfg)

        ca = ws[f"A{fila}"]
        ca.value = ing; ca.fill = fill(cf)
        ca.font  = fnt(bold=True, color="CC0000" if ausente else "000000")
        ca.alignment = aln("left"); ca.border = br

        cb = ws[f"B{fila}"]
        cb.value = horario_display; cb.fill = fill(cf)
        cb.font  = fnt(); cb.alignment = aln(); cb.border = br

        for j in range(5):
            col = get_column_letter(3 + j); cc = ws[f"{col}{fila}"]
            cc.border = br; cc.alignment = aln("left")
            if ausente:
                cc.value = "\n".join(f"⚠️ {c_} — REASIGNAR" for c_ in clientes) or "—"
                cc.fill = fill("FADBD8"); cc.font = fnt("922B21", True)
            else:
                lineas = [c_ + (" ⚠️ PMP largo" if c_ in largos else "") for c_ in clientes]
                cc.value = "\n".join(lineas) or "—"
                cc.fill  = fill("FEF9E7" if any(c_ in largos for c_ in clientes) else cf)
                cc.font  = fnt()
        fila += 1

    for texto, color in [
        (f"🌙  DISPO NOCTURNA N2 (V-V 10pm–7am): {n2}", "FFC300"),
        *([(f"🆘  ESCALAMIENTO N3 (cambios / desbordes): {n3}", "F0B27A")] if n3 else []),
        *([(f"⚠️  Ausentes: {', '.join(ausentes)} — revisar reasignación", "E74C3C")] if ausentes else []),
    ]:
        ws.merge_cells(f"A{fila}:{get_column_letter(N)}{fila}")
        cc = ws[f"A{fila}"]
        cc.value = texto; cc.fill = fill(color)
        cc.font  = fnt("FFFFFF" if color == "E74C3C" else "000000", True)
        cc.alignment = aln(); cc.border = br
        ws.row_dimensions[fila].height = 22; fila += 1

    ws.column_dimensions["A"].width = 24; ws.column_dimensions["B"].width = 22
    for i in range(5):
        ws.column_dimensions[get_column_letter(3+i)].width = 32

    # Hoja Config — parámetros de esta generación
    ws2 = wb.create_sheet("Config")
    ws2["A1"] = "Parámetros de esta generación"; ws2["A1"].font = fnt(bold=True)
    fin = (lunes + timedelta(4)).strftime("%d/%m/%Y")
    for r_, (k, v) in enumerate([
        ("Semana", f"{lunes.strftime('%d/%m/%Y')} → {fin}"),
        ("Generado el", date.today().strftime("%d/%m/%Y")),
        ("Dispo N2", n2), ("Dispo N3", n3 or "N/A"),
        ("Ausentes", ", ".join(ausentes) or "Ninguno"),
    ], 2):
        ws2[f"A{r_}"] = k; ws2[f"B{r_}"] = v; ws2[f"A{r_}"].font = fnt(bold=True)
    ws2.column_dimensions["A"].width = 16; ws2.column_dimensions["B"].width = 50

    nombre  = f"PMP_Semana_{lunes.strftime('%Y%m%d')}.xlsx"
    destino = Path(dir_destino) / nombre
    wb.save(destino)
    return str(destino)


# ══════════════════════════════════════════════════════════════════════════════
#  3 · UI  ·  banner, helpers visuales y prompts de questionary
# ══════════════════════════════════════════════════════════════════════════════
def limpiar() -> None:
    console.clear()


def banner() -> None:
    limpiar()
    titulo = Text("📋  PMP GENERATOR", style="bold white")
    titulo.append("  ·  Célula 3", style="bold cyan")
    sub = Text("Rotación semanal  ·  Mateo → Heiner → Estefania", style="dim")
    console.print(Panel(Group(titulo, sub), box=box.DOUBLE, border_style="cyan",
                        padding=(0, 2), expand=False))


def _ask(pregunta):
    """Ejecuta un prompt de questionary; si el usuario cancela (Ctrl+C) → Volver.
    Usa unsafe_ask() porque ask() se traga el Ctrl+C e imprime 'Cancelled by user'."""
    try:
        resp = pregunta.unsafe_ask()
    except (KeyboardInterrupt, EOFError):
        raise Volver()
    if resp is None:
        raise Volver()
    return resp


def confirmar(mensaje: str, default: bool = True) -> bool:
    return _ask(questionary.confirm(mensaje, default=default, style=QS, auto_enter=False))


def pedir_fecha(prompt: str, default: Optional[date] = None) -> date:
    """Pide una fecha con varios formatos aceptados; reintenta hasta que sea válida."""
    default_str = default.strftime("%d/%m/%Y") if default else ""
    while True:
        raw = _ask(questionary.text(
            prompt, default=default_str, style=QS,
            instruction="(DD/MM/AAAA · Enter para el valor sugerido)")).strip()
        try:
            return parse_fecha(raw)
        except ValueError:
            pass
        console.print(f"  [red]✗[/] Formato no válido: '{raw}'. Usa DD/MM/AAAA.")


def seleccionar_ausentes(ingenieros: List[str]) -> List[str]:
    """Checkbox con flechas + espacio. Sin marcar = nadie ausente."""
    return _ask(questionary.checkbox(
        "¿Alguien ausente esta semana?  (Espacio para marcar · Enter para confirmar)",
        choices=ingenieros, style=QS))


# ── Auto-detección de los Excel ───────────────────────────────────────────────
DIRS_BUSQUEDA = [Path.home() / "Downloads", Path.home() / "Desktop", Path.home() / "Documents"]


def buscar_excels(patron: str) -> List[Path]:
    """Busca *<patron>*.xlsx (SIN distinguir mayúsculas) en Downloads/Desktop/
    Documents y un nivel de subcarpetas."""
    encontrados: dict[str, Path] = {}
    aguja = patron.lower()

    def coincide(nombre: str) -> bool:
        n = nombre.lower()
        return aguja in n and n.endswith(".xlsx") and not nombre.startswith("~$")

    for base in DIRS_BUSQUEDA:
        if not base.is_dir():
            continue
        # nivel 0 (la carpeta base) + nivel 1 (sus subcarpetas directas)
        candidatos = list(base.iterdir())
        for sub in [d for d in list(candidatos) if d.is_dir()]:
            try:
                candidatos.extend(sub.iterdir())
            except OSError:
                continue          # subcarpeta sin permisos: se ignora
        for p in candidatos:
            if p.is_file() and coincide(p.name):
                encontrados[str(p.resolve())] = p
    return sorted(encontrados.values(), key=lambda p: (len(p.parts), p.name.lower()))


def _etiqueta_archivo(p: Path) -> str:
    try:
        rel = p.relative_to(Path.home())
        return f"~/{rel}"
    except ValueError:
        return str(p)


def resolver_ruta(cfg: dict, clave: str, patron: str, descripcion: str,
                  requerido: bool = True) -> Optional[str]:
    """
    Devuelve una ruta válida para `clave`:
      1. la guardada en config, si el archivo existe;
      2. auto-detección — 1 resultado se usa directo, varios se eligen con select;
      3. si no hay ninguno, se pide la ruta a mano (autocompletado de path).
    El resultado se persiste en la config.
    """
    actual = cfg.get(clave, "")
    if actual and Path(actual).is_file():
        return actual

    halladas = buscar_excels(patron)

    if len(halladas) == 1:
        ruta = str(halladas[0].resolve())
        console.print(f"  [green]✓[/] {descripcion} detectado: [cyan]{_etiqueta_archivo(halladas[0])}[/]")
        cfg[clave] = ruta; guardar_cfg(cfg)
        return ruta

    if len(halladas) > 1:
        opciones = [questionary.Choice(title=_etiqueta_archivo(p), value=str(p.resolve()))
                    for p in halladas]
        opciones.append(questionary.Choice(title="✎  Escribir la ruta manualmente…", value="__manual__"))
        elegida = _ask(questionary.select(
            f"Se encontraron varios candidatos para {descripcion}:", choices=opciones, style=QS))
        if elegida != "__manual__":
            cfg[clave] = elegida; guardar_cfg(cfg)
            return elegida

    # 0 resultados (o eligió escribir manualmente)
    if not requerido and len(halladas) == 0:
        if not confirmar(f"No se encontró {descripcion}. ¿Indicar la ruta manualmente?", default=False):
            return None
    return _pedir_ruta_manual(cfg, clave, descripcion, requerido)


def _pedir_ruta_manual(cfg: dict, clave: str, descripcion: str, requerido: bool) -> Optional[str]:
    while True:
        raw = _ask(questionary.path(f"Ruta a {descripcion}:", style=QS)).strip()
        if not raw and not requerido:
            return None
        ruta = str(Path(raw).expanduser())
        if Path(ruta).is_file() and ruta.lower().endswith(".xlsx"):
            cfg[clave] = ruta; guardar_cfg(cfg)
            return ruta
        console.print(f"  [red]✗[/] No es un .xlsx válido: '{raw}'")
        if not requerido and not confirmar("¿Intentar de nuevo?", default=True):
            return None


def auto_detectar_inicio(cfg: dict) -> None:
    """Al arrancar: rellena en silencio las rutas únicas y muestra un panel-resumen."""
    lineas: List[Text] = []
    for clave, patron, desc in [("ruta_pmp", "PMP", "Control de Gestión PMP"),
                                 ("ruta_matriz", "Matriz", "Matriz Unificada")]:
        actual = cfg.get(clave, "")
        if actual and Path(actual).is_file():
            lineas.append(Text.assemble(("✓ ", "green"), (f"{desc}: ", "bold"),
                                        (_etiqueta_archivo(Path(actual)), "cyan")))
            continue
        halladas = buscar_excels(patron)
        if len(halladas) == 1:
            cfg[clave] = str(halladas[0].resolve()); guardar_cfg(cfg)
            lineas.append(Text.assemble(("✓ ", "green"), (f"{desc}: ", "bold"),
                                        (_etiqueta_archivo(halladas[0]), "cyan")))
        elif len(halladas) > 1:
            lineas.append(Text.assemble(("• ", "yellow"), (f"{desc}: ", "bold"),
                                        (f"{len(halladas)} candidatos (se elegirá al generar)", "dim")))
        else:
            lineas.append(Text.assemble(("• ", "yellow"), (f"{desc}: ", "bold"),
                                        ("sin detectar (se pedirá al usar)", "dim")))
    console.print(Panel(Group(*lineas), title="🔍  Detección de archivos",
                        border_style="cyan", box=box.ROUNDED, expand=False, padding=(0, 2)))


# ── Preview de la rotación ─────────────────────────────────────────────────────
def tabla_preview(lunes: date, nuevas: Dict[str, List[str]], ausentes: List[str],
                  n2: str, n3: Optional[str], fuente_dispo: str, cfg: dict) -> Panel:
    largos = set(cfg["clientes_largos"])
    fin = (lunes + timedelta(4)).strftime("%d/%m/%Y")

    tabla = Table(box=box.SIMPLE_HEAVY, expand=False, show_lines=True,
                  header_style="bold white on red3")
    tabla.add_column("Ingeniero", style="bold", no_wrap=True)
    tabla.add_column("Horario", justify="center")
    tabla.add_column("Clientes asignados")

    for ing in cfg["rotacion_pmp"]:
        clientes = nuevas.get(ing, [])
        if ing in ausentes:
            ing_cell = Text(f"{ing}  (AUSENTE)", style="bold red")
            cli_txt  = Text("\n".join(f"⚠ {c} — REASIGNAR" for c in clientes) or "—", style="red")
            horario  = Text("—", style="dim")
        else:
            ing_cell = Text(ing, style="bold")
            horario  = Text(horario_predominante(clientes, cfg).replace(" (", "\n("))
            partes = []
            for c in clientes:
                t = Text(c)
                if c in largos:
                    t.append("  ⚠ PMP largo", style="yellow")
                partes.append(t)
            cli_txt = Text("\n").join(partes) if partes else Text("(sin clientes)", style="dim")
        tabla.add_row(ing_cell, horario, cli_txt)

    pie = Text()
    pie.append("🌙  Dispo N2 (V-V 10pm–7am): ", style="bold")
    pie.append(f"{n2}", style="yellow")
    pie.append(f"   [{fuente_dispo}]\n", style="dim")
    if n3:
        pie.append("🆘  Escalamiento N3: ", style="bold"); pie.append(f"{n3}\n", style="orange3")
    if ausentes:
        pie.append("⚠  Ausentes: ", style="bold red")
        pie.append(", ".join(ausentes), style="red")

    cuerpo = Group(tabla, Text(), pie)
    return Panel(cuerpo, title=f"👁  Vista previa — semana {lunes.strftime('%d/%m/%Y')} → {fin}",
                 border_style="green", box=box.ROUNDED, padding=(1, 2), expand=False)


def abrir_archivo(ruta: str) -> None:
    try:
        if sys.platform == "darwin":
            subprocess.run(["open", ruta], check=False)
        elif os.name == "nt":
            os.startfile(ruta)            # type: ignore[attr-defined]
        else:
            subprocess.run(["xdg-open", ruta], check=False)
    except Exception as e:
        console.print(f"  [yellow]⚠[/] No se pudo abrir automáticamente: {e}")


def _siguiente_en_rotacion(nombre: str, rotacion: List[str]) -> str:
    return rotacion[(rotacion.index(nombre) + 1) % len(rotacion)] if nombre in rotacion else nombre


def _fecha_excel(d: date) -> datetime:
    return datetime.combine(d, datetime.min.time())


def _headers_semana_pmp(ws) -> List[Tuple[int, date]]:
    headers: List[Tuple[int, date]] = []
    for row in range(1, ws.max_row + 1):
        valor = ws.cell(row, 4).value
        if hasattr(valor, "date") and str(ws.cell(row, 2).value).strip() == "Consultor":
            headers.append((row, valor.date()))
    return headers


def siguiente_lunes_en_pmp(ruta_pmp: str) -> Optional[date]:
    from openpyxl import load_workbook
    wb = load_workbook(ruta_pmp, read_only=True, data_only=False)
    ws = wb["2026"]
    headers = _headers_semana_pmp(ws)
    return max((lunes for _, lunes in headers), default=None) + timedelta(days=7) if headers else None


def diagnostico_fecha(ruta_pmp: str, lunes_sig: date) -> Tuple[bool, Optional[str], Optional[date]]:
    """Comprueba —sin escribir nada— si la semana pedida encaja en el Control.
    Es determinista (no IA): la coherencia de fechas es una resta exacta.
    Devuelve (ok, mensaje, sugerida), donde `sugerida` es la siguiente semana que
    falta. ok=False si la semana ya existe o si no es la consecutiva esperada."""
    try:
        from openpyxl import load_workbook
        ws = load_workbook(ruta_pmp, read_only=True, data_only=False)["2026"]
        headers = _headers_semana_pmp(ws)
    except Exception:
        return True, None, None          # si no se puede leer, lo maneja el flujo normal
    if not headers:
        return True, None, None
    fechas = {f for _, f in headers}
    sugerida = max(fechas) + timedelta(days=7)
    if lunes_sig in fechas:
        return (False,
                f"  [red]✗[/] La semana {lunes_sig.strftime('%d/%m/%Y')} ya está en el Control "
                f"(esa tabla ya se creó). La siguiente que falta es "
                f"[cyan]{sugerida.strftime('%d/%m/%Y')}[/].",
                sugerida)
    if lunes_sig != sugerida:
        return (False,
                f"  [yellow]⚠[/] Las semanas deben ir consecutivas: la siguiente esperada es "
                f"[cyan]{sugerida.strftime('%d/%m/%Y')}[/], no {lunes_sig.strftime('%d/%m/%Y')}.",
                sugerida)
    return True, None, sugerida


def _copiar_bloque(ws, fila_origen: int, fila_destino: int, alto: int = 21,
                  col_min: int = 1, col_max: int = 15) -> None:
    offset = fila_destino - fila_origen

    for row in range(fila_origen, fila_origen + alto):
        ws.row_dimensions[row + offset].height = ws.row_dimensions[row].height
        for col in range(col_min, col_max + 1):
            origen = ws.cell(row, col)
            destino = ws.cell(row + offset, col)
            destino.value = origen.value
            if origen.has_style:
                destino._style = copy.copy(origen._style)
            destino.number_format = origen.number_format
            destino.font = copy.copy(origen.font)
            destino.fill = copy.copy(origen.fill)
            destino.border = copy.copy(origen.border)
            destino.alignment = copy.copy(origen.alignment)
            destino.protection = copy.copy(origen.protection)

    for merged in list(ws.merged_cells.ranges):
        if (fila_origen <= merged.min_row <= fila_origen + alto - 1 and
                fila_origen <= merged.max_row <= fila_origen + alto - 1):
            ws.merge_cells(
                start_row=merged.min_row + offset,
                start_column=merged.min_col,
                end_row=merged.max_row + offset,
                end_column=merged.max_col,
            )


def actualizar_control_pmp(ruta_pmp: str, lunes: date, ausentes: List[str], cfg: dict,
                           salida: Optional[str] = None) -> str:
    """Copia el último bloque semanal del Control PMP debajo y rota Célula 3.

    No reconstruye la tabla: conserva el formato real de Excel y solo toca el
    nuevo bloque copiado. La rotación de consultores usa `siguiente_disponible`,
    la MISMA regla que `rotar()`, de modo que el Control actualizado coincide con
    el cuadro resumen incluso cuando hay ausentes (el cliente del ausente queda a
    cargo del siguiente disponible; el aviso visual "REASIGNAR" vive solo en el
    cuadro resumen). Escribe siempre a una copia, nunca sobre el original.
    """
    from openpyxl import load_workbook

    wb = load_workbook(ruta_pmp)
    ws = wb["2026"]
    headers = _headers_semana_pmp(ws)
    if not headers:
        raise ValueError("No encontré bloques semanales en la hoja 2026.")

    if any(fecha == lunes for _, fecha in headers):
        raise ValueError(f"Ya existe un bloque para {lunes.strftime('%d/%m/%Y')}.")

    source_row, source_lunes = max((h for h in headers if h[1] < lunes), default=headers[-1])
    if lunes != source_lunes + timedelta(days=7):
        raise ValueError(
            f"La siguiente semana esperada después de {source_lunes.strftime('%d/%m/%Y')} "
            f"es {(source_lunes + timedelta(days=7)).strftime('%d/%m/%Y')}."
        )

    target_row = source_row + 24
    if any(ws.cell(r, c).value is not None for r in range(target_row, target_row + 21) for c in range(1, 16)):
        raise ValueError(f"La zona destino desde la fila {target_row} no está vacía.")

    _copiar_bloque(ws, source_row, target_row)

    date_cols = [
        c for c in range(1, ws.max_column + 1)
        if hasattr(ws.cell(source_row, c).value, "date")
    ]
    for idx, col in enumerate(date_cols[:5]):
        ws.cell(target_row, col).value = _fecha_excel(lunes + timedelta(days=idx))

    for row in range(target_row + 1, target_row + 21):
        for col in (2, 13):
            valor = ws.cell(row, col).value
            if valor in cfg["rotacion_pmp"]:
                ws.cell(row, col).value = siguiente_disponible(valor, cfg["rotacion_pmp"], ausentes)

    destino = salida or str(Path(ruta_pmp).with_name(
        f"{Path(ruta_pmp).stem}_actualizado_{lunes.strftime('%Y%m%d')}.xlsx"
    ))
    wb.save(destino)
    return destino


def _ultima_fila_dispo(ws) -> Optional[int]:
    ultima = None
    for row in range(1, ws.max_row + 1):
        if hasattr(ws.cell(row, 2).value, "date") and hasattr(ws.cell(row, 3).value, "date"):
            ultima = row
    return ultima


def asegurar_dispo_en_matriz(ruta_matriz: str, inicio: date, cfg: dict,
                             salida: Optional[str] = None) -> str:
    """Asegura que la matriz tenga la fila V-V de `inicio`.

    Si ya existe, solo guarda una copia equivalente. Si falta, agrega filas
    copiando el formato de la última fila real y rotando N2/N3.
    """
    from openpyxl import load_workbook

    wb = load_workbook(ruta_matriz)
    ws = wb["ROTACION DISPO CELULA 3 "]
    ultima = _ultima_fila_dispo(ws)
    if not ultima:
        raise ValueError("No encontré filas de disponibilidad en la matriz.")

    for row in range(1, ws.max_row + 1):
        valor = ws.cell(row, 2).value
        if hasattr(valor, "date") and valor.date() == inicio:
            destino = salida or str(Path(ruta_matriz).with_name(
                f"{Path(ruta_matriz).stem}_actualizada_{inicio.strftime('%Y%m%d')}.xlsx"
            ))
            wb.save(destino)
            return destino

    rotacion_n3 = cfg.get("rotacion_n3") or []
    if not rotacion_n3:
        rotacion_n3 = []
        for row in range(1, ws.max_row + 1):
            nombre = ws.cell(row, 5).value
            if nombre and nombre not in rotacion_n3:
                rotacion_n3.append(str(nombre))

    while ws.cell(ultima, 2).value.date() < inicio:
        nueva = ultima + 1
        for col in range(2, 6):
            origen = ws.cell(ultima, col)
            destino = ws.cell(nueva, col)
            if origen.has_style:
                destino._style = copy.copy(origen._style)
            destino.number_format = origen.number_format
            destino.font = copy.copy(origen.font)
            destino.fill = copy.copy(origen.fill)
            destino.border = copy.copy(origen.border)
            destino.alignment = copy.copy(origen.alignment)
            destino.protection = copy.copy(origen.protection)
        ws.row_dimensions[nueva].height = ws.row_dimensions[ultima].height

        inicio_nuevo = ws.cell(ultima, 3).value.date()
        ws.cell(nueva, 2).value = _fecha_excel(inicio_nuevo)
        ws.cell(nueva, 3).value = _fecha_excel(inicio_nuevo + timedelta(days=7))
        ws.cell(nueva, 4).value = _siguiente_en_rotacion(str(ws.cell(ultima, 4).value), cfg["rotacion_n2"])
        ws.cell(nueva, 5).value = _siguiente_en_rotacion(str(ws.cell(ultima, 5).value), rotacion_n3)
        ultima = nueva

    destino = salida or str(Path(ruta_matriz).with_name(
        f"{Path(ruta_matriz).stem}_actualizada_{inicio.strftime('%Y%m%d')}.xlsx"
    ))
    wb.save(destino)
    return destino


# ══════════════════════════════════════════════════════════════════════════════
#  4 · MODOS
# ══════════════════════════════════════════════════════════════════════════════
def modo_dispo(cfg: dict) -> None:
    console.print(Panel("🌙  Consulta de disponibilidad nocturna",
                        border_style="magenta", box=box.ROUNDED, expand=False))
    lunes = pedir_fecha("Semana a consultar (lunes)", siguiente_lunes())

    n2 = n3 = None
    fuente = "cálculo automático"
    ruta_matriz = resolver_ruta(cfg, "ruta_matriz", "Matriz", "Matriz Unificada", requerido=False)
    if ruta_matriz:
        with console.status("[cyan]Leyendo Matriz Unificada…", spinner="dots"):
            try:
                n2, n3 = leer_dispo(lunes, ruta_matriz)
                fuente = "Matriz Unificada"
            except Exception as e:
                console.print(f"  [yellow]⚠[/] {e}")
    if not n2:
        n2 = calcular_dispo_fallback(lunes, cfg)

    fin = (lunes + timedelta(4)).strftime("%d/%m/%Y")
    cuerpo = Text()
    cuerpo.append("N2 (incidentes nocturnos) : ", style="bold"); cuerpo.append(f"{n2}\n", style="cyan")
    cuerpo.append("N3 (escalamiento)         : ", style="bold"); cuerpo.append(f"{n3 or 'N/A'}\n")
    cuerpo.append("Fuente                    : ", style="bold"); cuerpo.append(fuente, style="dim")
    console.print(Panel(cuerpo, title=f"🌙  Disponibilidad — {lunes.strftime('%d/%m/%Y')} → {fin}",
                        border_style="yellow", box=box.ROUNDED, expand=False, padding=(1, 2)))


def modo_pmp(cfg: dict) -> None:
    console.print(Panel("📋  Generar semana siguiente\n"
                        "[dim]Actualiza el Control PMP y la Matriz, y crea el cuadro resumen.[/]",
                        border_style="red", box=box.ROUNDED, expand=False))

    # 1 · Resolver archivos primero (para poder sugerir la fecha correcta del Control)
    ruta_pmp = resolver_ruta(cfg, "ruta_pmp", "PMP", "Control de Gestión PMP", requerido=True)
    if not ruta_pmp:
        console.print("  [red]✗[/] Sin PMP no se puede generar."); return
    ruta_matriz = resolver_ruta(cfg, "ruta_matriz", "Matriz", "Matriz Unificada", requerido=False)

    # 2 · Fecha: se sugiere la SIGUIENTE semana que falta en el Control (no el lunes
    #     del calendario). Con Enter se aplica la correcta automáticamente.
    try:
        sugerida = siguiente_lunes_en_pmp(ruta_pmp) or siguiente_lunes()
    except Exception:
        sugerida = siguiente_lunes()
    lunes_sig = pedir_fecha("Semana a GENERAR (lunes)", sugerida)

    # 2b · Validar coherencia AHORA, antes de leer / clasificar / previsualizar.
    ok, msg, sug = diagnostico_fecha(ruta_pmp, lunes_sig)
    if not ok:
        console.print(msg)
        if sug and confirmar(f"¿Usar {sug.strftime('%d/%m/%Y')} en su lugar?", default=True):
            lunes_sig = sug
        else:
            console.print("  [dim]Cancelado — no se escribió ningún archivo.[/]"); return
    lunes_act = lunes_sig - timedelta(weeks=1)

    # 3 · Ausentes
    ausentes = seleccionar_ausentes(cfg["rotacion_pmp"])

    # 4 · Leer asignaciones (con fallback de IA si cambió el formato del Excel)
    with console.status("[cyan]Leyendo asignaciones actuales…", spinner="dots"):
        actuales, nota_lectura = leer_asignaciones_resiliente(lunes_act, ruta_pmp, cfg)
    if all(not v for v in actuales.values()):
        console.print(f"  [yellow]⚠[/] No se hallaron clientes de Célula 3 para la semana "
                      f"{lunes_act.strftime('%d/%m/%Y')} en el PMP."); return
    if nota_lectura:
        console.print(f"  [magenta]🤖 {nota_lectura}[/]")

    # 4b · Clasificar clientes nuevos con IA (opcional, pregunta antes)
    clasificar_clientes_nuevos(cfg, [c for cs in actuales.values() for c in cs])

    # 4c · Rotar (regla determinista) + disponibilidad
    nuevas = rotar(actuales, ausentes, cfg["rotacion_pmp"])
    n2 = n3 = None
    fuente_dispo = "cálculo automático"
    if ruta_matriz:
        try:
            n2, n3 = leer_dispo(lunes_sig, ruta_matriz)
            if n2:
                fuente_dispo = "Matriz Unificada"
        except Exception as e:
            console.print(f"  [yellow]⚠[/] Dispo: {e}")
    if not n2:
        n2 = calcular_dispo_fallback(lunes_sig, cfg)

    # Mostrar de dónde salió cada cliente (semana de referencia)
    ref = Text(f"Semana de referencia leída: {lunes_act.strftime('%d/%m/%Y')}\n", style="dim")
    for ing, cls in actuales.items():
        if cls:
            ref.append(f"  {ing}: ", style="dim")
            ref.append(", ".join(cls) + "\n")
    console.print(ref)

    # 4d · Revisión de anomalías con IA (opcional, anonimizada)
    if _ia_on(cfg):
        with console.status("[magenta]Revisando anomalías con IA…", spinner="dots"):
            anomalias = pmp_ia.revisar_anomalias(cfg, actuales, nuevas, ausentes)
        if anomalias:
            cuerpo = Text()
            for a in anomalias:
                cuerpo.append(f"  • {a}\n", style="yellow")
            console.print(Panel(cuerpo, title="🤖  Posibles anomalías — revisar antes de generar",
                                border_style="yellow", box=box.ROUNDED, expand=False, padding=(0, 2)))

    # 5 · PREVIEW + confirmación
    console.print(tabla_preview(lunes_sig, nuevas, ausentes, n2, n3, fuente_dispo, cfg))
    if not confirmar("¿Confirmar y generar la semana siguiente?", default=True):
        console.print("  [dim]Cancelado — no se escribió ningún archivo.[/]"); return

    # 6 · A partir de la MISMA rotación: Control real + Matriz + cuadro resumen.
    #     Todo se escribe en COPIAS junto a los archivos fuente; los originales no
    #     se tocan. El Control es la operación con validaciones: si falla, se aborta
    #     sin generar artefactos parciales para no confundir.
    dir_destino = Path(ruta_pmp).resolve().parent
    try:
        with console.status("[green]Actualizando el Control PMP…", spinner="dots"):
            control = actualizar_control_pmp(ruta_pmp, lunes_sig, ausentes, cfg)
    except Exception as e:
        console.print(f"  [red]✗[/] No se pudo actualizar el Control PMP: {e}")
        return

    matriz = None
    if ruta_matriz:
        try:
            with console.status("[green]Actualizando la Matriz de recursos…", spinner="dots"):
                matriz = asegurar_dispo_en_matriz(ruta_matriz, lunes_sig - timedelta(days=3), cfg)
        except Exception as e:
            console.print(f"  [yellow]⚠[/] No se pudo actualizar la Matriz (se omite): {e}")

    try:
        with console.status("[green]Generando el cuadro resumen…", spinner="dots"):
            resumen = generar_excel(lunes_sig, nuevas, n2, n3, ausentes, cfg, dir_destino)
    except Exception as e:
        console.print(f"  [red]✗[/] Error al generar el cuadro resumen: {e}")
        resumen = None

    # 7 · Resultado
    fin = (lunes_sig + timedelta(4)).strftime("%d/%m/%Y")
    res = Text()
    res.append("Semana        : ", style="bold"); res.append(f"{lunes_sig.strftime('%d/%m/%Y')} → {fin}\n")
    res.append("Dispo N2      : ", style="bold"); res.append(f"{n2}\n", style="yellow")
    res.append("Dispo N3      : ", style="bold"); res.append(f"{n3 or 'N/A'}\n\n")
    res.append("Control PMP   : ", style="bold"); res.append(f"{control}\n", style="cyan")
    res.append("Matriz        : ", style="bold"); res.append(f"{matriz or 'no actualizada'}\n",
                                                              style="cyan" if matriz else "dim")
    res.append("Cuadro resumen: ", style="bold"); res.append(f"{resumen or 'no generado'}",
                                                              style="cyan" if resumen else "dim")
    console.print(Panel(res, title="✅  Semana siguiente generada",
                        border_style="green", box=box.DOUBLE, expand=False, padding=(1, 2)))
    console.print("  [dim]Siguiente paso: revisa el Control PMP actualizado y súbelo a "
                  "SharePoint. El cuadro resumen es para compartirlo al equipo.[/]\n")

    if confirmar("¿Abrir el Control PMP actualizado ahora?", default=True):
        abrir_archivo(control)
    if resumen and confirmar("¿Abrir también el cuadro resumen?", default=False):
        abrir_archivo(resumen)

    # 8 · Aviso para el equipo redactado por IA (opcional, anonimizado)
    if _ia_on(cfg) and confirmar("¿Redactar un aviso para el equipo con IA?", default=False):
        with console.status("[magenta]Redactando aviso con IA…", spinner="dots"):
            aviso = pmp_ia.redactar_aviso(cfg, lunes_sig, nuevas, ausentes, n2, n3)
        if aviso:
            console.print(Panel(aviso, title="🤖  Aviso para el equipo (revísalo antes de enviar)",
                                border_style="magenta", box=box.ROUNDED, expand=False, padding=(1, 2)))
            ruta_aviso = Path(dir_destino) / f"Aviso_PMP_{lunes_sig.strftime('%Y%m%d')}.txt"
            try:
                ruta_aviso.write_text(aviso, encoding="utf-8")
                console.print(f"  [green]✓[/] Guardado en [cyan]{ruta_aviso}[/]\n")
            except Exception:
                pass
        else:
            console.print("  [yellow]⚠[/] La IA no devolvió un aviso (revisa conexión / key).\n")


def modo_config(cfg: dict) -> None:
    console.print(Panel(f"⚙  Configuración de rutas\n[dim]Se guardan en {CONFIG_PATH}[/]",
                        border_style="blue", box=box.ROUNDED, expand=False))

    def _editar(clave: str, desc: str, patron: str, requerido: bool) -> None:
        actual = cfg.get(clave, "")
        estado = f"[cyan]{_etiqueta_archivo(Path(actual))}[/]" if actual and Path(actual).is_file() \
            else ("[yellow](sin configurar)[/]" if not actual else "[red](ruta inválida)[/]")
        console.print(f"  {desc}: {estado}")
        accion = _ask(questionary.select(
            f"{desc} — ¿qué hacer?",
            choices=[questionary.Choice("Auto-detectar / volver a buscar", "auto"),
                     questionary.Choice("Escribir ruta manual", "manual"),
                     questionary.Choice("Dejar como está", "skip")],
            style=QS))
        if accion == "skip":
            return
        if accion == "auto":
            cfg[clave] = ""                       # forzar re-detección
            resolver_ruta(cfg, clave, patron, desc, requerido=requerido)
        else:
            _pedir_ruta_manual(cfg, clave, desc, requerido)

    _editar("ruta_pmp", "Control de Gestión PMP", "PMP", True)
    _editar("ruta_matriz", "Matriz Unificada", "Matriz", False)
    _configurar_ia(cfg)
    guardar_cfg(cfg)
    console.print("  [green]✓[/] Configuración guardada.")


def _configurar_ia(cfg: dict) -> None:
    """Configura la IA (DeepSeek): pegar la API key y activar/desactivar. Opcional;
    sin key la herramienta funciona igual con la rotación determinista."""
    if pmp_ia is None:
        return
    tiene_key = bool((cfg.get("deepseek_api_key") or "").strip())
    estado = "[green]activada[/]" if _ia_on(cfg) else (
        "[yellow]con key pero desactivada[/]" if tiene_key else "[dim]sin configurar[/]")
    console.print(f"\n  Asistente IA (DeepSeek): {estado}")
    accion = _ask(questionary.select(
        "Asistente IA — ¿qué hacer?",
        choices=[
            questionary.Choice("Pegar / cambiar la API key", "key"),
            questionary.Choice("Activar IA" if not cfg.get("ia_habilitada", True) else "Desactivar IA", "toggle"),
            questionary.Choice("Quitar la API key", "borrar"),
            questionary.Choice("Dejar como está", "skip"),
        ], style=QS))
    if accion == "key":
        nueva = _ask(questionary.password("Pega la API key de DeepSeek:", style=QS)).strip()
        if nueva:
            cfg["deepseek_api_key"] = nueva
            cfg["ia_habilitada"] = True
            console.print("  [green]✓[/] Key guardada. IA activada.")
    elif accion == "toggle":
        cfg["ia_habilitada"] = not cfg.get("ia_habilitada", True)
        console.print(f"  [green]✓[/] IA {'activada' if cfg['ia_habilitada'] else 'desactivada'}.")
    elif accion == "borrar":
        cfg["deepseek_api_key"] = ""
        console.print("  [green]✓[/] Key eliminada.")


# ══════════════════════════════════════════════════════════════════════════════
#  5 · MENÚ PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
def menu(cfg: dict) -> None:
    auto_detectar_inicio(cfg)
    primera = True
    while True:
        if not primera:
            banner()
        primera = False

        rutas_ok = bool(cfg.get("ruta_pmp") and Path(cfg["ruta_pmp"]).is_file())
        estado = "[green]rutas OK ✓[/]" if rutas_ok else "[yellow]configura rutas ⚠[/]"
        ia_estado = "  ·  [magenta]IA ✓[/]" if _ia_on(cfg) else ""
        console.print(f"  Estado: {estado}{ia_estado}\n")

        try:
            op = questionary.select(
                "¿Qué deseas hacer?",
                choices=[
                    questionary.Choice("🌙  Consultar disponibilidad nocturna", "1"),
                    questionary.Choice("📋  Generar semana siguiente", "2"),
                    questionary.Choice("⚙   Configurar rutas de archivos", "3"),
                    questionary.Choice("🚪  Salir", "0"),
                ],
                style=QS, qmark="›", pointer="❯").unsafe_ask()
        except (KeyboardInterrupt, EOFError):
            break

        if op in (None, "0"):
            break

        accion = {"1": modo_dispo, "2": modo_pmp, "3": modo_config}.get(op)
        if accion:
            try:
                console.print()
                accion(cfg)
            except Volver:
                console.print("  [dim]↩  Volviendo al menú…[/]")
            except Exception as e:
                # Red de seguridad: nunca un traceback crudo delante del público.
                console.print(Panel(f"[red]{type(e).__name__}:[/] {e}",
                                    title="⚠  Ocurrió un error inesperado",
                                    border_style="red", box=box.ROUNDED, expand=False,
                                    padding=(0, 2)))
            try:
                questionary.text("", qmark="",
                                 instruction="(Enter para volver al menú)").unsafe_ask()
            except (KeyboardInterrupt, EOFError):
                break
        banner()

    limpiar()
    console.print("\n  [dim]Hasta luego 👋[/]\n")


def _selftest() -> None:
    import tempfile
    from openpyxl import Workbook, load_workbook

    cfg = DEFAULT_CFG.copy()
    assert parse_fecha("19/06/2026") == date(2026, 6, 19)
    assert parse_fecha("2026-06-19") == date(2026, 6, 19)
    actuales = {
        "Mateo Florez": ["A"],
        "Heiner Diaz": ["B"],
        "Estefania Sanabria": ["C"],
    }
    rotadas = rotar(actuales, ["Heiner Diaz"], cfg["rotacion_pmp"])
    assert rotadas["Estefania Sanabria"] == ["A", "B"]
    assert rotadas["Mateo Florez"] == ["C"]

    # siguiente_disponible: la regla única de rotación, sin y con ausentes.
    rot = cfg["rotacion_pmp"]
    assert siguiente_disponible("Mateo Florez", rot, []) == "Heiner Diaz"
    assert siguiente_disponible("Mateo Florez", rot, ["Heiner Diaz"]) == "Estefania Sanabria"
    assert siguiente_disponible("Estefania Sanabria", rot, ["Heiner Diaz"]) == "Mateo Florez"

    with tempfile.TemporaryDirectory() as tmp:
        pmp = Path(tmp) / "pmp.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "2026"
        ws["B1"] = "Consultor"
        for col, d in zip((4, 6, 8, 9, 15), [date(2026, 6, 22) + timedelta(days=i) for i in range(5)]):
            ws.cell(1, col).value = _fecha_excel(d)
        ws["B3"] = "Mateo Florez"
        ws["M3"] = "Estefania Sanabria"
        wb.save(pmp)

        # Sugerencia / validación de fecha (determinista): el bloque del 22/06 ya
        # existe, así que la siguiente que falta es el 29/06.
        assert siguiente_lunes_en_pmp(str(pmp)) == date(2026, 6, 29)
        assert diagnostico_fecha(str(pmp), date(2026, 6, 22))[0] is False   # ya existe
        assert diagnostico_fecha(str(pmp), date(2026, 6, 29))[0] is True    # la que falta

        # Sin ausentes: rotación 1-a-1 (Mateo→Heiner, Estefania→Mateo).
        out = Path(tmp) / "pmp_out.xlsx"
        actualizar_control_pmp(str(pmp), date(2026, 6, 29), [], cfg, str(out))
        ws = load_workbook(out, data_only=False)["2026"]
        assert ws["D25"].value.date() == date(2026, 6, 29)
        assert ws["O25"].value.date() == date(2026, 7, 3)
        assert ws["B27"].value == "Heiner Diaz"
        assert ws["M27"].value == "Mateo Florez"

        # Con ausentes: el renombrado salta al siguiente disponible (Mateo→Estefania).
        out_aus = Path(tmp) / "pmp_out_aus.xlsx"
        actualizar_control_pmp(str(pmp), date(2026, 6, 29), ["Heiner Diaz"], cfg, str(out_aus))
        ws = load_workbook(out_aus, data_only=False)["2026"]
        assert ws["B27"].value == "Estefania Sanabria"
        assert ws["M27"].value == "Mateo Florez"

        matriz = Path(tmp) / "matriz.xlsx"
        matriz_out = Path(tmp) / "matriz_out.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "ROTACION DISPO CELULA 3 "
        ws["B4"] = _fecha_excel(date(2026, 6, 19))
        ws["C4"] = _fecha_excel(date(2026, 6, 26))
        ws["D4"] = "Heiner Diaz"
        ws["E4"] = "Carlos Barrera"
        wb.save(matriz)
        asegurar_dispo_en_matriz(str(matriz), date(2026, 6, 26), cfg, str(matriz_out))
        wb = load_workbook(matriz_out, data_only=False)
        ws = wb["ROTACION DISPO CELULA 3 "]
        assert ws["B5"].value.date() == date(2026, 6, 26)
        assert ws["D5"].value == "Mateo Florez"

    # La capa de IA debe viajar con la app (también dentro del binario empaquetado).
    assert pmp_ia is not None, "pmp_ia no disponible (¿faltó en el empaquetado?)"
    pmp_ia._selftest()
    print("IA embebida:", "sí" if _ia_on(DEFAULT_CFG) else "no")
    print("self-test ok")


# ══════════════════════════════════════════════════════════════════════════════
#  6 · MAIN
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    try:
        cfg = cargar_cfg()
        if "--self-test" in sys.argv:
            _selftest()
            sys.exit(0)
        # Herramienta interactiva: necesita una terminal real. Al abrir el
        # ejecutable por doble clic, Windows (PMP.exe) y macOS ('Lanzar PMP.command')
        # la proveen. Si por algún motivo no la hay, se avisa en vez de fallar.
        if not (sys.stdin.isatty() and sys.stdout.isatty()):
            print("Esta herramienta es interactiva: ábrela con doble clic\n"
                  "  · Windows: PMP.exe\n"
                  "  · macOS:   'Lanzar PMP.command'")
            sys.exit(1)
        banner()
        menu(cfg)
    except (KeyboardInterrupt, EOFError):
        console.print("\n  [dim]Interrumpido. Hasta luego 👋[/]\n")
        sys.exit(0)
