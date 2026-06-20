#!/usr/bin/env python3
"""
pmp_ia.py — Capa OPCIONAL de IA (DeepSeek) para el PMP Generator.

REGLA DE ORO: la IA NUNCA decide la rotación ni las fechas. Eso es código
determinista en pmp_generator.py (siguiente_disponible, siguiente_lunes…) y debe
seguir siéndolo: es exacto, instantáneo y no depende de internet. Aquí la IA solo
ASISTE en lo que es difícil de forma determinista:

  1. leer_estructura  — interpretar un Excel con formato inesperado (fallback).
  2. clasificar_cliente — horario mañana/tarde y si es "PMP largo".
  3. redactar_aviso   — nota en lenguaje natural para el equipo.
  4. revisar_anomalias — detectar cosas raras antes de subir el Control.

PRIVACIDAD: todo lo que se envía a DeepSeek pasa por `Anonimizador`, que cambia
los nombres reales de consultores y clientes por códigos (Consultor A, Cliente 1)
y reconstruye la respuesta al volver. Los nombres reales no salen del equipo.
ÚNICA excepción: `clasificar_cliente` envía el nombre real del cliente AISLADO
(sin consultores ni asignaciones), porque clasificarlo exige saber cuál es; se
puede apagar con la clave de config `ia_clasificar_clientes`.

ROBUSTEZ: sin API key, sin internet o ante cualquier error, TODA función devuelve
None / un valor neutro y el flujo determinista continúa igual. Nunca lanza.

Sin dependencias nuevas: el cliente HTTP usa urllib de la biblioteca estándar.
"""
from __future__ import annotations

import json
import os
import re
import ssl
import urllib.error
import urllib.request
from datetime import date
from typing import Any, Dict, List, Optional

DEEPSEEK_URL = "https://api.deepseek.com/chat/completions"
MODELO_DEFECTO = "deepseek-v4-flash"   # barato y rápido; suficiente para estas tareas
TIMEOUT_S = 25


def _clave_embebida() -> str:
    """Key por defecto, para que la herramienta funcione SIN que nadie la
    configure. Prioridad: variable de entorno DEEPSEEK_API_KEY (útil en CI) >
    archivo `pmp_key.py` (NO versionado, se genera al construir y viaja dentro del
    ejecutable). Si no hay ninguno, "" (la IA queda desactivada)."""
    env = os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if env:
        return env
    try:
        from pmp_key import DEEPSEEK_API_KEY   # archivo local, fuera de git
        return (DEEPSEEK_API_KEY or "").strip()
    except Exception:
        return ""


# ── Configuración ──────────────────────────────────────────────────────────────
def ia_config(cfg: dict) -> dict:
    # La key explícita de la config gana; si no hay, se usa la embebida por defecto.
    return {
        "key": (cfg.get("deepseek_api_key") or "").strip() or _clave_embebida(),
        "modelo": (cfg.get("deepseek_model") or MODELO_DEFECTO).strip(),
        "habilitada": bool(cfg.get("ia_habilitada", True)),
        "clasificar": bool(cfg.get("ia_clasificar_clientes", True)),
    }


def ia_activa(cfg: dict) -> bool:
    """True si hay key y la IA está habilitada. Si es False, todo cae al modo
    determinista de siempre."""
    c = ia_config(cfg)
    return c["habilitada"] and bool(c["key"])


# ── Cliente DeepSeek (urllib) ───────────────────────────────────────────────────
def _chat(cfg: dict, system: str, user: str, json_mode: bool = True,
          max_tokens: int = 1024) -> Optional[str]:
    """Una llamada a DeepSeek. Devuelve el texto de respuesta, o None ante
    cualquier fallo (sin key, red caída, timeout, HTTP error, respuesta vacía).
    Nunca lanza: está pensada para que el llamador haga fallback en silencio."""
    c = ia_config(cfg)
    if not c["key"]:
        return None
    cuerpo: Dict[str, Any] = {
        "model": c["modelo"],
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        "stream": False,
        "max_tokens": max_tokens,
        "temperature": 0.1,   # tareas deterministas (clasificar/mapear) → poca variabilidad
        # Sin razonamiento en cadena: estas tareas (clasificar, redactar, mapear)
        # no lo necesitan, y con él activo el "thinking" se come el presupuesto de
        # max_tokens y deja el content vacío. Desactivarlo lo hace directo y barato.
        "thinking": {"type": "disabled"},
    }
    if json_mode:
        cuerpo["response_format"] = {"type": "json_object"}
    datos = json.dumps(cuerpo).encode("utf-8")
    req = urllib.request.Request(
        DEEPSEEK_URL, data=datos, method="POST",
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {c['key']}",
        },
    )
    try:
        ctx = ssl.create_default_context()
        with urllib.request.urlopen(req, timeout=TIMEOUT_S, context=ctx) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
        contenido = payload["choices"][0]["message"]["content"]
        return (contenido or "").strip() or None
    except Exception:
        return None


def _chat_json(cfg: dict, system: str, user: str, max_tokens: int = 1024) -> Optional[Any]:
    """Como _chat pero parsea JSON. Tolera vallas ```json. None si falla o no
    es JSON válido."""
    txt = _chat(cfg, system, user, json_mode=True, max_tokens=max_tokens)
    if not txt:
        return None
    if txt.startswith("```"):
        txt = re.sub(r"^```[a-zA-Z]*\n?", "", txt)
        txt = re.sub(r"\n?```$", "", txt).strip()
    try:
        return json.loads(txt)
    except Exception:
        return None


# ── Anonimizador ────────────────────────────────────────────────────────────────
class Anonimizador:
    """Sustituye nombres reales por códigos antes de enviar a DeepSeek y los
    restaura al recibir. Construye el mapa con los consultores de la config y con
    los clientes que se le indiquen (puede ampliarse con registrar_clientes)."""

    def __init__(self, cfg: dict, clientes: Optional[List[str]] = None):
        self._a_codigo: Dict[str, str] = {}
        self._a_real: Dict[str, str] = {}
        personas: List[str] = []
        for clave in ("rotacion_pmp", "rotacion_n2", "rotacion_n3"):
            for n in cfg.get(clave, []) or []:
                if n and n not in personas:
                    personas.append(n)
        for i, n in enumerate(personas):
            self._reg(n, f"Consultor {chr(65 + i)}")   # Consultor A, B, C…
        if clientes:
            self.registrar_clientes(clientes)

    def _reg(self, real: str, codigo: str) -> None:
        self._a_codigo[real] = codigo
        self._a_real[codigo] = real

    def registrar_clientes(self, clientes: List[str]) -> None:
        n = sum(1 for c in self._a_real if c.startswith("Cliente"))
        for cli in clientes:
            if cli and cli not in self._a_codigo:
                n += 1
                self._reg(cli, f"Cliente {n}")

    @staticmethod
    def _sub(texto: str, mapa: Dict[str, str]) -> str:
        """Sustitución en UNA sola pasada (regex): cada posición se reemplaza una
        vez y el texto insertado NO se vuelve a escanear. Así un código recién
        puesto (p. ej. 'Consultor A') no es alterado por otra clave más corta
        (p. ej. un cliente llamado 'A'). Las claves van de más larga a más corta
        para que ante solapamientos gane la coincidencia más específica."""
        if not mapa:
            return str(texto)
        claves = sorted(mapa, key=len, reverse=True)
        patron = re.compile("|".join(re.escape(k) for k in claves))
        return patron.sub(lambda m: mapa[m.group(0)], str(texto))

    def ocultar(self, texto: str) -> str:
        return self._sub(texto, self._a_codigo)

    def revelar(self, texto: str) -> str:
        return self._sub(texto, self._a_real)

    def ocultar_asig(self, asig: Dict[str, List[str]]) -> Dict[str, List[str]]:
        return {self.ocultar(k): [self.ocultar(c) for c in v] for k, v in asig.items()}


# ── 1 · Redactar el aviso para el equipo (anonimizado) ───────────────────────────
def redactar_aviso(cfg: dict, lunes: date, nuevas: Dict[str, List[str]],
                   ausentes: List[str], n2: Optional[str], n3: Optional[str]) -> Optional[str]:
    if not ia_activa(cfg):
        return None
    anon = Anonimizador(cfg, [c for cs in nuevas.values() for c in cs])
    datos = {
        "semana_inicio": lunes.strftime("%d/%m/%Y"),
        "asignaciones": anon.ocultar_asig(nuevas),
        "ausentes": [anon.ocultar(a) for a in ausentes],
        "n2": anon.ocultar(n2 or ""),
        "n3": anon.ocultar(n3 or ""),
    }
    system = (
        "Eres asistente de un equipo de soporte técnico. Redacta en español un "
        "aviso BREVE (4 a 6 líneas) para anunciar la rotación PMP de la semana. "
        "Si hay ausentes, indica quién cubre a quién. Menciona la disponibilidad "
        "nocturna N2 (incidentes) y N3 (escalamiento). Tono profesional y directo, "
        "sin saludos largos. Devuelve JSON: {\"aviso\": \"texto\"}."
    )
    user = "Datos de la semana (json):\n" + json.dumps(datos, ensure_ascii=False, indent=2)
    resp = _chat_json(cfg, system, user, max_tokens=700)
    if not isinstance(resp, dict) or "aviso" not in resp:
        return None
    return anon.revelar(str(resp["aviso"])).strip() or None


# ── 2 · Revisar anomalías (anonimizado) ──────────────────────────────────────────
def revisar_anomalias(cfg: dict, actuales: Dict[str, List[str]],
                      nuevas: Dict[str, List[str]], ausentes: List[str]) -> Optional[List[str]]:
    if not ia_activa(cfg):
        return None
    todos = [c for cs in list(actuales.values()) + list(nuevas.values()) for c in cs]
    anon = Anonimizador(cfg, todos)
    datos = {
        "semana_anterior": anon.ocultar_asig(actuales),
        "semana_nueva": anon.ocultar_asig(nuevas),
        "ausentes": [anon.ocultar(a) for a in ausentes],
    }
    system = (
        "Revisa una rotación semanal de clientes entre consultores en busca de "
        "ANOMALÍAS antes de publicarla. Reporta solo cosas claramente raras: un "
        "consultor con muchísimos más clientes que el resto, un cliente duplicado, "
        "un cliente que desaparece sin que su consultor esté ausente, o un "
        "consultor presente que se queda sin ningún cliente sin motivo. NO comentes "
        "la rotación normal ni las ausencias esperadas. Sé conciso. Devuelve JSON "
        "{\"anomalias\": [\"frase corta\", ...]} y lista vacía si todo está bien."
    )
    user = "json:\n" + json.dumps(datos, ensure_ascii=False, indent=2)
    resp = _chat_json(cfg, system, user, max_tokens=600)
    if not isinstance(resp, dict) or "anomalias" not in resp:
        return None
    salida = [anon.revelar(str(a)).strip() for a in resp.get("anomalias", [])]
    return [a for a in salida if a]


# ── 3 · Clasificar un cliente nuevo ──────────────────────────────────────────────
def clasificar_cliente(cfg: dict, cliente: str) -> Optional[dict]:
    """Devuelve {"horario": "manana"|"tarde", "pmp_largo": bool} o None.
    PRIVACIDAD: envía SOLO el nombre del cliente (sin consultores ni asignaciones).
    Se puede desactivar con cfg['ia_clasificar_clientes'] = False."""
    if not ia_activa(cfg) or not ia_config(cfg)["clasificar"] or not cliente.strip():
        return None
    system = (
        "Clasificas un cliente de soporte técnico por su nombre. Devuelve JSON "
        "{\"horario\": \"manana\" o \"tarde\", \"pmp_largo\": true o false}. "
        "'tarde' si por su naturaleza suele atenderse 12:00-17:00; 'pmp_largo' si "
        "el mantenimiento es extenso (p. ej. bases de datos grandes). Ante la duda, "
        "usa \"manana\" y false."
    )
    user = f"Cliente: {cliente}\nDevuelve el json."
    resp = _chat_json(cfg, system, user, max_tokens=120)
    if not isinstance(resp, dict):
        return None
    horario = "tarde" if str(resp.get("horario", "")).lower().startswith("t") else "manana"
    return {"horario": horario, "pmp_largo": bool(resp.get("pmp_largo"))}


# ── 4 · Leer estructura de un Excel inesperado ───────────────────────────────────
def mapear_estructura(cfg: dict, ws, max_filas: int = 25) -> Optional[dict]:
    """Fallback cuando la lectura exacta falla. Recibe una hoja openpyxl y pide al
    modelo dónde está la tabla. NO envía valores reales: solo el TIPO de cada celda
    (FECHA / PERSONA / TEXTO / VACIO) y su posición. Devuelve, en índices base-1:
      {"fila_encabezado": int, "col_fecha": int, "col_nombre": int,
       "cols_clientes": [int, ...]}  o None.
    """
    if not ia_activa(cfg):
        return None
    personas = set()
    for clave in ("rotacion_pmp", "rotacion_n2", "rotacion_n3"):
        personas.update(cfg.get(clave, []) or [])

    rejilla: List[List[str]] = []
    for r in range(1, min(max_filas, ws.max_row) + 1):
        fila: List[str] = []
        for c in range(1, min(ws.max_column, 16) + 1):
            v = ws.cell(r, c).value
            if v is None or str(v).strip() == "":
                fila.append("VACIO")
            elif hasattr(v, "date"):
                fila.append("FECHA")
            elif str(v).strip() in personas:
                fila.append("PERSONA")
            else:
                fila.append("TEXTO")
        rejilla.append(fila)

    system = (
        "Recibes una rejilla de una hoja de cálculo donde cada celda es solo su "
        "TIPO: FECHA, PERSONA (un consultor), TEXTO (p. ej. un cliente) o VACIO. "
        "Numera SIEMPRE desde 1: la 1.ª columna (A) es 1, la 2.ª (B) es 2, la 6.ª "
        "(F) es 6; la 1.ª fila es 1. Localiza la tabla de asignación "
        "semanal: una fila de encabezado con FECHAs, y debajo filas que empiezan "
        "con una PERSONA seguida de TEXTO (sus clientes). Devuelve JSON "
        "{\"fila_encabezado\": n, \"col_fecha\": n, \"col_nombre\": n, "
        "\"cols_clientes\": [n, ...]}. Si no la encuentras, todos en 0."
    )
    user = "Rejilla (fila: celdas):\n" + json.dumps(
        {str(i + 1): f for i, f in enumerate(rejilla)}, ensure_ascii=False)
    resp = _chat_json(cfg, system, user, max_tokens=400)
    if not isinstance(resp, dict):
        return None
    try:
        mapa = {
            "fila_encabezado": int(resp.get("fila_encabezado", 0)),
            "col_fecha": int(resp.get("col_fecha", 0)),
            "col_nombre": int(resp.get("col_nombre", 0)),
            "cols_clientes": [int(x) for x in resp.get("cols_clientes", [])],
        }
    except (TypeError, ValueError):
        return None
    if mapa["fila_encabezado"] <= 0 or mapa["col_nombre"] <= 0 or not mapa["cols_clientes"]:
        return None
    return mapa


def leer_asignaciones_con_mapeo(ws, mapa: dict, rotacion: List[str]) -> Dict[str, List[str]]:
    """Aplica el mapeo devuelto por mapear_estructura para extraer las
    asignaciones. Determinista: solo usa el mapa, no vuelve a llamar a la IA."""
    asig: Dict[str, List[str]] = {ing: [] for ing in rotacion}
    set_rot = set(rotacion)
    ignorar = {"", "-", "FESTIVO", "HORA DE ENVIO", *set_rot}
    col_nombre = mapa["col_nombre"]
    cols_cli = mapa["cols_clientes"]
    inicio = mapa["fila_encabezado"] + 1
    for r in range(inicio, inicio + 40):
        if r > ws.max_row:
            break
        nombre = ws.cell(r, col_nombre).value
        if nombre in set_rot:
            for c in cols_cli:
                v = ws.cell(r, c).value
                if v is None or hasattr(v, "date"):
                    continue
                cli = str(v).strip()
                if cli and cli not in ignorar and cli not in asig[nombre]:
                    asig[nombre].append(cli)
    return asig


# ── Self-test (no requiere API key: prueba anonimización y parsers) ──────────────
def _selftest() -> None:
    cfg = {
        "rotacion_pmp": ["Mateo Florez", "Heiner Diaz", "Estefania Sanabria"],
        "rotacion_n2": ["Estefania Sanabria", "Heiner Diaz", "Mateo Florez"],
        "rotacion_n3": ["Carlos Barrera"],
    }
    # IA desactivada (ia_habilitada=False) → fallback a None, sin depender de si
    # hay key embebida en el entorno de prueba.
    cfg_off = dict(cfg, ia_habilitada=False)
    assert ia_activa(cfg_off) is False
    assert redactar_aviso(cfg_off, date(2026, 6, 29), {"Heiner Diaz": ["Banco X"]}, [], "Mateo Florez", None) is None
    assert clasificar_cliente(cfg_off, "HOMI Oracle") is None

    # Anonimizador: round-trip ocultar/revelar.
    anon = Anonimizador(cfg, ["Banco X", "HOMI Oracle"])
    oculto = anon.ocultar("Mateo Florez cubre a Heiner Diaz con Banco X y HOMI Oracle")
    assert "Mateo Florez" not in oculto and "Banco X" not in oculto, oculto
    assert "Consultor A" in oculto and "Cliente 1" in oculto, oculto
    assert anon.revelar(oculto) == "Mateo Florez cubre a Heiner Diaz con Banco X y HOMI Oracle"
    assert anon.ocultar_asig({"Mateo Florez": ["Banco X"]}) == {"Consultor A": ["Cliente 1"]}

    # Colisión: un cliente llamado "A" no debe corromper el código "Consultor A"
    # (sustitución en una sola pasada).
    anon2 = Anonimizador(cfg, ["A"])
    o2 = anon2.ocultar("Mateo Florez atiende A")
    assert o2 == "Consultor A atiende Cliente 1", o2
    assert anon2.revelar(o2) == "Mateo Florez atiende A", anon2.revelar(o2)

    # Parser tolera vallas ```json y revela los nombres (respuesta simulada).
    global _chat
    cfg_fake = dict(cfg, deepseek_api_key="x")
    orig = _chat
    _chat = lambda *a, **k: '```json\n{"aviso": "Consultor A cubre a Consultor B."}\n```'  # noqa: E731
    try:
        txt = redactar_aviso(cfg_fake, date(2026, 6, 29), {"Heiner Diaz": ["Banco X"]}, [], "Mateo Florez", None)
        assert txt == "Mateo Florez cubre a Heiner Diaz.", txt
    finally:
        _chat = orig
    print("pmp_ia self-test ok")


if __name__ == "__main__":
    _selftest()
