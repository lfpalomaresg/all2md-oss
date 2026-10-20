"""
Preconversión de formatos que ningún motor lee directamente, con
herramientas gratuitas del propio equipo:

- textutil (viene con macOS): .doc/.rtf/.odt → HTML, que Docling sí lee.
  HTML y no .docx: el .docx de textutil aplana las tablas en párrafos.
- LibreOffice (soffice, si está instalado): .xls → .xlsx (motor_excel,
  conserva celdas combinadas) y .ppt → .pptx (Docling).

Sin dependencias de terceros: lo importan los motores (venv propio) y la
app (para saber si ofrecer .xls/.ppt).
"""
from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path

from all2md_core import nombre_seguro

# Lanzada desde Finder, la app no hereda el PATH de Homebrew: se miran
# también las rutas de instalación habituales en macOS.
RUTAS_SOFFICE = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/opt/homebrew/bin/soffice",
    "/usr/local/bin/soffice",
]

# Un fichero patológico no debe dejar el motor colgado para siempre.
TIMEOUT_SEGUNDOS = 300


def buscar_soffice() -> str | None:
    encontrado = shutil.which("soffice")
    if encontrado:
        return encontrado
    for ruta in RUTAS_SOFFICE:
        if os.path.isfile(ruta) and os.access(ruta, os.X_OK):
            return ruta
    return None


def _ejecutar(cmd: list[str], origen: Path, salida: Path, herramienta: str) -> Path:
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=TIMEOUT_SEGUNDOS)
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f"{herramienta} tardó más de {TIMEOUT_SEGUNDOS}s con {origen.name}") from e
    except OSError as e:
        raise RuntimeError(f"No se pudo lanzar {herramienta} para {origen.name}: {e}") from e
    # textutil devuelve 0 incluso si no pudo leer el fichero, y soffice
    # también cuando no reconoce el formato: lo que cuenta es la salida.
    if r.returncode != 0 or not salida.exists():
        detalle = (r.stderr or r.stdout).strip()
        raise RuntimeError(f"{herramienta} no pudo leer {origen.name}: {detalle}")
    return salida


def con_textutil(origen: Path, carpeta: Path) -> Path:
    """.doc/.rtf/.odt → HTML. Nombre de salida fijo: "{stem}.html" es un
    byte más largo que "{stem}.doc" y con nombres de 255 bytes superaba
    el máximo del sistema de ficheros (hallazgo de careo, BAJA, 04/10)."""
    carpeta.mkdir(parents=True, exist_ok=True)
    html = carpeta / "entrada.html"
    return _ejecutar(["textutil", "-convert", "html", "-encoding", "UTF-8",
                      "-output", str(html), str(origen)], origen, html, "textutil")


def con_soffice(origen: Path, formato: str, carpeta: Path) -> Path:
    """Convierte con LibreOffice a `formato` ("xlsx", "pptx"). Trabaja
    sobre una copia con nombre controlado (soffice nombra la salida como la
    entrada) y con un perfil de usuario propio: con el perfil por defecto,
    soffice no arranca si LibreOffice ya está abierto."""
    soffice = buscar_soffice()
    if soffice is None:
        raise RuntimeError(f"{origen.name} necesita LibreOffice, que no está instalado.")
    carpeta.mkdir(parents=True, exist_ok=True)
    # El nombre de la copia acaba como título de hoja cuando soffice importa
    # un ".xls" que en realidad es CSV/HTML (exportaciones de bancos): el
    # del documento, saneado y recortado (cabe de sobra en NAME_MAX), no
    # "entrada" (hallazgo de careo, BAJA, 04/10).
    base = nombre_seguro(origen.stem) or "entrada"
    base = base.encode("utf-8")[:100].decode("utf-8", "ignore").rstrip(" .") or "entrada"
    entrada = carpeta / f"{base}{origen.suffix.lower()}"
    try:
        shutil.copyfile(origen, entrada)
    except OSError as e:
        raise RuntimeError(f"No se pudo leer {origen.name}: {e}") from e
    salida = carpeta / f"{base}.{formato}"
    perfil = (carpeta / "perfil-lo").resolve().as_uri()
    return _ejecutar([soffice, f"-env:UserInstallation={perfil}", "--headless",
                      "--convert-to", formato, "--outdir", str(carpeta), str(entrada)],
                     origen, salida, "LibreOffice")
