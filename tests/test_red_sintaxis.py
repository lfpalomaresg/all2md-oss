# -*- coding: utf-8 -*-
"""RED DE REGRESIÓN UNIVERSAL — "todo el código del repo es Python válido y se importa".

Naturaleza: test de CARACTERIZACIÓN / red de regresión, NO TDD. Se escribe sobre código
que ya existe y se espera que pase a la primera. Su valor está en el día que deje de pasar.

Motivo: detectar errores de sintaxis o importación antes de distribuir una
versión, incluso en módulos que no se usan durante una conversión concreta.

ALCANCE: esta red solo garantiza que todo el código sea Python válido e
importable. La lógica de negocio (enrutamiento, nombres, comandos) vive en
`all2md_core.py` y la cubre `tests/test_core.py`.

GUARDA CONTRA EL FALSO VERDE: si el escaneo no encuentra ficheros .py, el test FALLA.
"""
import ast
import importlib
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent
IGNORADOS = {".git", "__pycache__", ".venv", "venv", "node_modules"}

# Mínimo de ficheros que DEBE encontrar el escaneo. Si baja de aquí, lo roto es el
# propio test (una ruta, un glob), no el código: fallar en vez de dar verde decorativo.
MINIMO_FICHEROS = 1


def _ficheros_py():
    return [
        p
        for p in RAIZ.rglob("*.py")
        if not any(parte in IGNORADOS for parte in p.relative_to(RAIZ).parts)
        and p.parent.name != "tests"
    ]


def test_el_escaneo_encuentra_ficheros():
    """Guarda contra el falso verde: sin ficheros, el resto de la red no prueba nada."""
    ficheros = _ficheros_py()
    assert len(ficheros) >= MINIMO_FICHEROS, (
        f"El escaneo de {RAIZ} devolvió {len(ficheros)} ficheros .py de aplicación. "
        "Si el repo no se ha vaciado, es este test el que está roto."
    )


@pytest.mark.parametrize("ruta", _ficheros_py(), ids=lambda p: p.name)
def test_el_fichero_es_python_valido(ruta):
    ast.parse(ruta.read_text(encoding="utf-8"), filename=str(ruta))


def test_la_app_se_importa_sin_abrir_ventana():
    """Importar ejecuta el nivel de módulo. El guard `if __name__ == "__main__"` debe
    seguir ahí: sin él, importar el fichero abriría la ventana de Tkinter y colgaría
    cualquier herramienta que lo cargue (tests, linters, empaquetado).

    Se salta (no falla) si el intérprete no trae _tkinter: el python@3.14 de
    Homebrew lo separa en el paquete python-tk -- eso es del entorno, no del
    código. La lógica de negocio vive en all2md_core y se prueba sin Tk."""
    pytest.importorskip("tkinter", reason="este Python no trae _tkinter (brew install python-tk)")
    sys.path.insert(0, str(RAIZ))
    try:
        modulo = importlib.import_module("all2md_repo_app")
        assert hasattr(modulo, "App")
        assert callable(modulo.main)
    finally:
        sys.path.remove(str(RAIZ))
