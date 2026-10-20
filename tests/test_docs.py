# -*- coding: utf-8 -*-
"""Red de deriva de documentación: el README es lo que lee quien retoma el
proyecto. Si se añade un fichero de tests o un formato enrutado y el README
no lo cuenta, esto falla."""
from pathlib import Path

import pytest

from all2md_core import EXTENSIONES_DOCLING, EXTENSIONES_EXCEL

RAIZ = Path(__file__).resolve().parent.parent
README = (RAIZ / "README.md").read_text(encoding="utf-8")


def test_el_readme_lista_cada_fichero_de_tests():
    ficheros = sorted(p.name for p in (RAIZ / "tests").glob("test_*.py"))
    assert ficheros, "no se encontró ningún test: el glob está roto"
    faltan = [f for f in ficheros if f"tests/{f}" not in README]
    assert not faltan, f"README sin documentar: {faltan}"


def test_el_readme_cuenta_cada_extension_enrutada():
    faltan = sorted(e for e in EXTENSIONES_EXCEL | EXTENSIONES_DOCLING if e not in README)
    assert not faltan, f"README no menciona: {faltan}"


def _lineas_de_enrutado(readme, motor):
    """Líneas del README que enrutan algo a `motor` ("... → ... motor.py")."""
    # Entre backticks: "motor_docling.py" a secas también casa con la fila
    # "tests/test_motor_docling.py" de la tabla de tests, que tapaba el
    # borrado de una línea de enrutado (hallazgo de careo, MEDIA, 04/10).
    return [l for l in readme.splitlines() if "→" in l and f"`{motor}.py`" in l]


def _faltan_en_enrutado(readme):
    from all2md_core import EXTENSIONES_SOFFICE, EXTENSIONES_TEXTUTIL
    esperado = {
        "motor_excel": EXTENSIONES_EXCEL | {e for e, m in EXTENSIONES_SOFFICE.items() if m == "motor_excel"},
        "motor_docling": EXTENSIONES_DOCLING | EXTENSIONES_TEXTUTIL
                         | {e for e, m in EXTENSIONES_SOFFICE.items() if m == "motor_docling"},
    }
    faltan = []
    for motor, extensiones in esperado.items():
        lineas = " ".join(_lineas_de_enrutado(readme, motor))
        faltan += [f"{e} → {motor}" for e in sorted(extensiones) if f"`{e}`" not in lineas]
    return faltan


def test_el_readme_enruta_cada_extension_a_su_motor():
    # No basta con que la extensión aparezca en algún sitio (".pdf" sale en
    # ejemplos, ".xlsx" en todas partes): tiene que estar en la línea que
    # dice a qué motor va.
    assert not _faltan_en_enrutado(README)


LINEAS_ENRUTADO = sorted({l for m in ("motor_excel", "motor_docling")
                          for l in _lineas_de_enrutado(README, m)})


def test_hay_lineas_de_enrutado():
    assert len(LINEAS_ENRUTADO) >= 3, "el README ya no tiene líneas de enrutado reconocibles"


@pytest.mark.parametrize("linea", LINEAS_ENRUTADO, ids=lambda l: l[:40])
def test_la_red_detecta_que_se_borre_cualquier_linea_de_enrutado(linea):
    # Guarda contra el falso verde: quitar UNA línea de enrutado, sea cual
    # sea, tiene que dejar alguna extensión sin motor documentado.
    sin_ella = "\n".join(l for l in README.splitlines() if l != linea)
    assert _faltan_en_enrutado(sin_ella), linea
