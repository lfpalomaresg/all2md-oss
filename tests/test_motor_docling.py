# -*- coding: utf-8 -*-
"""motor_docling: preconversión de Word antiguo/RTF/ODT con textutil
(incluido en macOS, gratis) a HTML, que Docling sí lee. HTML y no .docx:
el .docx que escribe textutil aplana las tablas en párrafos sueltos, y
reconstruir tablas es justo para lo que se usa Docling. Sin importar
docling. Datos sintéticos."""
import shutil
import subprocess

import pytest

from motor_docling import preconvertir

pytestmark = pytest.mark.skipif(shutil.which("textutil") is None,
                                reason="textutil solo existe en macOS")

HTML = ("<h1>Informe sintético ñ</h1><table border=1>"
        "<tr><td>Concepto</td><td>Importe</td></tr>"
        "<tr><td>Rooms</td><td>100</td></tr></table>")


def _documento(tmp_path, formato):
    fuente = tmp_path / "fuente.html"
    fuente.write_text(HTML, encoding="utf-8")
    origen = tmp_path / f"nota.{formato}"
    subprocess.run(["textutil", "-inputencoding", "UTF-8", "-convert", formato,
                    "-output", str(origen), str(fuente)], check=True)
    return origen


@pytest.mark.parametrize("formato", ["doc", "rtf", "odt"])
def test_preconvertir_conserva_texto_y_tablas(tmp_path, formato):
    html = preconvertir(_documento(tmp_path, formato), tmp_path / "trabajo")
    assert html.suffix == ".html"
    texto = html.read_text(encoding="utf-8")
    assert "Informe sintético ñ" in texto
    assert "<table" in texto and "Rooms" in texto


def test_preconvertir_deja_pasar_lo_que_docling_ya_lee(tmp_path):
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"%PDF-1.4")
    assert preconvertir(pdf, tmp_path / "trabajo") == pdf


def test_preconvertir_error_de_textutil_da_mensaje_legible(tmp_path):
    # textutil no falla con basura (la lee como texto plano) y devuelve 0
    # incluso si el fichero no existe: se comprueba que haya salida.
    with pytest.raises(RuntimeError, match="fantasma.doc"):
        preconvertir(tmp_path / "fantasma.doc", tmp_path / "trabajo")
