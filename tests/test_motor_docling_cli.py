# -*- coding: utf-8 -*-
"""Contrato del CLI de motor_docling cuando Docling falla, no devuelve nada,
no está instalado o devuelve texto que no se puede codificar en UTF-8:
código de salida + mensaje legible, sin traceback, y sin pisar un .md
previo. Docling real no hace falta: se pone delante en PYTHONPATH un
paquete `docling` falso cuyo comportamiento elige cada test. Datos
sintéticos."""
import os
import subprocess
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent

DOCLING_FALSO = '''
import os
from pathlib import Path


class _Doc:
    def __init__(self, md):
        self._md = md

    def export_to_markdown(self):
        return self._md


class _Resultado:
    def __init__(self, md):
        self.document = _Doc(md)


class DocumentConverter:
    def convert(self, ruta):
        modo = Path(os.environ["DOCLING_FALSO_MODO"]).read_text(encoding="utf-8")
        if modo == "revienta":
            raise ValueError("PDF dañado: xref roto")
        if modo == "vacio":
            return _Resultado("   \\n")
        if modo == "surrogate":
            return _Resultado("texto \\udcff con byte suelto")
        return _Resultado("# Informe ñ\\n\\n| a | b |\\n")
'''


@pytest.fixture
def docling_falso(tmp_path):
    paquete = tmp_path / "falso" / "docling"
    paquete.mkdir(parents=True)
    (paquete / "__init__.py").write_text("")
    (paquete / "document_converter.py").write_text(DOCLING_FALSO, encoding="utf-8")
    modo = tmp_path / "modo.txt"

    def ejecutar(modo_docling, origen, destino):
        modo.write_text(modo_docling, encoding="utf-8")
        entorno = dict(os.environ, PYTHONPATH=str(tmp_path / "falso"),
                       DOCLING_FALSO_MODO=str(modo))
        return subprocess.run([sys.executable, str(RAIZ / "motor_docling.py"),
                               str(origen), str(destino)],
                              capture_output=True, text=True, env=entorno)
    return ejecutar


@pytest.fixture
def pdf(tmp_path):
    origen = tmp_path / "informe año.pdf"
    origen.write_bytes(b"%PDF-1.4 sintetico")
    return origen


def test_convierte_y_escribe_utf8(docling_falso, pdf, tmp_path):
    destino = tmp_path / "salida" / "informe año.md"
    r = docling_falso("ok", pdf, destino)
    assert r.returncode == 0, r.stderr
    assert destino.read_text(encoding="utf-8") == "# Informe ñ\n\n| a | b |\n"


def test_docling_que_revienta_da_mensaje_limpio_y_no_pisa(docling_falso, pdf, tmp_path):
    destino = tmp_path / "informe.md"
    destino.write_text("versión buena anterior", encoding="utf-8")
    r = docling_falso("revienta", pdf, destino)
    assert r.returncode != 0
    assert "Traceback" not in r.stderr
    assert "motor_docling falló con informe año.pdf" in r.stderr and "xref roto" in r.stderr
    assert destino.read_text(encoding="utf-8") == "versión buena anterior"


def test_docling_sin_contenido_es_un_error(docling_falso, pdf, tmp_path):
    destino = tmp_path / "informe.md"
    r = docling_falso("vacio", pdf, destino)
    assert r.returncode != 0
    assert "Traceback" not in r.stderr and "no produjo contenido" in r.stderr
    assert not destino.exists()


def test_texto_no_codificable_no_deja_el_md_a_medias(docling_falso, pdf, tmp_path):
    # Un surrogate suelto (byte mal decodificado en el PDF) no se puede
    # escribir en UTF-8: el .md anterior debe quedar intacto y sin temporales.
    destino = tmp_path / "informe.md"
    destino.write_text("versión buena anterior", encoding="utf-8")
    r = docling_falso("surrogate", pdf, destino)
    assert r.returncode != 0
    assert "Traceback" not in r.stderr and "No se pudo escribir" in r.stderr
    assert destino.read_text(encoding="utf-8") == "versión buena anterior"
    assert sorted(p.name for p in tmp_path.iterdir() if p.name.startswith(".")) == []


def test_sin_docling_instalado_da_mensaje_limpio(pdf, tmp_path):
    # Con un venv a medio instalar: ImportError → mensaje, no traceback.
    entorno = dict(os.environ, PYTHONPATH=str(tmp_path / "vacio"))
    entorno.pop("PYTHONHOME", None)
    r = subprocess.run([sys.executable, "-c",
                        "import sys, runpy; sys.modules['docling'] = None; "
                        f"sys.argv = ['motor_docling.py', {str(pdf)!r}, {str(tmp_path / 'a.md')!r}]; "
                        f"runpy.run_path({str(RAIZ / 'motor_docling.py')!r}, run_name='__main__')"],
                       capture_output=True, text=True, env=entorno, cwd=RAIZ)
    assert r.returncode != 0
    assert "Traceback" not in r.stderr and "motor_docling falló con informe año.pdf" in r.stderr
    assert not (tmp_path / "a.md").exists()


def test_motor_excel_xls_sin_libreoffice_da_error_legible(tmp_path, monkeypatch):
    import preconversion
    from motor_excel import convertir
    monkeypatch.setattr(preconversion, "buscar_soffice", lambda: None)
    xls = tmp_path / "antiguo.xls"
    xls.write_bytes(b"\xd0\xcf\x11\xe0 sintetico")
    with pytest.raises(RuntimeError, match="antiguo.xls necesita LibreOffice"):
        convertir(xls)
