# -*- coding: utf-8 -*-
"""motor_docling.convertir sin Docling instalado: se inyecta un módulo
`docling.document_converter` falso para probar la orquestación (qué
fichero recibe Docling, limpieza del temporal, salida vacía) y el
enrutado de preconvertir. Datos sintéticos."""
import sys
import types
from pathlib import Path

import pytest

import motor_docling
from motor_docling import ConversionVacia, convertir, preconvertir


@pytest.fixture
def docling_falso(monkeypatch):
    """Registra un DocumentConverter falso. Devuelve un dict con lo que
    recibió convert() y con el markdown que devolverá."""
    estado = {"md": "# Informe sintético\n", "recibido": None, "existia": None}

    class DocumentConverter:
        def convert(self, ruta):
            estado["recibido"] = Path(ruta)
            estado["existia"] = Path(ruta).exists()
            documento = types.SimpleNamespace(export_to_markdown=lambda: estado["md"])
            return types.SimpleNamespace(document=documento)

    paquete = types.ModuleType("docling")
    modulo = types.ModuleType("docling.document_converter")
    modulo.DocumentConverter = DocumentConverter
    paquete.document_converter = modulo
    monkeypatch.setitem(sys.modules, "docling", paquete)
    monkeypatch.setitem(sys.modules, "docling.document_converter", modulo)
    return estado


def test_convertir_pdf_pasa_el_origen_tal_cual(docling_falso, tmp_path):
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"%PDF-1.4")
    assert convertir(pdf) == "# Informe sintético\n"
    assert docling_falso["recibido"] == pdf


@pytest.mark.parametrize("md", ["", "  \n\t\n"])
def test_convertir_sin_contenido_lanza_conversion_vacia(docling_falso, tmp_path, md):
    docling_falso["md"] = md
    pdf = tmp_path / "escaneo.pdf"
    pdf.write_bytes(b"%PDF-1.4")
    with pytest.raises(ConversionVacia, match=r"escaneo\.pdf"):
        convertir(pdf)


def test_convertir_rtf_preconvierte_en_temporal_y_lo_borra(docling_falso, monkeypatch, tmp_path):
    def textutil_falso(origen, carpeta):
        html = carpeta / "entrada.html"
        html.write_text("<p>hola</p>", encoding="utf-8")
        return html
    monkeypatch.setattr(motor_docling, "con_textutil", textutil_falso)
    rtf = tmp_path / "nota.RTF"
    rtf.write_text("{\\rtf1 hola}", encoding="utf-8")

    assert convertir(rtf) == "# Informe sintético\n"
    recibido = docling_falso["recibido"]
    assert recibido.name == "entrada.html" and docling_falso["existia"]
    assert recibido.parent.name.startswith("all2md-")
    # El temporal (con la copia del documento) no se queda en disco.
    assert not recibido.parent.exists()


def test_convertir_borra_el_temporal_aunque_docling_falle(docling_falso, monkeypatch, tmp_path):
    carpetas = []

    def textutil_falso(origen, carpeta):
        carpetas.append(carpeta)
        html = carpeta / "entrada.html"
        html.write_text("<p>x</p>", encoding="utf-8")
        return html

    def revienta(self, ruta):
        raise ValueError("documento corrupto")

    monkeypatch.setattr(motor_docling, "con_textutil", textutil_falso)
    monkeypatch.setattr(sys.modules["docling.document_converter"].DocumentConverter,
                        "convert", revienta)
    doc = tmp_path / "nota.doc"
    doc.write_bytes(b"\xd0\xcf\x11\xe0")
    with pytest.raises(ValueError, match="documento corrupto"):
        convertir(doc)
    assert carpetas and not carpetas[0].exists()


@pytest.mark.parametrize("nombre", ["diapos.ppt", "DIAPOS.PPT"])
def test_preconvertir_ppt_va_a_libreoffice_como_pptx(monkeypatch, tmp_path, nombre):
    llamadas = []
    monkeypatch.setattr(motor_docling, "con_soffice",
                        lambda o, f, c: llamadas.append((o, f, c)) or c / "x.pptx")
    origen = tmp_path / nombre
    assert preconvertir(origen, tmp_path / "t") == tmp_path / "t" / "x.pptx"
    assert llamadas == [(origen, "pptx", tmp_path / "t")]


@pytest.mark.parametrize("nombre", ["a.DOC", "a.Rtf", "a.ODT"])
def test_preconvertir_extension_en_mayusculas_va_a_textutil(monkeypatch, tmp_path, nombre):
    monkeypatch.setattr(motor_docling, "con_textutil", lambda o, c: c / "entrada.html")
    assert preconvertir(tmp_path / nombre, tmp_path / "t") == tmp_path / "t" / "entrada.html"


@pytest.mark.parametrize("nombre", ["a.docx", "a.pptx", "a.PDF"])
def test_preconvertir_no_toca_lo_que_docling_lee(monkeypatch, tmp_path, nombre):
    def no_llamar(*a):
        raise AssertionError("no debía preconvertir")
    monkeypatch.setattr(motor_docling, "con_textutil", no_llamar)
    monkeypatch.setattr(motor_docling, "con_soffice", no_llamar)
    assert preconvertir(tmp_path / nombre, tmp_path / "t") == tmp_path / nombre
