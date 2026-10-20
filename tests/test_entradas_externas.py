# -*- coding: utf-8 -*-
"""Entradas que llegan de fuera y no estaban cubiertas: la ruta .xls de
motor_excel (con LibreOffice sustituido por un doble), referencias de
matriciales malformadas en un libro hostil, y nombres/contenidos con
codificaciones raras en la cabecera de procedencia. Datos sintéticos."""
import datetime as dt
from pathlib import Path

import pytest
from openpyxl import Workbook

import motor_excel
from all2md_core import anadir_frontmatter, con_frontmatter
from motor_excel import ConversionVacia, convertir

FECHA = dt.date(2026, 10, 15)


# --- motor_excel: .xls vía LibreOffice ---

def _soffice_que_genera(monkeypatch, construir):
    """con_soffice falso: escribe en la carpeta de trabajo el .xlsx que
    devuelve `construir()`. Devuelve la lista de llamadas."""
    llamadas = []

    def falso(origen, formato, carpeta):
        llamadas.append((origen, formato, carpeta))
        salida = carpeta / "entrada.xlsx"
        construir().save(salida)
        return salida
    monkeypatch.setattr(motor_excel, "con_soffice", falso)
    return llamadas


def test_xls_se_convierte_via_xlsx_en_temporal_que_se_borra(monkeypatch, tmp_path):
    def libro():
        wb = Workbook()
        wb.active.title = "Cierre"
        wb.active.append(["Rooms", 100])
        return wb
    llamadas = _soffice_que_genera(monkeypatch, libro)
    origen = tmp_path / "viejo.XLS"
    origen.write_bytes(b"\xd0\xcf\x11\xe0")

    md, avisos = convertir(origen)
    assert "## Cierre" in md and "| Rooms | 100 |" in md
    assert avisos == 0
    (o, formato, carpeta), = llamadas
    assert o == origen and formato == "xlsx"
    assert carpeta.name.startswith("all2md-")
    assert not carpeta.exists()


def test_xls_vacio_nombra_el_documento_original_no_el_temporal(monkeypatch, tmp_path):
    _soffice_que_genera(monkeypatch, Workbook)
    origen = tmp_path / "extracto banco.xls"
    origen.write_bytes(b"\xd0\xcf\x11\xe0")
    with pytest.raises(ConversionVacia, match=r"extracto banco\.xls") as e:
        convertir(origen)
    assert "entrada" not in str(e.value)


def test_xls_error_de_libreoffice_se_propaga_legible(monkeypatch, tmp_path):
    def falla(origen, formato, carpeta):
        raise RuntimeError(f"{origen.name} necesita LibreOffice, que no está instalado.")
    monkeypatch.setattr(motor_excel, "con_soffice", falla)
    with pytest.raises(RuntimeError, match=r"a\.xls necesita LibreOffice"):
        convertir(tmp_path / "a.xls")


# --- motor_excel: matricial con referencia malformada ---

@pytest.mark.parametrize("ref", ["no-es-un-rango", "A0:B2", ""])
def test_matricial_con_ref_malformada_no_rompe_ni_expande(ref):
    wb = Workbook()
    hoja = wb.active
    ancla = hoja["B2"]
    marcadas = {}
    grandes = []
    presupuesto, f_max, c_max = motor_excel._expandir_matricial(
        hoja, ancla, ref, marcadas, 50, grandes)
    assert (presupuesto, f_max, c_max) == (50, 0, 0)
    assert marcadas == {} and grandes == []


# --- Cabecera de procedencia: codificaciones ---

@pytest.mark.parametrize("nombre", [
    "C:\\docs\\cierre.pdf",      # barra invertida: escape YAML \\
    "fin\\",                      # barra al final, antes de la comilla
    "a\\\"b.pdf",                 # barra + comilla
])
def test_frontmatter_con_barras_invertidas_es_yaml_valido(nombre):
    yaml = pytest.importorskip("yaml")
    md = con_frontmatter("cuerpo\n", Path(nombre), "MarkItDown", FECHA)
    cabecera = md.split("---\n")[1]
    assert yaml.safe_load(cabecera)["fuente"] == Path(nombre).name


def test_frontmatter_con_barras_invertidas_se_sustituye_sin_apilar():
    previo = con_frontmatter("cuerpo\n", Path("C:\\x\\a.pdf"), "MarkItDown", FECHA)
    md = con_frontmatter(previo, Path("C:\\x\\a.pdf"), "motor_docling", FECHA)
    assert md.count("fuente:") == 1
    assert md.endswith("---\n\ncuerpo\n")
    assert 'motor: "motor_docling"' in md


def test_anadir_frontmatter_con_salida_no_utf8_no_toca_el_fichero(tmp_path):
    # Un conversor que escribe en Latin-1: la app captura UnicodeError y
    # deja el .md del motor tal cual (sin cabecera, pero sin perderlo).
    destino = tmp_path / "a.md"
    destino.write_bytes("Año 2026: ocupación".encode("latin-1"))
    with pytest.raises(UnicodeError):
        anadir_frontmatter(destino, Path("a.pdf"), "markitdown", FECHA)
    assert destino.read_bytes() == "Año 2026: ocupación".encode("latin-1")
    assert list(tmp_path.iterdir()) == [destino]


def test_anadir_frontmatter_quita_bom_y_unifica_saltos_crlf(tmp_path):
    # Salida de un conversor con BOM y CRLF: el BOM no queda incrustado
    # tras la cabecera y los saltos quedan todos en LF (sin mezclar).
    destino = tmp_path / "a.md"
    destino.write_bytes("\ufeff# Título\r\nlínea\r\n".encode("utf-8"))
    anadir_frontmatter(destino, Path("a.docx"), "motor_docling", FECHA)
    texto = destino.read_bytes().decode("utf-8")
    assert texto.startswith("---\nfuente: \"a.docx\"\n")
    assert "\ufeff" not in texto and "\r" not in texto
    assert texto.endswith("---\n\n# Título\nlínea\n")
