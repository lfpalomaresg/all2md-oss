# -*- coding: utf-8 -*-
"""Los motores se invocan desde la app como subprocess: su contrato es el
código de salida + un mensaje legible (nunca un traceback crudo, que la
app mostraría tal cual en el diálogo de error). Datos sintéticos."""
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

RAIZ = Path(__file__).resolve().parent.parent


def _motor_excel(*args):
    return subprocess.run([sys.executable, str(RAIZ / "motor_excel.py"), *map(str, args)],
                          capture_output=True, text=True)


def _xlsx(tmp_path, nombre="a.xlsx"):
    wb = Workbook()
    wb.active["A1"] = "hola"
    ruta = tmp_path / nombre
    wb.save(ruta)
    return ruta


def test_sin_argumentos_explica_el_uso():
    r = _motor_excel()
    assert r.returncode != 0 and "Uso" in r.stderr


def test_origen_inexistente(tmp_path):
    r = _motor_excel(tmp_path / "no.xlsx", tmp_path / "a.md")
    assert r.returncode != 0 and "No existe" in r.stderr


def test_xlsx_corrupto_da_mensaje_limpio(tmp_path):
    corrupto = tmp_path / "roto.xlsx"
    corrupto.write_bytes(b"esto no es un zip")
    r = _motor_excel(corrupto, tmp_path / "a.md")
    assert r.returncode != 0
    assert "Traceback" not in r.stderr and "roto.xlsx" in r.stderr
    assert not (tmp_path / "a.md").exists()


def test_convierte_y_escribe_el_destino(tmp_path):
    r = _motor_excel(_xlsx(tmp_path), tmp_path / "a.md")
    assert r.returncode == 0, r.stderr
    assert "| hola |" in (tmp_path / "a.md").read_text(encoding="utf-8")


def test_crea_la_carpeta_destino_si_no_existe(tmp_path):
    destino = tmp_path / "nueva" / "sub" / "a.md"
    r = _motor_excel(_xlsx(tmp_path), destino)
    assert r.returncode == 0, r.stderr
    assert destino.exists()


def test_destino_no_escribible_da_mensaje_limpio(tmp_path):
    destino = tmp_path / "es_un_directorio.md"
    destino.mkdir()
    r = _motor_excel(_xlsx(tmp_path), destino)
    assert r.returncode != 0
    assert "Traceback" not in r.stderr and "No se pudo escribir" in r.stderr


@pytest.mark.parametrize("args", [[], ["solo_origen"]])
def test_motor_docling_sin_argumentos_explica_el_uso(args):
    # Sin importar docling: el error de uso sale antes del import diferido.
    r = subprocess.run([sys.executable, str(RAIZ / "motor_docling.py"), *args],
                       capture_output=True, text=True)
    assert r.returncode != 0 and "Uso" in r.stderr


# --- Automejora 2, loop 18: destino igual al origen ---
def test_motor_excel_no_escribe_encima_del_origen(tmp_path):
    origen = _xlsx(tmp_path)
    antes = origen.read_bytes()
    r = _motor_excel(origen, tmp_path / "." / origen.name)
    assert r.returncode != 0
    assert "Traceback" not in r.stderr and "origen" in r.stderr
    assert origen.read_bytes() == antes


def test_motor_docling_no_escribe_encima_del_origen(tmp_path):
    # La comprobación va antes del import diferido de docling.
    origen = tmp_path / "a.pdf"
    origen.write_bytes(b"%PDF-1.4 sintetico")
    r = subprocess.run([sys.executable, str(RAIZ / "motor_docling.py"), str(origen), str(origen)],
                       capture_output=True, text=True)
    assert r.returncode != 0 and "origen" in r.stderr
    assert origen.read_bytes() == b"%PDF-1.4 sintetico"
