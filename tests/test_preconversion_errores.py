# -*- coding: utf-8 -*-
"""preconversion.py sin LibreOffice ni textutil de verdad: subprocess.run
sustituido por un doble, para probar lo que pasa cuando la herramienta
externa se cuelga, no se puede lanzar, falla o "termina bien" sin escribir
nada -- y que el comando y los nombres que se le pasan son los esperados.
Datos sintéticos."""
import subprocess
from pathlib import Path
from urllib.parse import unquote, urlparse

import pytest

import preconversion
from preconversion import TIMEOUT_SEGUNDOS, con_soffice, con_textutil


class _Resultado:
    def __init__(self, returncode=0, stdout="", stderr=""):
        self.returncode, self.stdout, self.stderr = returncode, stdout, stderr


def _run_falso(monkeypatch, efecto):
    """Sustituye subprocess.run; `efecto(cmd, **kw)` decide qué pasa.
    Devuelve la lista de llamadas (cmd, kwargs)."""
    llamadas = []

    def run(cmd, **kw):
        llamadas.append((cmd, kw))
        return efecto(cmd, **kw)

    monkeypatch.setattr(preconversion.subprocess, "run", run)
    return llamadas


def _origen(tmp_path, nombre, contenido=b"datos sinteticos"):
    origen = tmp_path / nombre
    origen.write_bytes(contenido)
    return origen


# --- _ejecutar a través de con_textutil: fallos de la herramienta externa ---

def test_timeout_da_error_legible_y_pasa_el_limite(monkeypatch, tmp_path):
    def colgado(cmd, **kw):
        raise subprocess.TimeoutExpired(cmd, kw["timeout"])
    llamadas = _run_falso(monkeypatch, colgado)
    with pytest.raises(RuntimeError, match=rf"tardó más de {TIMEOUT_SEGUNDOS}s con nota\.doc") as e:
        con_textutil(_origen(tmp_path, "nota.doc"), tmp_path / "trabajo")
    assert isinstance(e.value.__cause__, subprocess.TimeoutExpired)
    # Sin timeout, un fichero patológico colgaría el motor para siempre.
    assert llamadas[0][1]["timeout"] == TIMEOUT_SEGUNDOS


def test_herramienta_que_no_se_puede_lanzar_da_error_legible(monkeypatch, tmp_path):
    def sin_permiso(cmd, **kw):
        raise PermissionError(13, "Permission denied")
    _run_falso(monkeypatch, sin_permiso)
    with pytest.raises(RuntimeError, match=r"No se pudo lanzar textutil para nota\.rtf"):
        con_textutil(_origen(tmp_path, "nota.rtf"), tmp_path / "trabajo")


def test_salida_cero_sin_fichero_es_error_con_el_detalle_de_stdout(monkeypatch, tmp_path):
    # textutil devuelve 0 aunque no haya podido leer el fichero.
    _run_falso(monkeypatch, lambda cmd, **kw: _Resultado(0, stdout="  no pude  \n"))
    with pytest.raises(RuntimeError, match=r"textutil no pudo leer nota\.odt: no pude$"):
        con_textutil(_origen(tmp_path, "nota.odt"), tmp_path / "trabajo")


def test_codigo_de_error_aunque_haya_salida_es_error_con_stderr(monkeypatch, tmp_path):
    def falla_a_medias(cmd, **kw):
        Path(cmd[cmd.index("-output") + 1]).write_text("<html>a medias", encoding="utf-8")
        return _Resultado(1, stdout="ruido", stderr="formato corrupto")
    _run_falso(monkeypatch, falla_a_medias)
    with pytest.raises(RuntimeError, match="formato corrupto"):
        con_textutil(_origen(tmp_path, "nota.doc"), tmp_path / "trabajo")


def test_textutil_usa_nombre_de_salida_fijo_aunque_el_origen_sea_largo(monkeypatch, tmp_path):
    def escribe(cmd, **kw):
        Path(cmd[cmd.index("-output") + 1]).write_text("<p>ok</p>", encoding="utf-8")
        return _Resultado(0)
    llamadas = _run_falso(monkeypatch, escribe)
    origen = _origen(tmp_path, "ñ" * 120 + ".doc")  # 240 bytes + ".doc"
    html = con_textutil(origen, tmp_path / "trabajo" / "anidada")
    assert html == tmp_path / "trabajo" / "anidada" / "entrada.html"
    cmd = llamadas[0][0]
    assert cmd[:5] == ["textutil", "-convert", "html", "-encoding", "UTF-8"]
    assert cmd[-1] == str(origen)


# --- con_soffice: comando, copia de trabajo y nombres ---

def _soffice_falso(monkeypatch, escribir=True):
    monkeypatch.setattr(preconversion, "buscar_soffice", lambda: "/falso/soffice")

    def convierte(cmd, **kw):
        if escribir:
            entrada = Path(cmd[-1])
            formato = cmd[cmd.index("--convert-to") + 1]
            outdir = Path(cmd[cmd.index("--outdir") + 1])
            (outdir / f"{entrada.stem}.{formato}").write_bytes(entrada.read_bytes())
        return _Resultado(0)
    return _run_falso(monkeypatch, convierte)


def test_con_soffice_sin_libreoffice_da_error_legible(monkeypatch, tmp_path):
    monkeypatch.setattr(preconversion, "buscar_soffice", lambda: None)
    with pytest.raises(RuntimeError, match=r"viejo\.xls necesita LibreOffice"):
        con_soffice(_origen(tmp_path, "viejo.xls"), "xlsx", tmp_path / "trabajo")


def test_con_soffice_comando_con_perfil_propio_y_copia_del_origen(monkeypatch, tmp_path):
    llamadas = _soffice_falso(monkeypatch)
    origen = _origen(tmp_path, "Cierre Julio.XLS", b"\xd0\xcf\x11\xe0 binario")
    carpeta = tmp_path / "trabajo"
    salida = con_soffice(origen, "xlsx", carpeta)

    cmd = llamadas[0][0]
    assert cmd[0] == "/falso/soffice"
    assert "--headless" in cmd
    assert cmd[cmd.index("--convert-to") + 1] == "xlsx"
    assert cmd[cmd.index("--outdir") + 1] == str(carpeta)
    # Perfil propio dentro de la carpeta de trabajo: con el de por defecto
    # soffice no arranca si LibreOffice ya está abierto.
    perfil = next(a for a in cmd if a.startswith("-env:UserInstallation="))
    uri = perfil.split("=", 1)[1]
    assert Path(unquote(urlparse(uri).path)) == (carpeta / "perfil-lo").resolve()
    # Trabaja sobre una copia (extensión en minúsculas), nunca sobre el origen.
    entrada = Path(cmd[-1])
    assert entrada == carpeta / "Cierre Julio.xls"
    assert entrada.read_bytes() == origen.read_bytes()
    assert salida == carpeta / "Cierre Julio.xlsx"
    assert salida.exists()


def test_con_soffice_sanea_el_nombre_de_la_copia(monkeypatch, tmp_path):
    llamadas = _soffice_falso(monkeypatch)
    # Control bidi y punto inicial: no deben acabar en el nombre (que soffice
    # usa como título de hoja de un ".xls" que es CSV).
    origen = _origen(tmp_path, ".‮extracto.xls")
    salida = con_soffice(origen, "xlsx", tmp_path / "trabajo")
    assert Path(llamadas[0][0][-1]).name == "extracto.xls"
    assert salida.name == "extracto.xlsx"


def test_con_soffice_recorta_nombres_largos_sin_dejar_espacio_final(monkeypatch, tmp_path):
    llamadas = _soffice_falso(monkeypatch)
    origen = _origen(tmp_path, "a" * 99 + " " + "b" * 100 + ".xls")
    salida = con_soffice(origen, "xlsx", tmp_path / "trabajo")
    assert Path(llamadas[0][0][-1]).name == "a" * 99 + ".xls"
    assert salida.name == "a" * 99 + ".xlsx"


def test_con_soffice_nombre_sin_nada_utilizable_usa_entrada(monkeypatch, tmp_path):
    llamadas = _soffice_falso(monkeypatch)
    # Solo invisibles: el saneado no deja nada.
    origen = _origen(tmp_path, "​　.xls")
    salida = con_soffice(origen, "xlsx", tmp_path / "trabajo")
    assert Path(llamadas[0][0][-1]).name == "entrada.xls"
    assert salida.name == "entrada.xlsx"


def test_con_soffice_que_no_escribe_salida_es_error(monkeypatch, tmp_path):
    # soffice devuelve 0 también cuando no reconoce el formato.
    _soffice_falso(monkeypatch, escribir=False)
    with pytest.raises(RuntimeError, match=r"LibreOffice no pudo leer raro\.ppt"):
        con_soffice(_origen(tmp_path, "raro.ppt"), "pptx", tmp_path / "trabajo")
