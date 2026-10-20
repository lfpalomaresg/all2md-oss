# -*- coding: utf-8 -*-
"""Frontera con las herramientas externas (textutil, LibreOffice): cómo se
comporta preconversion cuando la herramienta se cuelga, no se puede lanzar,
falla, o "acaba bien" sin escribir nada. Sin LibreOffice real: cada test
usa un ejecutable falso (script de shell) que simula el caso. Datos
sintéticos."""
import errno
import os
import stat

import pytest

import preconversion
from preconversion import _ejecutar, con_soffice


def _script(tmp_path, nombre, cuerpo):
    ruta = tmp_path / nombre
    ruta.write_text("#!/bin/sh\n" + cuerpo, encoding="utf-8")
    ruta.chmod(ruta.stat().st_mode | stat.S_IXUSR)
    return ruta


def _origen(tmp_path, nombre="libro.xls", contenido=b"binario sintetico"):
    ruta = tmp_path / nombre
    ruta.write_bytes(contenido)
    return ruta


def test_herramienta_colgada_corta_por_timeout_con_mensaje(tmp_path, monkeypatch):
    lento = _script(tmp_path, "lento", "sleep 30\n")
    monkeypatch.setattr(preconversion, "TIMEOUT_SEGUNDOS", 0.5)
    origen = _origen(tmp_path)
    with pytest.raises(RuntimeError, match=r"tardó más de 0\.5s con libro\.xls"):
        _ejecutar([str(lento)], origen, tmp_path / "salida.xlsx", "LibreOffice")


def test_herramienta_que_no_se_puede_lanzar_da_runtimeerror(tmp_path):
    # Un soffice sin permiso de ejecución (instalación rota) lanza OSError
    # en subprocess.run: debe llegar como RuntimeError legible, no traceback.
    roto = tmp_path / "soffice"
    roto.write_text("#!/bin/sh\n", encoding="utf-8")
    roto.chmod(0o644)
    with pytest.raises(RuntimeError, match="No se pudo lanzar LibreOffice para libro.xls"):
        _ejecutar([str(roto)], _origen(tmp_path), tmp_path / "salida.xlsx", "LibreOffice")


def test_codigo_cero_sin_fichero_de_salida_es_un_fallo(tmp_path):
    # soffice devuelve 0 cuando no reconoce el formato: lo que cuenta es
    # que exista la salida.
    mudo = _script(tmp_path, "mudo", "echo 'Error: source file could not be loaded' >&2\nexit 0\n")
    with pytest.raises(RuntimeError, match="could not be loaded"):
        _ejecutar([str(mudo)], _origen(tmp_path), tmp_path / "salida.xlsx", "LibreOffice")


def test_codigo_distinto_de_cero_falla_aunque_haya_salida(tmp_path):
    # Una salida a medias de un proceso que ha fallado no se da por buena.
    salida = tmp_path / "salida.xlsx"
    salida.write_bytes(b"a medias")
    falla = _script(tmp_path, "falla", "echo 'disco lleno' >&2\nexit 3\n")
    with pytest.raises(RuntimeError, match="disco lleno"):
        _ejecutar([str(falla)], _origen(tmp_path), salida, "LibreOffice")


def test_sin_stderr_usa_stdout_como_detalle(tmp_path):
    falla = _script(tmp_path, "falla", "echo 'detalle por stdout'\nexit 1\n")
    with pytest.raises(RuntimeError, match="detalle por stdout"):
        _ejecutar([str(falla)], _origen(tmp_path), tmp_path / "x.xlsx", "LibreOffice")


def test_con_soffice_sin_libreoffice_lo_dice(tmp_path, monkeypatch):
    monkeypatch.setattr(preconversion, "buscar_soffice", lambda: None)
    with pytest.raises(RuntimeError, match="libro.xls necesita LibreOffice"):
        con_soffice(_origen(tmp_path), "xlsx", tmp_path / "trabajo")


def _soffice_falso(tmp_path):
    """soffice que apunta sus argumentos y escribe <outdir>/<stem>.<formato>
    como haría LibreOffice."""
    registro = tmp_path / "argv.txt"
    falso = _script(tmp_path, "soffice", f"""
for a in "$@"; do printf '%s\\n' "$a"; done > "{registro}"
formato=""; outdir=""; entrada=""
while [ $# -gt 0 ]; do
  case "$1" in
    --convert-to) formato="$2"; shift 2;;
    --outdir) outdir="$2"; shift 2;;
    -*) shift;;
    *) entrada="$1"; shift;;
  esac
done
base=$(basename "$entrada"); base="${{base%.*}}"
printf 'convertido' > "$outdir/$base.$formato"
""")
    return falso, registro


def test_con_soffice_convierte_una_copia_con_perfil_propio(tmp_path, monkeypatch):
    falso, registro = _soffice_falso(tmp_path)
    monkeypatch.setattr(preconversion, "buscar_soffice", lambda: str(falso))
    origen = _origen(tmp_path, "Cierre Julio.XLS")
    carpeta = tmp_path / "trabajo"

    salida = con_soffice(origen, "xlsx", carpeta)

    assert salida == carpeta / "Cierre Julio.xlsx"
    assert salida.read_text() == "convertido"
    argv = registro.read_text(encoding="utf-8").splitlines()
    # Perfil propio (si no, soffice no arranca con LibreOffice abierto) y
    # conversión de la COPIA, nunca del original.
    assert argv[0].startswith("-env:UserInstallation=file://")
    assert "--headless" in argv
    assert argv[-1] == str(carpeta / "Cierre Julio.xls")
    assert origen.read_bytes() == b"binario sintetico"


@pytest.mark.parametrize("nombre", [
    "..xls",                 # stem "." : nada utilizable
    ".oculto.xls",           # no debe quedar un fichero punto
    "a\nb.xls",              # control dentro del nombre
    ("ñ" * 200) + ".xls",    # 400 bytes: superaría NAME_MAX
])
def test_con_soffice_nombre_de_la_copia_saneado(tmp_path, monkeypatch, nombre):
    falso, registro = _soffice_falso(tmp_path)
    monkeypatch.setattr(preconversion, "buscar_soffice", lambda: str(falso))
    try:
        origen = _origen(tmp_path, nombre)
    except OSError as exc:
        if exc.errno == errno.ENAMETOOLONG:
            pytest.skip("El sistema de archivos no admite el nombre largo de origen")
        raise
    carpeta = tmp_path / "trabajo"

    salida = con_soffice(origen, "xlsx", carpeta)

    assert salida.parent == carpeta
    assert not salida.name.startswith(".")
    assert "\n" not in salida.name
    assert len(salida.name.encode("utf-8")) <= 255
    assert salida.exists()


@pytest.mark.skipif(os.geteuid() == 0, reason="root ignora los permisos")
def test_con_soffice_origen_sin_permiso_de_lectura(tmp_path, monkeypatch):
    falso, _ = _soffice_falso(tmp_path)
    monkeypatch.setattr(preconversion, "buscar_soffice", lambda: str(falso))
    origen = _origen(tmp_path, "privado.xls")
    origen.chmod(0)
    try:
        with pytest.raises(RuntimeError, match="No se pudo leer privado.xls"):
            con_soffice(origen, "xlsx", tmp_path / "trabajo")
    finally:
        origen.chmod(0o644)
