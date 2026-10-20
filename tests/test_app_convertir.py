# -*- coding: utf-8 -*-
"""Orquestación de "Convertir y guardar" (App.convert): qué hace la app
con un .md que ya existe, con un motor que falla o no se puede lanzar, y
con una salida que no es UTF-8. Es donde un fallo cuesta más: pisar un .md
retocado a mano o dar por buena una conversión que no ha ocurrido.

Sin ventana y sin Tk real: tkinter se sustituye por un doble mínimo (el
Python de Homebrew puede venir sin _tkinter) y subprocess.run por un
motor falso. Datos sintéticos."""
import importlib
import subprocess
import sys
import types
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent


class _Var:
    def __init__(self, value=""):
        self._v = value

    def get(self):
        return self._v

    def set(self, v):
        self._v = v


class _MessageBox(types.ModuleType):
    """Doble de tkinter.messagebox: apunta cada diálogo y responde lo que
    el test haya preparado en `respuesta`."""
    NO = "no"

    def __init__(self):
        super().__init__("tkinter.messagebox")
        self.llamadas = []
        self.respuesta = None

    def _apunta(self, tipo, *args, **kwargs):
        self.llamadas.append((tipo, args, kwargs))

    def showerror(self, *a, **k):
        self._apunta("error", *a, **k)

    def showinfo(self, *a, **k):
        self._apunta("info", *a, **k)

    def askyesnocancel(self, *a, **k):
        self._apunta("pregunta", *a, **k)
        return self.respuesta

    def tipos(self):
        return [t for t, _, _ in self.llamadas]


@pytest.fixture
def entorno(monkeypatch, tmp_path):
    """Importa la app con un tkinter falso y devuelve (modulo, App, messagebox)."""
    tk = types.ModuleType("tkinter")
    tk.Tk = object
    tk.StringVar = _Var
    mb = _MessageBox()
    fd = types.ModuleType("tkinter.filedialog")
    tk.messagebox, tk.filedialog = mb, fd
    monkeypatch.setitem(sys.modules, "tkinter", tk)
    monkeypatch.setitem(sys.modules, "tkinter.messagebox", mb)
    monkeypatch.setitem(sys.modules, "tkinter.filedialog", fd)
    monkeypatch.syspath_prepend(str(RAIZ))
    monkeypatch.delitem(sys.modules, "all2md_repo_app", raising=False)
    modulo = importlib.import_module("all2md_repo_app")
    monkeypatch.delitem(sys.modules, "all2md_repo_app")

    # El venv de motores "existe" (un fichero cualquiera basta: no se lanza).
    python_falso = tmp_path / "venv-python"
    python_falso.write_text("")
    monkeypatch.setattr(modulo, "VENV_TOOLS_PY", python_falso)
    return modulo, mb


def _app(modulo, origen, repo, nombre=""):
    app = modulo.App.__new__(modulo.App)  # sin build_ui: no hay ventana
    app.pdf_path = _Var(str(origen))
    app.repo_path = _Var(str(repo))
    app.filename = _Var(nombre)
    app.status = _Var("")
    app.hay_soffice = False
    return app


def _motor_que_escribe(contenido: bytes = b"| hola |\n", llamadas=None):
    """subprocess.run falso: escribe `contenido` en el destino (último
    argumento, o el que sigue a -o) como haría el motor."""
    def run(cmd, **kwargs):
        if llamadas is not None:
            llamadas.append(cmd)
        destino = Path(cmd[cmd.index("-o") + 1] if "-o" in cmd else cmd[-1])
        destino.write_bytes(contenido)
        return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="")
    return run


@pytest.fixture
def documento(tmp_path):
    origen = tmp_path / "docs" / "cierre.xlsx"
    origen.parent.mkdir()
    origen.write_bytes(b"xlsx sintetico")
    return origen


def test_convierte_y_anade_la_cabecera_de_procedencia(entorno, documento, tmp_path, monkeypatch):
    modulo, mb = entorno
    repo = tmp_path / "repo"
    llamadas = []
    monkeypatch.setattr(modulo.subprocess, "run", _motor_que_escribe(llamadas=llamadas))
    app = _app(modulo, documento, repo)

    app.convert()

    destino = repo / "cierre.md"
    texto = destino.read_text(encoding="utf-8")
    assert texto.startswith('---\nfuente: "cierre.xlsx"\nmotor: "motor_excel"\n')
    assert texto.endswith("| hola |\n")
    # El motor escribe en un temporal dentro de la carpeta repo (mismo disco
    # para el os.replace), no en el destino.
    assert Path(llamadas[0][-1]).parent.parent == repo and llamadas[0][-1] != str(destino)
    assert sorted(p.name for p in repo.iterdir()) == ["cierre.md"]
    assert mb.tipos() == ["info"]
    assert app.filename.get() == "cierre"  # nombre vacío: se rellena


def test_md_existente_y_cancelar_no_toca_nada(entorno, documento, tmp_path, monkeypatch):
    modulo, mb = entorno
    repo = tmp_path / "repo"
    repo.mkdir()
    previo = repo / "cierre.md"
    previo.write_text("retocado a mano", encoding="utf-8")
    llamadas = []
    monkeypatch.setattr(modulo.subprocess, "run", _motor_que_escribe(llamadas=llamadas))
    mb.respuesta = None  # Cancelar

    _app(modulo, documento, repo).convert()

    assert llamadas == []
    assert previo.read_text(encoding="utf-8") == "retocado a mano"
    # Enter (respuesta por defecto) no debe pisar nada.
    pregunta = [k for t, _, k in mb.llamadas if t == "pregunta"][0]
    assert pregunta["default"] == mb.NO


def test_md_existente_y_no_guarda_una_copia_numerada(entorno, documento, tmp_path, monkeypatch):
    modulo, mb = entorno
    repo = tmp_path / "repo"
    repo.mkdir()
    (repo / "cierre.md").write_text("retocado a mano", encoding="utf-8")
    (repo / "cierre-2.md").write_text("otra copia", encoding="utf-8")
    monkeypatch.setattr(modulo.subprocess, "run", _motor_que_escribe())
    mb.respuesta = False  # No → copia

    _app(modulo, documento, repo).convert()

    assert (repo / "cierre.md").read_text(encoding="utf-8") == "retocado a mano"
    assert (repo / "cierre-2.md").read_text(encoding="utf-8") == "otra copia"
    copia = (repo / "cierre-3.md").read_text(encoding="utf-8")
    assert copia.startswith("---\nfuente:") and "| hola |" in copia


def test_md_existente_y_si_lo_sobrescribe(entorno, documento, tmp_path, monkeypatch):
    modulo, mb = entorno
    repo = tmp_path / "repo"
    repo.mkdir()
    (repo / "cierre.md").write_text("viejo", encoding="utf-8")
    monkeypatch.setattr(modulo.subprocess, "run", _motor_que_escribe())
    mb.respuesta = True

    _app(modulo, documento, repo).convert()

    assert "| hola |" in (repo / "cierre.md").read_text(encoding="utf-8")
    assert not (repo / "cierre-2.md").exists()
    assert sorted(p.name for p in repo.iterdir()) == ["cierre.md"]  # sin temporales


def test_sobrescribir_y_motor_que_acaba_bien_sin_escribir_conserva_el_previo(
        entorno, documento, tmp_path, monkeypatch):
    # El .md ya existía: que exista al terminar no prueba que el motor lo
    # haya escrito. No debe darse por convertido ni recibir otra cabecera.
    modulo, mb = entorno
    repo = tmp_path / "repo"
    repo.mkdir()
    (repo / "cierre.md").write_text("retocado a mano", encoding="utf-8")
    monkeypatch.setattr(modulo.subprocess, "run",
                        lambda cmd, **k: subprocess.CompletedProcess(cmd, 0, stdout="", stderr=""))
    mb.respuesta = True  # Sí → sobrescribir
    app = _app(modulo, documento, repo)

    app.convert()

    assert (repo / "cierre.md").read_text(encoding="utf-8") == "retocado a mano"
    assert mb.tipos() == ["pregunta", "error"]
    assert "OK" not in app.status.get()
    assert sorted(p.name for p in repo.iterdir()) == ["cierre.md"]


@pytest.mark.parametrize("motor_ok", [True, False])
def test_el_temporal_no_pisa_ficheros_ajenos_de_la_carpeta(
        entorno, documento, tmp_path, monkeypatch, motor_ok):
    # El temporal tiene que ser propio de esta ejecución: un fichero que ya
    # estaba en la carpeta con el nombre "esperable" no se borra ni se
    # reemplaza, salga bien o mal la conversión.
    modulo, mb = entorno
    repo = tmp_path / "repo"
    repo.mkdir()
    ajeno = repo / ".convirtiendo-cierre.md"
    ajeno.write_text("de otra conversión", encoding="utf-8")
    vistos = []

    def run(cmd, **kwargs):
        vistos.append(Path(cmd[-1]))
        if not motor_ok:
            raise subprocess.CalledProcessError(1, cmd, output="", stderr="roto")
        Path(cmd[-1]).write_bytes(b"| hola |\n")
        return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="")
    monkeypatch.setattr(modulo.subprocess, "run", run)

    _app(modulo, documento, repo).convert()

    assert ajeno.read_text(encoding="utf-8") == "de otra conversión"
    assert vistos[0] != ajeno and vistos[0].suffix == ".md"
    esperados = {".convirtiendo-cierre.md"} | ({"cierre.md"} if motor_ok else set())
    assert {p.name for p in repo.iterdir()} == esperados  # sin restos propios


def test_sobrescribir_y_motor_que_escribe_a_medias_y_falla_conserva_el_previo(
        entorno, documento, tmp_path, monkeypatch):
    # El motor abre (trunca) su destino, escribe algo y revienta: el .md
    # retocado a mano no puede perderse aunque se eligiera "sobrescribir".
    modulo, mb = entorno
    repo = tmp_path / "repo"
    repo.mkdir()
    (repo / "cierre.md").write_text("retocado a mano", encoding="utf-8")

    def run(cmd, **kwargs):
        Path(cmd[-1]).write_bytes(b"| a me")
        raise subprocess.CalledProcessError(1, cmd, output="", stderr="disco lleno")
    monkeypatch.setattr(modulo.subprocess, "run", run)
    mb.respuesta = True  # Sí → sobrescribir

    _app(modulo, documento, repo).convert()

    assert (repo / "cierre.md").read_text(encoding="utf-8") == "retocado a mano"
    assert mb.tipos() == ["pregunta", "error"]
    assert sorted(p.name for p in repo.iterdir()) == ["cierre.md"]


def test_motor_que_falla_se_muestra_sin_traceback(entorno, documento, tmp_path, monkeypatch):
    modulo, mb = entorno

    def run(cmd, **kwargs):
        raise subprocess.CalledProcessError(1, cmd, output="", stderr="motor_excel falló con cierre.xlsx: zip roto")
    monkeypatch.setattr(modulo.subprocess, "run", run)
    app = _app(modulo, documento, tmp_path / "repo")

    app.convert()

    assert mb.tipos() == ["error"]
    assert "zip roto" in mb.llamadas[0][1][1]
    assert "zip roto" in app.status.get()
    assert not (tmp_path / "repo" / "cierre.md").exists()


def test_motor_que_no_se_puede_lanzar_da_error(entorno, documento, tmp_path, monkeypatch):
    modulo, mb = entorno

    def run(cmd, **kwargs):
        raise PermissionError(13, "Permission denied", cmd[0])
    monkeypatch.setattr(modulo.subprocess, "run", run)
    app = _app(modulo, documento, tmp_path / "repo")

    app.convert()

    assert mb.tipos() == ["error"]
    assert "Permission denied" in app.status.get()


def test_sin_venv_de_motores_avisa_y_no_lanza_nada(entorno, documento, tmp_path, monkeypatch):
    modulo, mb = entorno
    monkeypatch.setattr(modulo, "VENV_TOOLS_PY", tmp_path / "no-existe" / "python3")
    llamadas = []
    monkeypatch.setattr(modulo.subprocess, "run", _motor_que_escribe(llamadas=llamadas))

    _app(modulo, documento, tmp_path / "repo").convert()

    assert llamadas == []
    assert mb.llamadas[0][1][0] == "Falta el venv de motores"


def test_salida_que_no_es_utf8_no_se_pierde(entorno, tmp_path, monkeypatch):
    # MarkItDown (u otro motor) deja un .md en latin-1: la cabecera no se
    # puede añadir, pero el .md generado debe quedar intacto y avisado.
    modulo, mb = entorno
    origen = tmp_path / "pagina.html"
    origen.write_text("<p>hola</p>", encoding="utf-8")
    monkeypatch.setattr(modulo.App, "get_markitdown_cmd", lambda self: "/usr/bin/markitdown")
    bruto = "Año contable".encode("latin-1")
    monkeypatch.setattr(modulo.subprocess, "run", _motor_que_escribe(bruto))
    repo = tmp_path / "repo"

    _app(modulo, origen, repo).convert()

    assert (repo / "pagina.md").read_bytes() == bruto
    assert mb.tipos() == ["info"]
    assert "No se pudo añadir la cabecera" in mb.llamadas[0][1][1]


def test_motor_que_acaba_bien_sin_escribir_no_se_da_por_convertido(entorno, documento, tmp_path, monkeypatch):
    # Código 0 sin fichero de salida (como textutil/soffice ante un formato
    # que no reconocen): decir "OK, Markdown generado en ..." sería mentir.
    modulo, mb = entorno
    monkeypatch.setattr(modulo.subprocess, "run",
                        lambda cmd, **k: subprocess.CompletedProcess(cmd, 0, stdout="", stderr=""))
    app = _app(modulo, documento, tmp_path / "repo")

    app.convert()

    assert not (tmp_path / "repo" / "cierre.md").exists()
    assert mb.tipos() == ["error"]
    assert "OK" not in app.status.get()


def test_validacion_fallida_no_crea_la_carpeta_ni_lanza(entorno, tmp_path, monkeypatch):
    modulo, mb = entorno
    llamadas = []
    monkeypatch.setattr(modulo.subprocess, "run", _motor_que_escribe(llamadas=llamadas))
    repo = tmp_path / "repo"

    _app(modulo, tmp_path / "no-existe.pdf", repo).convert()

    assert llamadas == [] and not repo.exists()
    assert mb.tipos() == ["error"]
