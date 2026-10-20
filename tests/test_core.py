# -*- coding: utf-8 -*-
"""Tests de la lógica pura de ALL2MD (sin Tkinter, sin subprocess).
Solo datos sintéticos -- ningún documento real."""
from pathlib import Path

import pytest

from all2md_core import elegir_motor, nombre_seguro


@pytest.mark.parametrize("nombre,esperado", [
    ("informe", "informe"),
    ("  informe  ", "informe"),
    ("../../otro", "otro"),
    ("/tmp/salida", "salida"),
    ("sub/dir/final", "final"),
])
def test_nombre_seguro_descarta_componentes_de_ruta(nombre, esperado):
    assert nombre_seguro(nombre) == esperado


@pytest.mark.parametrize("nombre", ["", "   ", ".", "..", "/"])
def test_nombre_seguro_rechaza_nombres_vacios_o_solo_ruta(nombre):
    assert nombre_seguro(nombre) is None


@pytest.mark.parametrize("fichero,motor", [
    ("a.xlsx", "motor_excel"),
    ("a.XLSM", "motor_excel"),
    ("a.pdf", "motor_docling"),
    ("a.DOCX", "motor_docling"),
    ("a.pptx", "motor_docling"),
    ("a.xls", "motor_excel"),   # vía LibreOffice → .xlsx
    ("a.doc", "motor_docling"),
    ("a.rtf", "motor_docling"),
    ("a.odt", "motor_docling"),
    ("a.ppt", "motor_docling"),  # vía LibreOffice → .pptx
    ("sin_extension", None),
])
def test_elegir_motor_enruta_por_extension(fichero, motor):
    assert elegir_motor(Path(fichero)) == motor


# --- Loop 2: un único constructor de comandos para convertir y para copiar ---
import shlex
import unicodedata

from all2md_core import comando_shell, construir_comando

PY = Path("/opt/repo/.venv/bin/python3")
REPO = Path("/opt/repo")


def test_construir_comando_motor_propio_usa_python_del_venv():
    cmd = construir_comando(Path("/d/a.xlsx"), Path("/r/a.md"), python_motores=PY,
                            repo_dir=REPO, markitdown="/x/markitdown")
    assert cmd == [str(PY), "/opt/repo/motor_excel.py", "/d/a.xlsx", "/r/a.md"]


def test_construir_comando_markitdown_por_defecto():
    cmd = construir_comando(Path("/d/a.html"), Path("/r/a.md"), python_motores=PY,
                            repo_dir=REPO, markitdown="/x/markitdown")
    assert cmd == ["/x/markitdown", "/d/a.html", "-o", "/r/a.md"]


def test_comando_shell_no_deja_tilde_entrecomillada():
    # shlex.quote('~/x') -> "'~/x'": la shell NO expande la tilde dentro de
    # comillas, así que el comando copiado no funcionaría.
    cmd = construir_comando(Path("~/d/a.html"), Path("/r/a.md"), python_motores=PY,
                            repo_dir=REPO, markitdown="~/env/bin/markitdown")
    texto = comando_shell(cmd)
    assert "~" not in texto
    assert shlex.split(texto)[0] == str(Path.home() / "env/bin/markitdown")


def test_comando_shell_escapa_nombres_hostiles():
    cmd = construir_comando(Path('/d/informe"$(rm -rf x)".pdf'), Path("/r/a.md"),
                            python_motores=PY, repo_dir=REPO, markitdown="m")
    assert shlex.split(comando_shell(cmd))[2] == '/d/informe"$(rm -rf x)".pdf'


def test_construir_comando_no_revienta_con_tilde_de_usuario_inexistente():
    cmd = construir_comando(Path("~usuario_que_no_existe_zz/a.html"), Path("/r/a.md"),
                            python_motores=PY, repo_dir=REPO, markitdown="m")
    assert cmd[1] == "~usuario_que_no_existe_zz/a.html"


# --- Loop 6: frontmatter de procedencia para RAG ---
import datetime as dt

from all2md_core import anadir_frontmatter, con_frontmatter

FECHA = dt.date(2026, 10, 4)


def test_con_frontmatter_cabecera_yaml_con_procedencia():
    md = con_frontmatter("# Hola\n", Path("/Users/x/docs/cierre.xlsx"), "motor_excel", FECHA)
    assert md == (
        "---\n"
        'fuente: "cierre.xlsx"\n'
        'motor: "motor_excel"\n'
        "convertido: 2026-10-04\n"
        "---\n\n"
        "# Hola\n"
    )


def test_con_frontmatter_no_filtra_la_ruta_completa():
    md = con_frontmatter("x", Path("/Users/secreto/carpeta/a.pdf"), "MarkItDown", FECHA)
    assert "/Users" not in md and "secreto" not in md


def test_con_frontmatter_escapa_nombres_hostiles_para_yaml():
    import json
    nombre = 'raro: "x"\n---\ny.pdf'
    md = con_frontmatter("x", Path(nombre), "MarkItDown", FECHA)
    cabecera = md.split("\n---\n")[0].splitlines()
    linea = next(l for l in cabecera if l.startswith("fuente: "))
    assert json.loads(linea[len("fuente: "):]) == nombre


def test_con_frontmatter_sustituye_uno_previo_en_vez_de_apilar():
    previo = con_frontmatter("cuerpo\n", Path("a.pdf"), "MarkItDown", FECHA)
    md = con_frontmatter(previo, Path("a.pdf"), "motor_docling", FECHA)
    assert md.count("---\n") == 2
    assert 'motor: "motor_docling"' in md
    assert md.endswith("cuerpo\n")


def test_anadir_frontmatter_reescribe_el_fichero(tmp_path):
    destino = tmp_path / "a.md"
    destino.write_text("| a | b |\n", encoding="utf-8")
    anadir_frontmatter(destino, Path("a.xlsx"), "motor_excel", FECHA)
    texto = destino.read_text(encoding="utf-8")
    assert texto.startswith("---\nfuente: \"a.xlsx\"")
    assert texto.endswith("| a | b |\n")


def test_con_frontmatter_no_toca_un_frontmatter_ajeno_que_empieza_por_fuente():
    ajeno = "---\nfuente: Ministerio\n\nTexto importante\n\n---\n\nResto\n"
    md = con_frontmatter(ajeno, Path("a.md"), "MarkItDown", FECHA)
    assert md.endswith(ajeno)


@pytest.mark.parametrize("nombre", [
    'raro: "x" #y.pdf', "x\x7fy.pdf", "x\x85y.pdf", "x\x9fy.pdf", "x￾y.pdf",
    "tab\there.pdf", "emoji_🏨.pdf",
])
def test_con_frontmatter_es_yaml_valido_y_conserva_el_nombre(nombre):
    yaml = pytest.importorskip("yaml")
    md = con_frontmatter("x", Path(nombre), "MarkItDown", FECHA)
    cabecera = md.split("\n---\n", 1)[0][len("---\n"):]
    assert yaml.safe_load(cabecera)["fuente"] == nombre


def test_anadir_frontmatter_no_deja_el_fichero_a_medias_si_falla(tmp_path, monkeypatch):
    destino = tmp_path / "a.md"
    destino.write_text("contenido original\n", encoding="utf-8")

    escribir_original = Path.write_text

    def falla(self, *a, **k):
        escribir_original(self, "", encoding="utf-8")  # trunca, como open("w")
        raise OSError("disco lleno")
    monkeypatch.setattr(Path, "write_text", falla)
    with pytest.raises(OSError):
        anadir_frontmatter(destino, Path("a.pdf"), "MarkItDown", FECHA)
    monkeypatch.undo()
    assert destino.read_text(encoding="utf-8") == "contenido original\n"


# --- Loop 7: no pisar un .md existente sin preguntar ---
from all2md_core import ruta_libre


def test_ruta_libre_devuelve_el_nombre_si_no_existe(tmp_path):
    assert ruta_libre(tmp_path, "informe") == tmp_path / "informe.md"


def test_ruta_libre_numera_si_ya_existe(tmp_path):
    (tmp_path / "informe.md").write_text("x")
    (tmp_path / "informe-2.md").write_text("x")
    assert ruta_libre(tmp_path, "informe") == tmp_path / "informe-3.md"


# --- Loop 8: escritura atómica compartida por motores y frontmatter ---
from all2md_core import escribir_atomico


def test_escribir_atomico_escribe_y_no_deja_temporales(tmp_path):
    destino = tmp_path / "a.md"
    escribir_atomico(destino, "hola\n")
    assert destino.read_text(encoding="utf-8") == "hola\n"
    assert [f.name for f in tmp_path.iterdir()] == ["a.md"]


def test_escribir_atomico_conserva_el_anterior_si_falla_la_codificacion(tmp_path):
    destino = tmp_path / "a.md"
    destino.write_text("anterior\n", encoding="utf-8")
    with pytest.raises(UnicodeError):
        escribir_atomico(destino, "surrogate suelto \udc80")
    assert destino.read_text(encoding="utf-8") == "anterior\n"
    assert [f.name for f in tmp_path.iterdir()] == ["a.md"]


# --- Automejora 2, loop 8: el usuario teclea la extensión ---
@pytest.mark.parametrize("nombre,esperado", [
    ("informe.md", "informe"),
    ("informe.MD", "informe"),
    ("informe.md  ", "informe"),
    ("cierre.v2", "cierre.v2"),        # otros puntos se respetan
    ("informe.md.md", "informe.md"),   # solo una extensión, la final
])
def test_nombre_seguro_quita_la_extension_md_tecleada(nombre, esperado):
    # La app añade ".md": "informe.md" acababa en "informe.md.md".
    assert nombre_seguro(nombre) == esperado


@pytest.mark.parametrize("nombre", [".md", "  .MD "])
def test_nombre_seguro_solo_extension_no_es_un_nombre(nombre):
    assert nombre_seguro(nombre) is None


# --- Automejora 2, loop 9: rutas de Windows pegadas y traversal con .md ---
@pytest.mark.parametrize("nombre,esperado", [
    ("C:\\docs\\informe", "informe"),
    ("..\\..\\otro", "otro"),
    ("carpeta\\sub/final.md", "final"),
])
def test_nombre_seguro_entiende_la_barra_invertida_como_separador(nombre, esperado):
    # En macOS "\" es un carácter válido: sin esto el .md se llamaba
    # literalmente "C:\docs\informe.md".
    assert nombre_seguro(nombre) == esperado


@pytest.mark.parametrize("nombre", ["..md", "...md", "x/..md", "../.md", "..\\", "a\\..", "..\\.md"])
def test_nombre_seguro_traversal_con_md_o_barra_invertida(nombre):
    assert nombre_seguro(nombre) is None


# --- Automejora 2, loop 10: caracteres de control en el nombre ---
@pytest.mark.parametrize("nombre,esperado", [
    ("info\x00rme", "informe"),        # NUL: open() lanza ValueError
    ("a\tb", "ab"),
    ("lin1\nlin2", "lin1lin2"),        # pegado desde una celda/correo
    ("x\x7fy\x85z", "xyz"),
    ("emoji_🏨", "emoji_🏨"),           # lo imprimible se respeta
])
def test_nombre_seguro_quita_caracteres_de_control(nombre, esperado):
    assert nombre_seguro(nombre) == esperado


def test_nombre_seguro_solo_control_no_es_un_nombre():
    assert nombre_seguro("\x00\x01") is None


def test_un_nombre_saneado_se_puede_escribir(tmp_path):
    nombre = nombre_seguro("info\x00rme\n")
    escribir_atomico(tmp_path / f"{nombre}.md", "x")
    assert (tmp_path / "informe.md").exists()


# --- Automejora 2, loop 11: controles bidi y separadores de línea Unicode ---
@pytest.mark.parametrize("nombre,esperado", [
    ("informe‮fdp.exe", "informefdp.exe"),   # RLO: en Finder se vería "informeexe.pdf"
    ("a‪b‬c", "abc"),
    ("a⁦b⁩c", "abc"),
    ("a‏b", "ab"),                           # RLM
    ("lin1 lin2 ", "lin1lin2"),
    ("familia_👨‍👩‍👧", "familia_👨‍👩‍👧"),  # ZWJ se respeta
])
def test_nombre_seguro_quita_bidi_y_separadores_de_linea(nombre, esperado):
    assert nombre_seguro(nombre) == esperado


@pytest.mark.parametrize("nombre", ["..\x00/", ".‮.", " .. "])
def test_nombre_seguro_controles_intercalados_no_abren_traversal(nombre):
    assert nombre_seguro(nombre) is None


# --- Automejora 2, loop 12: nombres que empiezan por punto ---
@pytest.mark.parametrize("nombre,esperado", [
    (".informe", "informe"),
    ("...informe", "informe"),
    (". informe", "informe"),
    ("a.b", "a.b"),
])
def test_nombre_seguro_no_genera_ficheros_ocultos(nombre, esperado):
    # ".informe.md" queda oculto en Finder y muchos indexadores RAG se
    # saltan los ficheros punto: la conversión "desaparecía".
    assert nombre_seguro(nombre) == esperado


def test_ruta_libre_con_nombre_saneado_no_es_oculta(tmp_path):
    assert not ruta_libre(tmp_path, nombre_seguro(".x")).name.startswith(".")


# --- Automejora 2, loop 13: nombres demasiado largos ---
from all2md_core import MAX_BYTES_NOMBRE


@pytest.mark.parametrize("nombre", ["a" * 300, "ñ" * 200, "🏨" * 100, "x" * 250 + "🏨"])
def test_nombre_seguro_cabe_en_el_limite_del_sistema(nombre, tmp_path):
    # APFS/ext4: 255 bytes por componente. Con ".md", "-NN" de ruta_libre
    # y el ".tmp" de escribir_atomico tiene que seguir cabiendo.
    limpio = nombre_seguro(nombre)
    assert limpio and nombre.startswith(limpio)
    assert len(limpio.encode("utf-8")) <= MAX_BYTES_NOMBRE
    destino = ruta_libre(tmp_path, limpio)
    escribir_atomico(destino, "x")
    escribir_atomico(ruta_libre(tmp_path, limpio), "x")   # la copia "-2"
    assert len(list(tmp_path.iterdir())) == 2


def test_nombre_seguro_corto_no_se_toca():
    assert nombre_seguro("informe 2026") == "informe 2026"


# --- Automejora 2, loop 14: un único nombre de salida para convertir y copiar ---
from all2md_core import nombre_salida


@pytest.mark.parametrize("tecleado,documento,esperado", [
    ("informe", "/d/cierre.xlsx", "informe"),
    ("", "/d/cierre.xlsx", "cierre"),
    ("   ", "/d/cierre.xlsx", "cierre"),
    ("", "/d/.oculto.pdf", "oculto"),          # el stem también se sanea
    ("", "/d/a‮fdp.exe.pdf", "afdp.exe"),
    ("..", "/d/cierre.xlsx", None),            # tecleado inválido: no se adivina
    ("", "/", None),                           # sin documento del que tomarlo
])
def test_nombre_salida(tecleado, documento, esperado):
    assert nombre_salida(tecleado, Path(documento)) == esperado


# --- Automejora 2, loop 15: ficheros de bloqueo de Office ---
from all2md_core import es_bloqueo_office


@pytest.mark.parametrize("fichero,esperado", [
    ("/d/~$cierre.xlsx", True),       # Excel abierto deja este "fichero" de 165 bytes
    ("/d/~$informe.docx", True),
    ("/d/~$presenta.pptx", True),
    ("/d/cierre~$.xlsx", False),
    ("/d/~cierre.xlsx", False),
    ("/d/cierre.xlsx", False),
])
def test_es_bloqueo_office(fichero, esperado):
    assert es_bloqueo_office(Path(fichero)) is esperado


# --- Automejora 2, loop 16: plantillas de Excel ---
@pytest.mark.parametrize("fichero", ["a.xltx", "a.XLTM"])
def test_plantillas_de_excel_van_al_motor_excel(fichero):
    # Una plantilla de cierre (.xltx) tiene la misma estructura que el
    # .xlsx: openpyxl la lee igual; MarkItDown no la reconoce.
    assert elegir_motor(Path(fichero)) == "motor_excel"


# --- Automejora 2, loop 17: el diálogo ofrece todo lo que se enruta ---
from all2md_core import (EXTENSIONES_DOCLING, EXTENSIONES_EXCEL,
                         EXTENSIONES_TEXTUTIL, tipos_dialogo)


def test_tipos_dialogo_incluye_todas_las_extensiones_enrutadas():
    tipos = tipos_dialogo()
    soportados = dict(tipos)["Documentos soportados"].split()
    for ext in EXTENSIONES_EXCEL | EXTENSIONES_DOCLING | EXTENSIONES_TEXTUTIL:
        assert f"*{ext}" in soportados, ext
    assert tipos[-1] == ("Todos los ficheros", "*.*")
    # Sin LibreOffice nada convierte .ppt ni .xls: no se ofrecen.
    for ext in (".ppt", ".xls"):
        assert all(f"*{ext}" not in p.split() for _, p in tipos), ext


def test_tipos_dialogo_con_libreoffice_ofrece_xls_y_ppt():
    tipos = dict(tipos_dialogo(hay_soffice=True))
    assert "*.xls" in tipos["Documentos soportados"].split()
    assert "*.xls" in tipos["Excel"].split()
    assert "*.ppt" in tipos["PowerPoint"].split()


def test_tipos_dialogo_cada_grupo_es_subconjunto_de_soportados():
    tipos = dict(tipos_dialogo())
    soportados = set(tipos.pop("Documentos soportados").split())
    tipos.pop("Todos los ficheros")
    for grupo, patrones in tipos.items():
        assert set(patrones.split()) <= soportados, grupo


def test_la_app_toma_los_tipos_del_dialogo_de_core():
    fuente = (Path(__file__).resolve().parent.parent / "all2md_repo_app.py").read_text(encoding="utf-8")
    assert "filetypes=all2md_core.tipos_dialogo(self.hay_soffice)" in fuente


# --- Automejora 2, loop 18: nunca escribir encima del documento de origen ---
from all2md_core import misma_ruta


def test_misma_ruta_detecta_alias(tmp_path):
    doc = tmp_path / "a.md"
    doc.write_text("original")
    assert misma_ruta(doc, doc)
    assert misma_ruta(doc, tmp_path / "sub" / ".." / "a.md")
    enlace = tmp_path / "enlace.md"
    enlace.symlink_to(doc)
    assert misma_ruta(doc, enlace)
    assert not misma_ruta(doc, tmp_path / "b.md")


# --- Automejora 2, loop 19: BOM al principio del .md del motor ---
def test_con_frontmatter_quita_el_bom_inicial():
    md = con_frontmatter("﻿# Hola\n", Path("a.pdf"), "MarkItDown", FECHA)
    assert "﻿" not in md
    assert md.endswith("---\n\n# Hola\n")


def test_con_frontmatter_no_apila_si_el_previo_traia_bom():
    previo = "﻿" + con_frontmatter("cuerpo\n", Path("a.pdf"), "MarkItDown", FECHA)
    md = con_frontmatter(previo, Path("a.pdf"), "motor_docling", FECHA)
    assert md.count("fuente:") == 1 and md.endswith("cuerpo\n")


def test_con_frontmatter_conserva_un_bom_que_no_esta_al_principio():
    md = con_frontmatter("a﻿b", Path("a.pdf"), "MarkItDown", FECHA)
    assert md.endswith("a﻿b")


def test_con_frontmatter_varios_bom_seguidos():
    previo = "\ufeff\ufeff" + con_frontmatter("cuerpo\n", Path("a.pdf"), "MarkItDown", FECHA)
    md = con_frontmatter(previo, Path("a.pdf"), "motor_docling", FECHA)
    assert "\ufeff" not in md
    assert md.count("fuente:") == 1 and md.endswith("cuerpo\n")


# --- Ronda 3, loop A: .doc/.rtf/.odt vía textutil; .ppt/.xls con aviso ---
from all2md_core import motivo_no_soportado


@pytest.mark.parametrize("fichero,pista", [
    ("viejo.ppt", ".pptx"), ("VIEJO.PPT", ".pptx"), ("viejo.xls", ".xlsx"),
])
def test_motivo_no_soportado_explica_como_resolverlo(fichero, pista):
    motivo = motivo_no_soportado(Path(fichero), hay_soffice=False)
    assert motivo and pista in motivo and "LibreOffice" in motivo


@pytest.mark.parametrize("fichero", ["viejo.ppt", "viejo.xls"])
def test_con_libreoffice_ppt_y_xls_si_se_convierten(fichero):
    assert motivo_no_soportado(Path(fichero), hay_soffice=True) is None


@pytest.mark.parametrize("fichero", ["a.pdf", "a.doc", "a.xlsx", "a.html", "a"])
def test_motivo_no_soportado_none_para_lo_que_si_se_convierte(fichero):
    assert motivo_no_soportado(Path(fichero), hay_soffice=False) is None


# --- Ronda 3, loop B: una sola validación para Convertir y Copiar CLI ---
from all2md_core import validar_conversion


def _doc(tmp_path, nombre="cierre.xlsx"):
    ruta = tmp_path / nombre
    ruta.write_bytes(b"sintetico")
    return ruta


def test_validar_conversion_ok_da_el_destino(tmp_path):
    v = validar_conversion(_doc(tmp_path), "", tmp_path / "repo", hay_soffice=False)
    assert v.error is None
    assert v.destino == tmp_path / "repo" / "cierre.md"


@pytest.mark.parametrize("caso,titulo", [
    ("no_existe", "Error"),
    ("bloqueo", "No es un documento"),
    ("no_soportado", "Formato no soportado"),
    ("nombre", "Nombre inválido"),
    ("mismo", "Destino igual al origen"),
])
def test_validar_conversion_rechaza(tmp_path, caso, titulo):
    repo = tmp_path / "repo"
    repo.mkdir()
    tecleado = ""
    if caso == "no_existe":
        origen = tmp_path / "fantasma.pdf"
    elif caso == "bloqueo":
        origen = _doc(tmp_path, "~$cierre.xlsx")
    elif caso == "no_soportado":
        origen = _doc(tmp_path, "viejo.ppt")
    elif caso == "nombre":
        origen, tecleado = _doc(tmp_path), ".."
    else:  # el .md caería encima del propio origen
        origen = _doc(repo, "notas.md")
    v = validar_conversion(origen, tecleado, repo, hay_soffice=False)
    assert v.destino is None
    assert v.error[0] == titulo and v.error[1]


def test_validar_conversion_sin_comprobar_existencia_para_la_plantilla_de_copy_cli(tmp_path):
    v = validar_conversion(Path("/ruta/a/tu.pdf"), "", tmp_path, hay_soffice=False,
                           comprobar_existe=False)
    assert v.destino == tmp_path / "tu.md"


def test_misma_ruta_mayusculas_y_enlace_duro(tmp_path):
    doc = tmp_path / "Notas.md"
    doc.write_text("x")
    import os
    os.link(doc, tmp_path / "duro.md")
    assert misma_ruta(doc, tmp_path / "duro.md")
    sensible = not (tmp_path / "notas.md").exists()  # APFS por defecto no distingue
    assert misma_ruta(doc, tmp_path / "notas.md") is (not sensible)


def test_convertir_y_copiar_usan_la_misma_validacion():
    # Las reglas viven (y se prueban) en validar_conversion; aquí solo se
    # fija que ninguno de los dos botones se la salta.
    import ast
    fuente = (Path(__file__).resolve().parent.parent / "all2md_repo_app.py").read_text(encoding="utf-8")
    metodos = {n.name: ast.unparse(n) for n in ast.walk(ast.parse(fuente))
               if isinstance(n, ast.FunctionDef)}
    for metodo in ("convert", "copy_cli"):
        assert "validar_conversion(" in metodos[metodo], metodo
        assert "misma_ruta(" not in metodos[metodo] and "es_bloqueo_office(" not in metodos[metodo], metodo
        assert ".stem" not in metodos[metodo], metodo  # el nombre lo decide core


@pytest.mark.parametrize("ruta", ["", "."])
def test_validar_conversion_campo_vacio_o_carpeta_no_es_un_documento(tmp_path, ruta, monkeypatch):
    # Path("") es "." y existe (es una carpeta): antes pasaba el control de
    # existencia y el error acababa siendo un engañoso "Nombre inválido".
    monkeypatch.chdir(tmp_path)
    v = validar_conversion(Path(ruta), "", tmp_path / "repo", hay_soffice=False)
    assert v.error and v.error[0] == "Error"


def test_validar_conversion_recorta_espacios_de_la_ruta_pegada(tmp_path):
    doc = _doc(tmp_path)
    v = validar_conversion(f"  {doc}  ", "", f" {tmp_path / 'repo'} ", hay_soffice=False)
    assert v.error is None and v.destino == tmp_path / "repo" / "cierre.md"


def test_validar_conversion_tilde_de_usuario_inexistente_no_revienta(tmp_path):
    v = validar_conversion("~usuario_que_no_existe_zz/x.pdf", "", tmp_path, hay_soffice=False)
    assert v.error and v.error[0] == "Error"


def test_validar_conversion_repo_que_es_un_fichero(tmp_path):
    fichero = _doc(tmp_path, "no_soy_carpeta")
    v = validar_conversion(_doc(tmp_path), "", fichero, hay_soffice=False)
    assert v.error and v.error[0] == "Carpeta repo inválida"


def test_validar_conversion_paquete_o_carpeta_lo_dice(tmp_path):
    paquete = tmp_path / "informe.pages"   # en macOS, un paquete es una carpeta
    paquete.mkdir()
    v = validar_conversion(paquete, "", tmp_path / "repo", hay_soffice=False)
    assert v.error and "carpeta" in v.error[1]


def test_validar_conversion_devuelve_el_origen_normalizado(tmp_path):
    doc = _doc(tmp_path)
    v = validar_conversion(f" {doc} ", "", tmp_path, hay_soffice=False)
    assert v.origen == doc


def test_validar_conversion_respeta_un_espacio_final_legitimo(tmp_path):
    doc = _doc(tmp_path, "informe.pdf ")   # macOS admite el espacio final
    v = validar_conversion(str(doc), "", tmp_path / "repo", hay_soffice=False)
    assert v.error is None and v.origen == doc


# --- Ronda 3, loop D: nombre_seguro con invisibles, NFD y recorte limpio ---


@pytest.mark.parametrize("nombre", ["​", "⁠", "﻿", "\xad", "​ ​", "​.md"])
def test_nombre_seguro_rechaza_nombres_solo_invisibles(nombre):
    # En Finder se verían como ".md": un fichero sin nombre visible.
    assert nombre_seguro(nombre) is None


def test_nombre_seguro_invisible_inicial_no_esconde_un_punto():
    # "​.informe" se ve como ".informe.md" (parece oculto).
    assert nombre_seguro("​.informe") == "informe"


def test_nombre_seguro_normaliza_a_nfc():
    nfd = unicodedata.normalize("NFD", "acción")
    assert nombre_seguro(nfd) == unicodedata.normalize("NFC", "acción")


def test_nombre_seguro_recorte_no_separa_letra_y_tilde():
    nfd = unicodedata.normalize("NFD", "é" * 150)
    limpio = nombre_seguro(nfd)
    assert limpio == "é" * (MAX_BYTES_NOMBRE // 2)


@pytest.mark.parametrize("nombre", [
    "a" * 197 + ".md" + "zzz",   # el recorte deja ".md" al final
    "a" * 199 + " b",            # el recorte deja un espacio al final
])
def test_nombre_seguro_recorte_no_deja_md_ni_espacio_final(nombre):
    limpio = nombre_seguro(nombre)
    assert not limpio.lower().endswith(".md") and not limpio.endswith(" ")
    assert len(limpio.encode("utf-8")) <= MAX_BYTES_NOMBRE


@pytest.mark.parametrize("nombre,esperado", [
    ("informe /", "informe"),
    ("informe.md.", "informe"),
    ("informe. ", "informe"),
])
def test_nombre_seguro_sin_basura_final(nombre, esperado):
    assert nombre_seguro(nombre) == esperado


@pytest.mark.parametrize("nombre,esperado", [
    ("a\n\u0301b", "\u00e1b"),               # NFC tras quitar el control
    ("informe\u00a0.md", "informe"),          # NBSP antes de .md
    ("\u200b\u3000.informe", "informe"),      # espacio ideográfico inicial
    ("informe.md\u200b", "informe"),           # Cf final sin recorte
    ("informe\u200b.md", "informe"),
    ("\u3164.informe", "informe"),             # relleno Hangul (se ve en blanco)
    ("\u0301.oculto", "oculto"),               # tilde suelta inicial
])
def test_nombre_seguro_invisibles_unicode(nombre, esperado):
    assert nombre_seguro(nombre) == esperado


@pytest.mark.parametrize("nombre", ["\u3164", "\u2800", "\u0301", "\u00a0\u3000"])
def test_nombre_seguro_rechaza_rellenos_que_se_ven_en_blanco(nombre):
    assert nombre_seguro(nombre) is None


@pytest.mark.parametrize("nombre", [
    "a\n\u0301b", "informe\u00a0.md", "\u200b\u3000.informe", "x\u202e\u00a0\\",
    "\u3164.informe", "infórme  ", "a" * 199 + "\u00e9",
])
def test_nombre_seguro_es_idempotente(nombre):
    una = nombre_seguro(nombre)
    assert una is None or nombre_seguro(una) == una


@pytest.mark.parametrize("nombre", ["❤️", "Informe 1️⃣", "Café́"])
def test_nombre_seguro_conserva_marcas_unidas_al_final(nombre):
    # Selector de variación (U+FE0F), keycap (U+20E3): al final van unidos
    # a su carácter. Quitarlos convertía "❤️" en "❤" (corazón en modo texto).
    assert nombre_seguro(nombre) == unicodedata.normalize("NFC", nombre)
