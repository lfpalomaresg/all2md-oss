"""
Lógica pura de ALL2MD -- sin Tkinter ni subprocess, para poder probarla
con pytest en cualquier Python (el de Homebrew puede venir sin _tkinter).
all2md_repo_app.py solo pinta la ventana y delega aquí las decisiones.
"""
from __future__ import annotations

import datetime as dt
import os
import re
import shlex
import unicodedata
from dataclasses import dataclass
from pathlib import Path

# Enrutamiento por tipo de documento: un Excel nativo trae la estructura de
# celdas combinadas accesible directamente (motor_excel, sin IA, exacto);
# PDF/Word/PowerPoint necesitan que algo reconstruya la tabla visualmente
# (motor_docling, gratis, local); el resto sigue por MarkItDown, que ya
# cubre bien lo que no tiene tablas complejas que reconstruir.
# Las plantillas (.xltx/.xltm) son el mismo formato OOXML: openpyxl las lee
# igual y MarkItDown no las reconoce.
EXTENSIONES_EXCEL = {".xlsx", ".xlsm", ".xltx", ".xltm"}
EXTENSIONES_DOCLING = {".pdf", ".docx", ".pptx"}
# Word antiguo, RTF y OpenDocument: Docling no los lee y MarkItDown
# tampoco (solo .docx). motor_docling los pasa antes a HTML con textutil,
# que viene con macOS (gratis, sin instalar nada).
EXTENSIONES_TEXTUTIL = {".doc", ".rtf", ".odt"}

# Office 97-2003 binario: MarkItDown no tiene conversor para .ppt y para
# .xls necesita xlrd, que no está instalado -- antes el diálogo los
# ofrecía y la conversión fallaba (hallazgo de careo, MEDIA, 04/10). Con
# LibreOffice instalado (gratis), preconversion.py los pasa a .xlsx (para
# motor_excel: conserva combinadas) y .pptx (para Docling).
EXTENSIONES_SOFFICE = {".xls": "motor_excel", ".ppt": "motor_docling"}

_SIN_LIBREOFFICE = {
    ".ppt": "PowerPoint 97-2003 (.ppt) necesita LibreOffice para convertirse, "
            "y no está instalado. Instálalo (gratis, y reabre esta app) o abre el fichero en "
            "PowerPoint o Keynote y guárdalo como .pptx.",
    ".xls": "Excel 97-2003 (.xls) necesita LibreOffice para convertirse, y "
            "no está instalado. Instálalo (gratis, y reabre esta app) o abre el fichero en Excel "
            "o Numbers y guárdalo como .xlsx.",
}


def motivo_no_soportado(origen: Path, hay_soffice: bool) -> str | None:
    """Mensaje para el usuario si el formato no se puede convertir con lo
    instalado en el equipo, o None."""
    if hay_soffice:
        return None
    return _SIN_LIBREOFFICE.get(Path(origen).suffix.lower())


def tipos_dialogo(hay_soffice: bool = False) -> list[tuple[str, str]]:
    """filetypes del diálogo "Elegir documento", derivado de las mismas
    extensiones que enrutan: a mano se había quedado sin .xlsm (y sin las
    plantillas), que solo se podían elegir con "Todos los ficheros".
    .xls/.ppt solo si hay LibreOffice para convertirlos."""
    def patrones(*exts):
        return " ".join(f"*{e}" for e in sorted(exts))

    excel = set(EXTENSIONES_EXCEL)
    powerpoint = {".pptx"}
    if hay_soffice:
        excel.add(".xls")
        powerpoint.add(".ppt")
    word = {".docx"} | EXTENSIONES_TEXTUTIL
    return [
        ("Documentos soportados", patrones(".pdf", *word, *excel, *powerpoint)),
        ("PDF", patrones(".pdf")),
        ("Word", patrones(*word)),
        ("Excel", patrones(*excel)),
        ("PowerPoint", patrones(*powerpoint)),
        ("Todos los ficheros", "*.*"),
    ]


def elegir_motor(origen: Path) -> str | None:
    """Nombre del motor propio que toca según la extensión, o None si va
    por el motor por defecto (MarkItDown)."""
    ext = origen.suffix.lower()
    if ext in EXTENSIONES_SOFFICE:
        return EXTENSIONES_SOFFICE[ext]
    if ext in EXTENSIONES_EXCEL:
        return "motor_excel"
    if ext in EXTENSIONES_DOCLING or ext in EXTENSIONES_TEXTUTIL:
        return "motor_docling"
    return None


# Marcas y controles de dirección (LRM/RLM, embeddings, overrides,
# isolates): un RLO hace que "informe\u202efdp.exe" se VEA como
# "informeexe.pdf" en Finder. El resto de caracteres de formato (Cf) se
# respeta: el ZWJ forma parte de emojis compuestos.
_BIDI = set("\u200e\u200f\u061c") | {chr(c) for c in range(0x202A, 0x202F)} \
    | {chr(c) for c in range(0x2066, 0x206A)}


def es_bloqueo_office(origen: Path) -> bool:
    """¿Es el fichero de bloqueo "~$nombre.xlsx" que Word/Excel/PowerPoint
    dejan junto al documento mientras está abierto? Pesa ~165 bytes, no es
    un documento: se enrutaba al motor y fallaba con un error críptico."""
    return Path(origen).name.startswith("~$")


# APFS/ext4 admiten 255 bytes por componente de ruta. El nombre tiene que
# caber con ".md", el "-N" de ruta_libre y el ".<nombre>.tmp" de
# escribir_atomico: 200 deja margen de sobra. Sin tope, un nombre largo
# (o con muchos emojis, 4 bytes cada uno) fallaba con "File name too long".
MAX_BYTES_NOMBRE = 200


# Caracteres que no son Cf ni espacio pero se ven en blanco: rellenos de
# Hangul y la celda braille vacía.
_RELLENOS = set("\u3164\u115f\u1160\uffa0\u2800")


def _invisible(ch: str) -> bool:
    """¿Se ve en blanco? Espacios Unicode (NBSP, U+3000...), formato (Cf:
    ZWSP, BOM, guion blando...), rellenos, o una marca combinante suelta."""
    return (ch.isspace() or ch in _RELLENOS
            or unicodedata.category(ch) in ("Cf", "Mn", "Me"))


def _limpia_extremos(limpio: str, quitar_md: bool) -> str:
    """Quita puntos e invisibles de los dos extremos y, si `quitar_md`, una
    extensión ".md" final (la app la añade). En bucle: "informe\u00a0.md"
    deja un NBSP final al quitar ".md", y "\u200b\u3000.informe" esconde un
    punto inicial detrás de dos invisibles (hallazgos de careo, BAJA, 04/10)."""
    def recorta_final(t):
        # Sin tocar marcas combinantes: al final van unidas a su carácter
        # (U+FE0F de "❤️", el keycap de "1️⃣"); solo se quitan al principio,
        # donde van sueltas.
        while t and (t[-1] == "." or (_invisible(t[-1])
                                      and unicodedata.category(t[-1]) not in ("Mn", "Me"))):
            t = t[:-1]
        return t

    # El final antes que el principio: ".md" a secas debe quedar en "", no
    # perder su punto inicial y convertirse en un fichero llamado "md".
    limpio = recorta_final(limpio)
    if quitar_md and limpio.lower().endswith(".md"):
        limpio = recorta_final(limpio[:-3])
    while limpio and (limpio[0] == "." or _invisible(limpio[0])):
        limpio = limpio[1:]
    return limpio


def nombre_seguro(nombre: str) -> str | None:
    """Nombre de fichero sin componentes de ruta, o None si no queda nada
    utilizable. .name descarta separadores, ".." intermedios y rutas
    absolutas pegadas por error -- sin esto, "../../otro" escribiría fuera
    del repo (hallazgo de careo, ALTA, 23/08). Path("..").name es ".." --
    se rechaza explícitamente."""
    # Fuera los caracteres de control (Cc), separadores de línea Unicode
    # (Zl/Zp) y controles bidi: un NUL hacía que open() lanzara
    # ValueError, y un salto pegado desde una celda o un correo acababa
    # dentro del nombre del fichero.
    nombre = "".join(ch for ch in nombre
                     if unicodedata.category(ch) not in ("Cc", "Zl", "Zp")
                     and ch not in _BIDI)
    # NFC después de filtrar (un control entre letra y tilde impedía
    # componerlas): un nombre pegado desde Finder llega en NFD y el recorte
    # por bytes podía separar la letra de su tilde.
    nombre = unicodedata.normalize("NFC", nombre)
    # "\" también separa: una ruta de Windows pegada en el nombre
    # ("C:\docs\informe") se quedaba entera, con barras, en macOS.
    limpio = Path(nombre.strip().replace("\\", "/")).name
    # Sin puntos ni invisibles en los extremos: ".informe.md" queda oculto
    # en Finder ("\u200b.informe" lo parece) y muchos indexadores RAG se
    # saltan los ficheros punto; "." y ".." quedan en "". La app añade
    # ".md": si el usuario ya la tecleó, "informe.md" acababa en
    # "informe.md.md" -- se quita UNA, sin distinguir mayúsculas.
    limpio = _limpia_extremos(limpio, quitar_md=True)
    # Recorte por BYTES UTF-8 (no por caracteres) sin partir un carácter.
    cortado = limpio.encode("utf-8")[:MAX_BYTES_NOMBRE].decode("utf-8", "ignore")
    if cortado != limpio:
        # El corte puede dejar ".md" (acabaría en ".md.md"), un espacio,
        # una tilde o un ZWJ sueltos al final.
        while True:
            antes = cortado
            cortado = _limpia_extremos(cortado, quitar_md=True)
            if cortado == antes:
                break
    limpio = cortado
    if not limpio or limpio in (".", ".."):
        return None
    return limpio


def nombre_salida(tecleado: str, documento: Path) -> str | None:
    """Nombre del .md (sin extensión): el tecleado, o el del documento si
    no se tecleó nada -- saneado en ambos casos. Lo usan convert() y
    copy_cli() para que no calculen el nombre cada uno a su manera (copy_cli
    caía al stem del documento sin sanear: ".x.pdf" daba un ".x.md"
    oculto). Un nombre tecleado inválido da None, no se sustituye en
    silencio por el del documento."""
    if tecleado.strip():
        return nombre_seguro(tecleado)
    return nombre_seguro(Path(documento).stem)


def construir_comando(origen: Path, destino: Path, *, python_motores: Path,
                      repo_dir: Path, markitdown: str) -> list[str]:
    """Lista argv para convertir `origen` en `destino`. La usan tanto
    convert() como copy_cli(), para que el comando copiado sea exactamente
    el que ejecuta la app. Expande "~" en todas las rutas: shlex.join la
    entrecomillaría y la shell ya no la expandiría."""
    def ruta(p) -> str:
        try:
            return str(Path(p).expanduser())
        except RuntimeError:  # "~usuario" que no existe: se deja tal cual
            return str(p)

    motor = elegir_motor(Path(origen))
    if motor:
        return [ruta(python_motores), ruta(Path(repo_dir) / f"{motor}.py"),
                ruta(origen), ruta(destino)]
    return [ruta(markitdown), ruta(origen), "-o", ruta(destino)]


def comando_shell(cmd: list[str]) -> str:
    """argv → texto pegable en una shell. shlex.join escapa comillas,
    backticks, $() y espacios de verdad (hallazgo de careo, MEDIA, 23/08)."""
    return shlex.join(cmd)


# Solo NUESTRA cabecera exacta: un documento de origen con su propio
# frontmatter que empiece por "fuente:" no se toca (hallazgo de careo,
# MEDIA, 04/10: una detección por prefijo borraba contenido legítimo).
_NUESTRO_FRONTMATTER = re.compile(
    r'\A---\nfuente: "(?:[^"\\\n]|\\.)*"\nmotor: "(?:[^"\\\n]|\\.)*"\n'
    r'convertido: \d{4}-\d{2}-\d{2}\n---\n\n?'
)


def _yaml_str(texto: str) -> str:
    r"""Escalar YAML entre comillas dobles. Escapa como \uXXXX todo lo que
    YAML no admite tal cual o reinterpreta (controles C0/C1, DEL, NEL,
    separadores de línea, BOM, no-caracteres): con json.dumps a secas un
    nombre con \x85 o \x7f rompía o alteraba la cabecera."""
    salida = []
    for ch in texto:
        cp = ord(ch)
        if ch == '"':
            salida.append('\\"')
        elif ch == "\\":
            salida.append("\\\\")
        elif (cp < 0x20 or 0x7F <= cp <= 0x9F or cp in (0x2028, 0x2029, 0xFEFF)
              or 0xD800 <= cp <= 0xDFFF or cp in (0xFFFE, 0xFFFF)):
            salida.append(f"\\u{cp:04x}")
        else:
            salida.append(ch)
    return '"' + "".join(salida) + '"'


def con_frontmatter(md: str, origen: Path, motor: str, fecha: dt.date) -> str:
    """Antepone una cabecera YAML de procedencia (fichero de origen, motor,
    fecha) para que el RAG pueda citar y filtrar por fuente. Solo el nombre
    del fichero, nunca la ruta: la ruta local no aporta al RAG y filtraría
    la estructura de carpetas del equipo. Si el texto ya trae una cabecera
    nuestra, se sustituye, no se apila."""
    # Un BOM inicial (algunos conversores lo escriben) quedaba incrustado
    # detrás de la cabecera, y escondía una cabecera nuestra previa al
    # regex anclado en \A: se apilaban dos.
    md = md.lstrip("\ufeff")
    md = _NUESTRO_FRONTMATTER.sub("", md, count=1)
    cabecera = (
        "---\n"
        f"fuente: {_yaml_str(Path(origen).name)}\n"
        f"motor: {_yaml_str(motor)}\n"
        f"convertido: {fecha.isoformat()}\n"
        "---\n\n"
    )
    return cabecera + md


def escribir_atomico(destino: Path, texto: str) -> None:
    """Escribe a un temporal del mismo directorio y lo renombra: write_text
    directo trunca primero, y un fallo a mitad (disco lleno, surrogate
    suelto que no codifica en UTF-8) dejaba vacío o a medias un .md que
    ya estaba bien (hallazgos de careo, BAJA, 04/10). Crea la carpeta."""
    destino.parent.mkdir(parents=True, exist_ok=True)
    temporal = destino.with_name(f".{destino.name}.tmp")
    try:
        temporal.write_text(texto, encoding="utf-8")
        os.replace(temporal, destino)
    finally:
        temporal.unlink(missing_ok=True)


def anadir_frontmatter(destino: Path, origen: Path, motor: str, fecha: dt.date) -> None:
    """Reescribe `destino` (ya generado por el motor) con la cabecera."""
    md = destino.read_text(encoding="utf-8")
    escribir_atomico(destino, con_frontmatter(md, origen, motor, fecha))


def misma_ruta(a: Path, b: Path) -> bool:
    """¿Apuntan `a` y `b` al mismo fichero (alias, "..", symlink, o
    mayúsculas en un disco que no las distingue)? Si destino == origen, el
    motor machacaba el documento original con el Markdown."""
    try:
        return os.path.samefile(Path(a).expanduser(), Path(b).expanduser())
    except OSError:  # alguno no existe: comparar rutas resueltas
        return Path(a).expanduser().resolve() == Path(b).expanduser().resolve()


def ruta_libre(repo: Path, nombre: str) -> Path:
    """repo/nombre.md, o nombre-2.md, nombre-3.md... el primero libre. Para
    ofrecer "guardar como copia" en vez de pisar un .md existente (que
    puede llevar retoques a mano o ya estar indexado en el RAG)."""
    candidata = repo / f"{nombre}.md"
    n = 2
    while candidata.exists():
        candidata = repo / f"{nombre}-{n}.md"
        n += 1
    return candidata


def _ruta_de_campo(valor) -> Path:
    """Ruta escrita o pegada en un campo. Se recortan los espacios de los
    extremos (una ruta pegada suele traerlos) salvo que la ruta tal cual
    exista: macOS admite nombres como "Informe .pdf"."""
    def expandir(texto):
        try:
            return Path(texto).expanduser()
        except RuntimeError:  # "~usuario" que no existe: se deja tal cual
            return Path(texto)

    tal_cual = expandir(str(valor))
    return tal_cual if tal_cual.exists() else expandir(str(valor).strip())


@dataclass
class Validacion:
    destino: Path | None = None
    origen: Path | None = None  # ruta ya recortada y expandida
    nombre: str | None = None
    error: tuple[str, str] | None = None  # (título, mensaje) para el diálogo


def validar_conversion(origen: str | Path, tecleado: str, repo: str | Path, *, hay_soffice: bool,
                       comprobar_existe: bool = True) -> Validacion:
    """Todas las comprobaciones previas a convertir, en un solo sitio: las
    usan "Convertir" y "Copiar comando CLI". Antes copy_cli no aplicaba
    ninguna y copiaba comandos que escribían encima del origen, aceptaban
    "~$x.xlsx" o cambiaban en silencio un nombre ".." por "salida"
    (hallazgo de careo, BAJA, 04/10). `comprobar_existe=False` para la
    plantilla de comando que se copia sin documento elegido."""
    # Acepta el texto tal cual de los campos: Convertir y Copiar recortaban
    # (o no) los espacios cada uno a su manera, y "~usuario" inexistente
    # hacía reventar expanduser() (hallazgos de careo, BAJA, 04/10).
    origen = _ruta_de_campo(origen)
    repo = _ruta_de_campo(repo)
    if comprobar_existe and origen.is_dir():
        return Validacion(error=(
            "Error",
            f"{origen.name} es una carpeta (o un paquete de macOS, como los "
            "de Pages/Keynote), no un documento.",
        ))
    if comprobar_existe and not origen.is_file():
        return Validacion(error=("Error", "El documento no existe o no lo has seleccionado bien."))
    if repo.exists() and not repo.is_dir():
        return Validacion(error=(
            "Carpeta repo inválida",
            f"{repo} es un fichero, no una carpeta. Elige la carpeta donde guardar el .md.",
        ))
    if es_bloqueo_office(origen):
        return Validacion(error=(
            "No es un documento",
            f"{origen.name} es el fichero de bloqueo que deja Office mientras "
            "el documento está abierto. Elige el documento original, el que "
            "no empieza por \"~$\" (Word puede recortar sus primeras letras "
            "en el nombre del bloqueo).",
        ))
    motivo = motivo_no_soportado(origen, hay_soffice)
    if motivo:
        return Validacion(error=("Formato no soportado", motivo))
    nombre = nombre_salida(tecleado, origen)
    if nombre is None:
        return Validacion(error=(
            "Nombre inválido",
            "El nombre del .md no puede estar vacío ni ser solo una ruta (\".\"/\"..\").",
        ))
    destino = Path(repo) / f"{nombre}.md"
    if misma_ruta(origen, destino):
        return Validacion(error=(
            "Destino igual al origen",
            f"{destino.name} es el propio documento de origen: convertirlo "
            "ahí lo sobrescribiría. Cambia el nombre del .md o la carpeta.",
        ))
    return Validacion(destino=destino, nombre=nombre, origen=origen)
