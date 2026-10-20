"""
Conversor de PDF/Word/PowerPoint (y escaneados) a Markdown con Docling.
Se usa para estos formatos por su reconstrucción de tablas; los resultados
requieren revisión. La descarga inicial de modelos puede necesitar red.

Import diferido: docling carga modelos pesados (layout/tabla) la primera
vez que se usa. Se importa dentro de convertir(), no al importar el
módulo, para que cargar este fichero sea barato -- solo paga el coste
quien de verdad convierte un documento (invocado siempre como subprocess
en un venv aparte, nunca dentro del proceso de la GUI Tkinter).

CLI: python motor_docling.py <origen> <destino.md>
Librería: from motor_docling import convertir
"""
import sys
import tempfile
from pathlib import Path

from all2md_core import EXTENSIONES_TEXTUTIL, escribir_atomico, misma_ruta
from preconversion import con_soffice, con_textutil


class ConversionVacia(Exception):
    """Docling no produjo ningún contenido."""


def preconvertir(origen: Path, carpeta: Path) -> Path:
    """Lo que Docling no lee bien, a un formato que sí:
    - .doc/.rtf/.odt → HTML con textutil (macOS). HTML y no .docx: el
      .docx de textutil aplana las tablas, y la ruta nativa de Docling para
      .doc también (pasa por LibreOffice y saca "ConceptoImporte...").
    - .ppt → .pptx con LibreOffice.
    Lo demás se devuelve tal cual."""
    ext = origen.suffix.lower()
    if ext in EXTENSIONES_TEXTUTIL:
        return con_textutil(origen, carpeta)
    if ext == ".ppt":
        return con_soffice(origen, "pptx", carpeta)
    return origen


def convertir(origen: Path) -> str:
    from docling.document_converter import DocumentConverter

    converter = DocumentConverter()
    with tempfile.TemporaryDirectory(prefix="all2md-") as tmp:
        resultado = converter.convert(str(preconvertir(origen, Path(tmp))))
    md = resultado.document.export_to_markdown()
    if not md.strip():
        raise ConversionVacia(f"Docling no produjo contenido para {origen.name}.")
    return md


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit("Uso: python motor_docling.py <origen> <destino.md>")
    origen = Path(sys.argv[1]).expanduser()
    destino = Path(sys.argv[2]).expanduser()
    if not origen.exists():
        sys.exit(f"No existe: {origen}")
    if misma_ruta(origen, destino):
        sys.exit(f"El destino es el propio documento de origen ({origen.name}): "
                 "no se sobrescribe.")

    try:
        md = convertir(origen)
    except ConversionVacia as e:
        sys.exit(str(e))
    except Exception as e:
        sys.exit(f"motor_docling falló con {origen.name}: {e}")

    try:
        escribir_atomico(destino, md)
    except (OSError, UnicodeError) as e:
        # Sin esto, un destino no escribible acababa en traceback crudo,
        # que la app enseña tal cual en el diálogo de error.
        sys.exit(f"No se pudo escribir {destino}: {e}")
    print(f"Guardado en {destino} ({len(md)} caracteres)")
