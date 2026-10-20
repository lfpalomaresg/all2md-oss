from __future__ import annotations
import datetime
import os
import subprocess
import sys
import shutil
import tempfile
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import all2md_core
import preconversion

HOME = Path.home()
REPO_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = HOME / "md-repo"
VENV_MARKITDOWN = HOME / "markitdown-env" / "bin" / "markitdown"

VENV_TOOLS_PY = REPO_DIR / ".venv" / "bin" / "python3"

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Documentos a Markdown → Repo")
        self.root.geometry("760x520")
        self.root.configure(bg="#f5f3ef")

        self.pdf_path = tk.StringVar()
        self.repo_path = tk.StringVar(value=str(DEFAULT_OUTPUT_DIR))
        self.filename = tk.StringVar()
        self.status = tk.StringVar(value="Selecciona un documento y una carpeta destino.")
        # LibreOffice (opcional) habilita .xls/.ppt: se mira una vez al abrir.
        self.hay_soffice = preconversion.buscar_soffice() is not None

        self.build_ui()

    def build_ui(self):
        wrap = tk.Frame(self.root, bg="#f5f3ef", padx=24, pady=24)
        wrap.pack(fill="both", expand=True)

        tk.Label(
            wrap,
            text="Documentos a Markdown → Repo",
            font=("Helvetica", 20, "bold"),
            bg="#f5f3ef",
            fg="#1f2937"
        ).pack(anchor="w")

        tk.Label(
            wrap,
            text="Mini app local para convertir PDF/Word/Excel/PowerPoint a .md (motor según tipo de documento) y guardarlo en tu repositorio de conocimiento.",
            font=("Helvetica", 11),
            bg="#f5f3ef",
            fg="#4b5563",
            wraplength=680,
            justify="left",
        ).pack(anchor="w", pady=(8, 20))

        self._field(wrap, "1) Documento de entrada (PDF/Word/Excel/PowerPoint)", self.pdf_path, self.pick_pdf, "Elegir documento")
        self._field(wrap, "2) Carpeta repo destino", self.repo_path, self.pick_repo, "Elegir carpeta")
        self._text_field(wrap, "3) Nombre del .md (sin extensión)", self.filename)

        btns = tk.Frame(wrap, bg="#f5f3ef")
        btns.pack(fill="x", pady=(18, 10))

        tk.Button(
            btns,
            text="Convertir y guardar",
            command=self.convert,
            bg="#0f766e",
            fg="white",
            padx=16,
            pady=10
        ).pack(side="left")

        tk.Button(
            btns,
            text="Abrir carpeta repo",
            command=self.open_repo,
            bg="#e5e7eb",
            fg="#111827",
            padx=16,
            pady=10
        ).pack(side="left", padx=10)

        tk.Button(
            btns,
            text="Copiar comando CLI",
            command=self.copy_cli,
            bg="#e5e7eb",
            fg="#111827",
            padx=16,
            pady=10
        ).pack(side="left")

        notes = (
            "Flujo recomendado:\n"
            "- Soporta PDF, Word (.docx/.doc/.rtf/.odt), Excel (.xlsx/.xlsm; .xls con LibreOffice) y PowerPoint (.pptx; .ppt con LibreOffice).\n"
            "- Deja tu repo de Markdown en ~/md-repo o cambia la carpeta aquí.\n"
            "- Enrutamiento automático por tipo (23/08/2026):\n"
            "    · .xlsx/.xlsm/.xltx/.xltm → motor Excel propio (openpyxl, sin IA, exacto con celdas combinadas)\n"
            "    · .pdf/.docx/.pptx → Docling (local, gratis, reconstruye tablas complejas)\n"
            "    · .doc/.rtf/.odt → textutil (macOS) a HTML → Docling\n"
            "    · .xls/.ppt (97-2003) → LibreOffice a .xlsx/.pptx → motor Excel / Docling\n"
            "    · el resto (HTML, CSV, EPUB...) → MarkItDown\n"
            "- El archivo resultante se guarda en la carpeta repo lista para tu RAG/IA."
        )

        tk.Label(
            wrap,
            text=notes,
            bg="#eef2f7",
            fg="#374151",
            justify="left",
            anchor="w",
            padx=14,
            pady=14
        ).pack(fill="x", pady=(12, 18))

        tk.Label(
            wrap,
            textvariable=self.status,
            bg="#f5f3ef",
            fg="#7c2d12",
            justify="left",
            wraplength=690
        ).pack(anchor="w")

    def _field(self, parent, label, var, command, button_text):
        tk.Label(
            parent,
            text=label,
            font=("Helvetica", 11, "bold"),
            bg="#f5f3ef",
            fg="#111827"
        ).pack(anchor="w", pady=(8, 6))

        row = tk.Frame(parent, bg="#f5f3ef")
        row.pack(fill="x")

        tk.Entry(
            row,
            textvariable=var,
            font=("Helvetica", 11),
            relief="solid",
            bd=1
        ).pack(side="left", fill="x", expand=True, ipady=8)

        tk.Button(
            row,
            text=button_text,
            command=command,
            bg="#d1fae5",
            fg="#065f46",
            padx=12,
            pady=8
        ).pack(side="left", padx=(10, 0))

    def _text_field(self, parent, label, var):
        tk.Label(
            parent,
            text=label,
            font=("Helvetica", 11, "bold"),
            bg="#f5f3ef",
            fg="#111827"
        ).pack(anchor="w", pady=(16, 6))

        tk.Entry(
            parent,
            textvariable=var,
            font=("Helvetica", 11),
            relief="solid",
            bd=1
        ).pack(fill="x", ipady=8)

    def pick_pdf(self):
        # El filtro sale de las mismas extensiones que enrutan (all2md_core):
        # así no se vuelve a quedar atrás cuando se añade un formato.
        path = filedialog.askopenfilename(filetypes=all2md_core.tipos_dialogo(self.hay_soffice))
        if path:
            self.pdf_path.set(path)
            if not self.filename.get():
                self.filename.set(Path(path).stem)

    def pick_repo(self):
        path = filedialog.askdirectory()
        if path:
            self.repo_path.set(path)

    def get_markitdown_cmd(self):
        if VENV_MARKITDOWN.exists():
            return str(VENV_MARKITDOWN)
        found = shutil.which("markitdown")
        return found

    def convert(self):
        name = self.filename.get().strip()
        v = all2md_core.validar_conversion(self.pdf_path.get(), name, self.repo_path.get(),
                                           hay_soffice=self.hay_soffice)
        if v.error:
            messagebox.showerror(*v.error)
            return
        pdf, repo = v.origen, v.destino.parent
        nombre, out_file = v.nombre, v.destino
        if not name:
            self.filename.set(nombre)
        try:
            repo.mkdir(parents=True, exist_ok=True)
        except OSError as e:  # un tramo de la ruta es un fichero, sin permisos...
            messagebox.showerror("Carpeta repo inválida", f"No se pudo crear {repo}:\n\n{e}")
            return

        # No se comprueba en all2md_core si `.venv` existe -- se decide aquí
        # para avisar con un mensaje específico en vez de degradar en
        # silencio a MarkItDown (hallazgo de careo, MEDIA, 23/08).
        motor = all2md_core.elegir_motor(pdf)
        if motor:
            if not VENV_TOOLS_PY.exists():
                messagebox.showerror(
                    "Falta el venv de motores",
                    f"{pdf.suffix} necesita el motor {motor}, pero no encuentro "
                    f"{VENV_TOOLS_PY}. Instálalo con:\n\n"
                    f"cd {REPO_DIR}\npython3 -m venv .venv\n"
                    ".venv/bin/pip install -r requirements-motores.txt"
                )
                return
            markitdown_cmd = ""
            motor_legible = {"motor_excel": "motor Excel (openpyxl, sin IA)", "motor_docling": "Docling (local, gratis)"}[motor]
        else:
            markitdown_cmd = self.get_markitdown_cmd()
            if not markitdown_cmd:
                messagebox.showerror(
                    "MarkItDown no encontrado",
                    "No encuentro markitdown. Instálalo en ~/markitdown-env o en tu PATH."
                )
                return
            motor_legible = "MarkItDown"
        # Después de comprobar motores: no preguntar para luego fallar.
        if out_file.exists():
            copia = all2md_core.ruta_libre(repo, nombre)
            respuesta = messagebox.askyesnocancel(
                "El .md ya existe",
                f"Ya existe {out_file.name} en la carpeta repo.\n\n"
                "Sí → sobrescribirlo\n"
                f"No → guardar como {copia.name}\n"
                "Cancelar → no convertir",
                default=messagebox.NO,  # Enter no debe pisar nada
            )
            if respuesta is None:
                return
            if respuesta is False:
                out_file = copia

        # El motor escribe en un temporal de la misma carpeta y el destino
        # solo se reemplaza tras validar el resultado: con "sobrescribir",
        # un motor que trunca su salida y falla, o que acaba bien sin
        # escribir nada, destruía o daba por convertido el .md retocado a
        # mano (hallazgos de careo, IMPORTANTE, 15/10).
        # Carpeta temporal exclusiva de esta ejecución (no un nombre fijo,
        # que borraba un fichero ajeno o el temporal de otra conversión), y
        # dentro el fichero aún sin crear: así "existe" sigue probando que
        # el motor escribió. Misma carpeta → os.replace atómico.
        try:
            carpeta_temporal = Path(tempfile.mkdtemp(prefix=".convirtiendo-", dir=repo))
        except OSError as e:
            self.status.set(f"No se pudo escribir en {repo}: {e}")
            messagebox.showerror("Carpeta repo inválida", f"No se pudo escribir en {repo}:\n\n{e}")
            return
        temporal = carpeta_temporal / out_file.name
        cmd = all2md_core.construir_comando(
            pdf, temporal, python_motores=VENV_TOOLS_PY, repo_dir=REPO_DIR,
            markitdown=markitdown_cmd,
        )

        try:
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            # Código 0 sin fichero (un conversor que no reconoce el formato
            # y "acaba bien"): no se puede anunciar un .md que no existe.
            if not temporal.exists():
                detalle = (result.stderr or result.stdout or "").strip()
                self.status.set(f"{motor_legible} terminó sin generar {out_file.name}.\n{detalle[:1200]}")
                messagebox.showerror(
                    "Falló la conversión",
                    f"Motor: {motor_legible}\n\nTerminó sin errores pero no generó "
                    f"{out_file}.\n\n{detalle[:1500]}",
                )
                return
            msg = f"OK ({motor_legible}). Markdown generado en:\n{out_file}"
            try:
                all2md_core.anadir_frontmatter(
                    temporal, pdf, motor or "markitdown", datetime.date.today()
                )
            except (OSError, UnicodeError) as e:
                # La conversión ya está hecha: sin cabecera sigue sirviendo.
                msg += f"\n\n(No se pudo añadir la cabecera de procedencia: {e})"
            try:
                os.replace(temporal, out_file)
            except OSError as e:
                self.status.set(f"No se pudo guardar {out_file}: {e}")
                messagebox.showerror("Falló la conversión", f"No se pudo guardar {out_file}:\n\n{e}")
                return
            salida = result.stdout.replace(str(temporal), str(out_file))
            if salida.strip():
                msg += f"\n\nSalida:\n{salida[:800]}"
            self.status.set(msg)
            messagebox.showinfo("Conversión completada", msg)
        except subprocess.CalledProcessError as e:
            err = e.stderr or e.stdout or str(e)
            self.status.set(f"Error al convertir con {motor_legible}:\n{err[:1200]}")
            messagebox.showerror("Falló la conversión", f"Motor: {motor_legible}\n\n{err[:1500]}")
        except OSError as e:
            # El ejecutable existe pero no se pudo lanzar (permisos, venv
            # roto, etc.) -- CalledProcessError no cubre esto (hallazgo de
            # careo, MEDIA, 23/08: antes terminaba en traceback crudo).
            self.status.set(f"No se pudo lanzar {motor_legible}: {e}")
            messagebox.showerror("Falló la conversión", f"No se pudo lanzar {motor_legible}:\n\n{e}")
        finally:
            # Solo lo creado por esta ejecución; un fallo al limpiar no debe
            # tapar el mensaje ya mostrado con un traceback.
            shutil.rmtree(carpeta_temporal, ignore_errors=True)

    def open_repo(self):
        repo = Path(self.repo_path.get()).expanduser()
        repo.mkdir(parents=True, exist_ok=True)
        subprocess.run(["open", str(repo)])

    def copy_cli(self):
        elegido = self.pdf_path.get().strip()

        # Mismas comprobaciones que convert(): sin esto se copiaban comandos
        # que pisaban el origen o convertían un "~$x.xlsx". Sin documento
        # elegido se copia una plantilla, sin exigir que exista.
        v = all2md_core.validar_conversion(elegido or "/ruta/a/tu.pdf", self.filename.get().strip(),
                                           self.repo_path.get(), hay_soffice=self.hay_soffice,
                                           comprobar_existe=bool(elegido))
        if v.error:
            self.status.set(f"No se copió el comando: {v.error[1]}")
            messagebox.showerror(*v.error)
            return

        # Mismo constructor que convert(): el comando copiado es el que
        # ejecutaría la app, con MarkItDown resuelto a ruta absoluta.
        cmd = all2md_core.construir_comando(
            v.origen, v.destino, python_motores=VENV_TOOLS_PY, repo_dir=REPO_DIR,
            markitdown=self.get_markitdown_cmd() or "markitdown",
        )

        self.root.clipboard_clear()
        self.root.clipboard_append(all2md_core.comando_shell(cmd))
        self.status.set("Comando copiado al portapapeles.")

def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    if sys.platform != "darwin":
        print("Esta mini app está pensada para macOS, pero se puede adaptar fácilmente.")
    main()
