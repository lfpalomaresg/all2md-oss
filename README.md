# ALL2MD

Aplicación local Python/Tkinter para convertir documentos a Markdown y guardarlos en una carpeta de conocimiento para IA y RAG.

ALL2MD aporta la interfaz gráfica, el enrutamiento entre motores, el tratamiento específico de Excel y el control del archivo de salida. Integra openpyxl, Docling y Microsoft MarkItDown, que mantienen sus propias atribuciones y licencias.

## Estado

Proyecto temprano, orientado a macOS, con instalación manual. El uso documentado es propio y puntual; no hay adopción externa acreditada, instalador distribuido ni soporte verificado para Windows o Linux.

## Formatos y motores

- `.xlsx`, `.xlsm`, `.xltx`, `.xltm` → `motor_excel.py` mediante openpyxl.
- `.pdf`, `.docx`, `.pptx` → `motor_docling.py` mediante Docling.
- `.doc`, `.rtf`, `.odt` → textutil de macOS, HTML y `motor_docling.py`.
- `.xls` → LibreOffice, `.xlsx` y `motor_excel.py`.
- `.ppt` → LibreOffice, `.pptx` y `motor_docling.py`.
- Otros formatos → MarkItDown, según sus dependencias y formatos disponibles.

## Instalación en macOS

Requiere Python 3 con Tkinter funcional. El código propio usa sintaxis de Python 3.10 o posterior, pero las dependencias pueden exigir versiones adicionales. Comprueba también sus requisitos. La instalación nueva completa aún debe verificarse.

Desde la carpeta del repositorio:

    python3 -m tkinter
    python3 -m venv .venv
    .venv/bin/python3 -m pip install -r requirements-motores.txt
    python3 -m venv "$HOME/markitdown-env"
    "$HOME/markitdown-env/bin/python3" -m pip install "markitdown[all]"

La primera orden debe abrir una ventana de prueba de Tkinter. Se ha utilizado Tk 8.6. La aplicación busca MarkItDown primero en ~/markitdown-env/bin/markitdown y después en PATH. LibreOffice es opcional para .xls/.ppt.

Las dependencias y ciertos modelos requieren descargas iniciales. La conversión prevista es local; no se afirma que todas las fases funcionen siempre sin red.

## Uso

    .venv/bin/python3 all2md_repo_app.py

Selecciona documento, carpeta de destino y nombre. Pulsa Convertir y guardar y revisa el Markdown. Si existe el destino, puedes conservar una copia, sobrescribir o cancelar. La GUI usa un temporal antes de reemplazar el archivo de destino.

La interfaz puede quedar ocupada durante la conversión; no hay cancelación del proceso en curso. Copiar comando CLI copia el comando del motor, sin todas las comprobaciones e interacciones de la GUI. ALL2MD no incluye servidor MCP propio.

## Límites

- Excel no se recalcula: guarda el libro recalculado antes de convertirlo.
- Las hojas ocultas se incluyen y se etiquetan: revisa el resultado antes de compartirlo.
- Se omiten con aviso hojas de más de 5000 filas o 500 columnas.
- textutil no conserva imágenes ni toda la estructura de títulos de formatos antiguos.
- El OCR y las tablas requieren revisión; no se garantiza conversión perfecta.
- No se garantiza ejecutar varios CLI simultáneamente sobre un mismo destino.
- Las versiones de Docling y openpyxl no están fijadas; la instalación no es completamente reproducible.
- Se conserva el pin de OpenCV por compatibilidad con el entorno original.

## Pruebas

    .venv/bin/python3 -m pip install -r requirements-dev.txt
    .venv/bin/python3 -m pytest -q

Se usan datos sintéticos, temporales y motores simulados. La suite completa está orientada a macOS/POSIX. Algunas pruebas requieren textutil, LibreOffice o /bin/sh. Una ejecución parcial no verifica los motores reales ni la GUI.

- tests/test_core.py: enrutamiento, nombres, comandos, procedencia y escritura.
- tests/test_motor_excel.py: formatos, celdas combinadas, fórmulas y límites.
- tests/test_cli_motores.py: contratos de los CLI y protección del origen.
- tests/test_motor_docling.py: preconversión con textutil.
- tests/test_preconversion.py: detección y conversión con LibreOffice.
- tests/test_preconversion_errores.py: fallos y timeouts simulados.
- tests/test_motor_docling_convertir.py: enrutamiento y Docling simulado.
- tests/test_entradas_externas.py: nombres, rangos y codificaciones.
- tests/test_preconversion_ejecutables_falsos.py: herramientas falsas de shell.
- tests/test_app_convertir.py: GUI con Tkinter y motores simulados.
- tests/test_motor_docling_cli.py: errores y salidas del CLI simulado.
- tests/test_red_sintaxis.py: sintaxis e importación sin abrir la ventana.
- tests/test_docs.py: correspondencia entre README, formatos y pruebas.

## Contribuciones y seguridad

Aporta casos reproducibles y documentos sintéticos. No publiques datos de clientes, documentos de trabajo, claves o rutas personales. Las propuestas deben explicar el problema, el cambio y las verificaciones realizadas.

Consulta SECURITY.md antes de comunicar una vulnerabilidad. El canal privado está pendiente de establecer; no publiques detalles sensibles en una incidencia pública.

## Licencia

MIT para el código propio incluido; consulta LICENSE. Las dependencias y modelos conservan sus respectivas licencias.
