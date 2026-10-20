# Contribuir a ALL2MD

Se agradecen informes de errores, mejoras de documentación y cambios pequeños con un caso reproducible. ALL2MD es una aplicación Python/Tkinter orientada a macOS; no presupongas compatibilidad con Windows o Linux.

## Informar de un problema

Describe el resultado esperado y el observado. Incluye la versión de Python, el sistema operativo, el formato del documento y el motor utilizado, junto con los pasos mínimos para reproducirlo. Si compartes un error o una ruta, elimina primero los datos personales.

Utiliza documentos sintéticos creados para la prueba. No adjuntes documentos de clientes o del trabajo, credenciales, archivos `.env`, información privada ni resultados que la contengan. Para posibles vulnerabilidades, consulta [SECURITY.md](SECURITY.md).

## Preparar un cambio

1. Trabaja en una rama y limita el cambio a un problema concreto.
2. Mantén la separación entre la interfaz, el enrutamiento, los motores y las preconversiones.
3. Si cambia el comportamiento, añade o adapta un caso de prueba con datos sintéticos y archivos temporales.
4. Actualiza el README si cambian los formatos, el enrutamiento o las instrucciones de uso. La prueba de documentación comprueba esa correspondencia.
5. Explica en la propuesta qué problema resuelve el cambio y cómo lo has comprobado.

No incluyas entornos virtuales, modelos descargados, archivos de salida, cachés ni documentos reales. No introduzcas servicios de pago, envío de documentos a servicios externos ni nuevas dependencias sin explicar el motivo y su efecto.

## Comprobaciones

Desde la raíz del repositorio, con un entorno virtual de Python creado:

```sh
.venv/bin/python3 -m pip install -r requirements-dev.txt
.venv/bin/python3 -m pytest -q
```

La suite combina pruebas unitarias, datos sintéticos y motores simulados. Algunas comprobaciones dependen de Tkinter, herramientas de macOS, LibreOffice o un entorno POSIX. Indica las pruebas omitidas o que no hayas podido ejecutar; una suite parcial no valida la interfaz gráfica ni una conversión con los motores reales.

Para cambios en la interfaz o en la integración de un motor, realiza también una comprobación manual con un documento sintético en macOS y explica el resultado. No afirmes compatibilidad con otra plataforma sin verificarla.

## Atribución

Incluye únicamente material que tengas derecho a aportar y conserva las atribuciones de terceros. Las dependencias y los modelos mantienen sus propias licencias.
