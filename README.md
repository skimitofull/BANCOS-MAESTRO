# BANCOS-MAESTRO

Respaldo del convertidor local **Estados de cuenta: PDF → Excel → PDF**, tomado el 10 de noviembre de 2026.

Incluye el programa, lectores bancarios, generador didáctico de movimientos ficticios, periodos enlazados, revisión tipográfica, fuentes, plantillas, ejemplos y respaldos históricos presentes en la carpeta original. Solo se excluye la caché regenerable de Python.

## Restaurar y ejecutar en Windows

1. Descargar o clonar este repositorio.
2. Abrir la carpeta `convertidor_estados_cuenta_local`.
3. Con Python instalado, ejecutar `instalar_dependencias.bat`.
4. Ejecutar `ejecutar_app.bat`.

Los ejercicios ficticios son material didáctico. Conservar las leyendas que los identifican.

## Integridad del respaldo

`RESPALDO_SHA256.json` registra tamaño y SHA-256 de cada archivo copiado. Se compararon los bytes de origen y destino y se comprobó la sintaxis de los archivos Python. Los archivos Office se copiaron sin modificar; esta operación no prueba su apertura o recálculo en Excel.

`.gitattributes` evita que Git transforme saltos de línea: se conservan los bytes originales del respaldo.
