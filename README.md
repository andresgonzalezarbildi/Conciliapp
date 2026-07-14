# Conciliación contable

Aplicación web para comparar movimientos de un sistema contable con una planilla de caja o banco. Funciona directamente en el navegador con HTML, CSS, JavaScript y SheetJS, sin backend ni base de datos.

![Vista de revisión de la conciliación](conciliacion-contable.jpg)

## Qué permite hacer

- Cargar archivos XLSX, XLS y CSV sin recortarlos previamente.
- Elegir la hoja, el rango de filas, la tabla interna y las columnas que se deben utilizar.
- Interpretar importes con signo o columnas separadas de Débito y Crédito.
- Buscar coincidencias uno a uno, uno a varios y varios a uno.
- Configurar tolerancias de fecha e importe, signos y similitud de descripciones.
- Revisar, quitar, editar o crear conciliaciones manualmente.
- Procesar tablas grandes mediante índices, límites de búsqueda y Web Worker cuando el navegador lo admite.
- Exportar un archivo Excel con resumen, conciliaciones, pendientes, datos originales y errores de importación.

Los archivos se procesan en memoria y sus movimientos no se envían a un servidor ni se guardan en `localStorage`.

## Puesta en marcha

No requiere instalar Node.js ni configurar una base de datos.

1. Descomprime el archivo ZIP manteniendo todos los archivos en la misma carpeta.
2. Abre `index.html` con un navegador actualizado.
3. Carga las dos planillas y sigue los pasos de importación, configuración y revisión.

La primera carga necesita conexión a Internet para obtener SheetJS y Lucide Icons desde sus CDN. El procesamiento de las planillas continúa realizándose localmente.

## Configuración

Los parámetros se pueden modificar desde el paso **Configuración** antes de cada conciliación. Sus valores iniciales están definidos al comienzo de `app.js`, dentro de `DEFAULT_CONFIG`.

| Parámetro | Valor inicial |
| --- | ---: |
| Tolerancia de fecha | 1 día |
| Tolerancia absoluta de importe | 0,10 |
| Máximo por agrupación | 8 movimientos |
| Límite de parejas candidatas | 2.000.000 |
| Límite de combinaciones agrupadas | 25.000 |
| Conciliación automática | 70 puntos |
| Posible conciliación | 55 puntos |
