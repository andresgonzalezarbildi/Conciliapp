# Conciliación contable

Aplicación web para comparar movimientos de un sistema contable con una planilla de caja o banco. Funciona directamente en el navegador con HTML, CSS, JavaScript y SheetJS, sin backend ni base de datos.

![Vista de revisión de la conciliación](conciliacion-contable.png)

## Qué permite hacer

- Cargar archivos XLSX, XLS y CSV sin recortarlos previamente.
- Elegir la hoja, el rango de filas, la tabla interna y las columnas que se deben utilizar.
- Interpretar importes con signo o columnas separadas de Débito y Crédito, incluso cuando ambas columnas traen valores positivos.
- Elegir la dirección contable: Débito positivo y Crédito negativo, la convención inversa o conservar el signo escrito.
- Buscar coincidencias uno a uno, uno a varios, varios a uno y agrupaciones masivas exactas con importes de signos mixtos.
- Detectar y crear compensaciones internas cuando Débitos y Créditos de una misma tabla dejan un neto de cero.
- Interpretar automáticamente un extracto Débito/Crédito y un mayor Debe/Haber con la orientación de signos correspondiente.
- Priorizar monto y fecha exactos cuando la pareja es única, aunque la descripción use otro nombre comercial.
- Evitar la penalización por tamaño en agrupaciones que comparten claramente el mismo nombre o referencia comercial.
- Configurar tolerancias de fecha e importe, signos y similitud de descripciones.
- Revisar, quitar, editar o crear conciliaciones manualmente.
- Consultar resultados en tablas compactas que se reorganizan para evitar desplazamiento horizontal en pantallas medianas.
- Ejecutar automáticamente una segunda pasada moderadamente flexible sobre los pendientes.
- Abrir una búsqueda avanzada para ajustar fechas, montos y similitud, incluyendo propuestas uno a uno sin límite de fecha.
- Recordar durante la sesión las agrupaciones rechazadas para no volver a proponerlas.
- Agregar observaciones visibles que se incluyen en el Excel exportado.
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
| Movimientos máximos por lado en una agrupación | 100 movimientos |
| Tope de candidatos uno a uno | 2.000.000 |
| Intentos de suma por movimiento | 25.000 |
| Agrupaciones con Débitos y Créditos mezclados | Activado |
| Compensaciones dentro de la misma tabla | Activado |
| Conciliación automática | 70 puntos |
| Posible conciliación | 55 puntos |
