# Conciliación contable

Aplicación web para comparar movimientos de un sistema contable con una planilla de caja o banco. Permite importar archivos reales, revisar las coincidencias propuestas y exportar un Excel con el resultado, sin enviar la información contable a un servidor.

**Link to project:** <https://andresgonzalez.netlify.app/conciliapp/>

![Vista de Conciliación contable](./conciliacion-contable.png)

## How It's Made:

**Tech used:** HTML5, CSS3, JavaScript, SheetJS, Lucide Icons, Web Workers

La aplicación está construida como una herramienta completamente del lado del cliente. SheetJS se utiliza para leer archivos XLSX, XLS y CSV, y también para generar el libro Excel final con hojas de resumen, conciliaciones, pendientes, datos originales y errores de importación.

Como las planillas contables no siempre comienzan en la primera fila ni tienen una única tabla, el importador busca posibles encabezados, permite elegir la hoja y el bloque de columnas, y deja corregir manualmente el mapeo de fecha, descripción, monto, débito y crédito. Los valores originales se conservan para la exportación, mientras que las fechas, importes y descripciones se normalizan internamente para poder compararlos.

El motor de conciliación busca coincidencias uno a uno y agrupaciones uno a varios o varios a uno. Cada propuesta recibe un puntaje basado en importe, cercanía de fecha, similitud de descripción y referencias numéricas. Los resultados claros pueden confirmarse automáticamente, mientras que los casos ambiguos quedan disponibles para revisión manual.

## Optimizations

La primera versión comparaba potencialmente cada movimiento de una tabla contra todos los de la otra. El motor fue refactorizado para crear índices por fecha, signo e importe y trabajar solamente con movimientos dentro de la ventana configurada. En una prueba sintética de 10.000 movimientos por tabla, esto redujo 100 millones de cruces teóricos a aproximadamente 820.000 parejas candidatas.

El procesamiento principal se ejecuta en un Web Worker cuando el navegador lo permite, evitando bloquear la interfaz. También se agregaron una estimación previa de carga, límites independientes para parejas y agrupaciones, procesamiento cancelable, paginación y advertencias para configuraciones costosas.

## Lessons Learned:

La parte más compleja no fue leer un archivo Excel, sino aceptar planillas con títulos, filas auxiliares, varias tablas y convenciones de signos diferentes sin depender de un formato rígido. Esto hizo importante combinar detección automática con controles manuales visibles.

También aprendí que una conciliación contable no debería tratar una coincidencia aproximada como una certeza. Separar coincidencias confirmadas, posibles y pendientes, explicar el puntaje y permitir editar o quitar agrupaciones resultó tan importante como el algoritmo de búsqueda.

Finalmente, trabajar con miles de movimientos mostró la diferencia entre optimizar solamente el tiempo de cálculo y cuidar toda la experiencia: estimar el costo, mantener la interfaz disponible, permitir cancelar y limitar las combinaciones son partes del mismo problema.
