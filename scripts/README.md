# scripts/

En esta carpeta están los scripts en R para el procesamiento de datos.

## Aistencias:

- [1.1_asistencias_excel.R](/scripts/1.1_asistencias_excel.R)

Este script en R procesa un archivo descargado desde el Aula Virtual en formato `.xlsx` con registros de asistencia de estudiantes, realiza limpieza y transformación de datos, y genera un nuevo archivo Excel con las siguientes características:

1. **Limpieza de datos**:
   - Elimina columnas innecesarias.
   - Normaliza los valores de asistencia (`P`, `R`, `FJ`, `FI`) y reemplaza valores desconocidos (`?`) por `FI`.

2. **Separación por grupos**:
   - Divide los datos en dos grupos (`GRUPO 2` y `GRUPO 3`) según la columna `Grupos`.
   - Elimina columnas irrelevantes para cada grupo.

3. **Cálculo del porcentaje de asistencia**:
   - Agrega una columna que calcula el porcentaje de asistencia considerando `P`, `R` y `FJ` como asistencias válidas.

4. **Formato condicional**:
   - Resalta en rojo (`FI`) y verde (`R`) los valores de asistencia.
   - Resalta en rojo los estudiantes con porcentaje de asistencia menor al 80%.

5. **Exportación**:
   - Genera un archivo Excel con dos hojas (`GRUPO_2` y `GRUPO_3`), incluyendo los datos procesados y el formato condicional.

Este script es útil para analizar y visualizar la asistencia de estudiantes de manera clara y organizada.

## Calificaciones:

![alt text](<Captura de pantalla 2025-05-06 a la(s) 8.57.28 p.m..png>)

Después de escoger la opción **.xlsx** para descargar los datos, obtendrás un archivo llamado **Historial de notas.xlsx**. Este archivo va en la carpeta **data** y es el input para el script **1.1_process_from_excel.R**

- [1.1_process_from_excel.R](/scripts/1.1_process_from_excel.R)

Este script en R procesa un archivo descargado desde el Aula Virtual llamado `Historial de notas.xlsx` para consolidar y analizar las calificaciones de los estudiantes. Realiza las siguientes tareas:

1. **Limpieza y filtrado de datos**:
   - Elimina registros no deseados, como los relacionados con "asistencia", "cuestionario inicial" o "PRE-TEST".
   - Convierte las calificaciones a formato numérico (por ejemplo, de `8,00` a `8.00`).
   - Elimina filas con calificaciones vacías o valores no válidos.

2. **Consolidación de calificaciones**:
   - Agrupa las calificaciones por estudiante y actividad.
   - Conserva la calificación más alta en caso de múltiples intentos para una misma actividad.

3. **Transformación a formato ancho**:
   - Convierte los datos a un formato donde cada actividad es una columna, facilitando el análisis.

4. **Cálculo de promedios**:
   - Agrega una columna llamada `Promedio` que calcula el promedio de las calificaciones de cada estudiante.

5. **Formato condicional**:
   - Resalta en **rojo** las celdas con valor `0`.
   - Resalta en **naranja** las celdas con valores entre `1` y `5`, incluyendo la columna `Promedio`.

6. **Exportación**:
   - Genera un archivo Excel llamado `calificaciones_consolidado.xlsx` en la carpeta `output`.
   - El archivo contiene las siguientes hojas:
     - **Grupo2_SinFormato**: Datos originales del Grupo 2.
     - **Grupo3_SinFormato**: Datos originales del Grupo 3.
     - **Grupo2_ConFormato**: Datos del Grupo 2 con formato condicional aplicado.
     - **Grupo3_ConFormato**: Datos del Grupo 3 con formato condicional aplicado.

Este script es útil para consolidar y visualizar las calificaciones de los estudiantes de manera clara y organizada, facilitando la identificación de áreas de mejora.