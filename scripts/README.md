# scripts/

En esta carpeta están los scripts en R para el procesamiento de datos.

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

