# Cargar paquetes necesarios
library(readxl)
library(dplyr)
library(tidyverse)
library(stringr)
library(openxlsx)
# Cargar el archivo Excel
datos <- read_excel("data/Historial de notas.xlsx")

# Seleccionar y renombrar columnas (manteniendo ID como texto)
datos_procesados <- datos %>%
  select(
    nombre = "Nombre",
    id = "Número de ID",        # Se mantendrá como character
    email = "Dirección de correo",
    item = "Ítem de calificación",
    calificacion = "Calificación revisada"
  ) %>%
  mutate(
    nombre = as.character(nombre),  # Asegurar que nombre es texto
    id = as.character(id)           # Asegurar que id es texto
  )

# Verificación de tipos
cat("\nTipos de datos finales:\n")
glimpse(datos_procesados)

# Mostrar muestra de datos
cat("\nMuestra de datos:\n")
head(datos_procesados)





# 1. Filtrar y limpiar los datos
datos_filtrados <- datos_procesados %>%
  # Eliminar registros no deseados (case insensitive)
  filter(!str_detect(tolower(item), "asistencia|cuestionario inicial|pre-test")) %>%
  # Eliminar filas con calificación vacía o NA
  filter(!is.na(calificacion) & calificacion != "") %>%
  # Convertir calificaciones a numérico (si están como texto "8,00")
  mutate(calificacion = as.numeric(str_replace(calificacion, ",", ".")))

# 2. Consolidar múltiples calificaciones (quedarse con la máxima)
datos_consolidados <- datos_filtrados %>%
  group_by(nombre, id, item) %>%
  summarise(
    calificacion = max(calificacion, na.rm = TRUE),  # Conservar la máxima calificación
    .groups = "drop"
  ) %>%
  # Reemplazar Inf (cuando todos son NA) por NA
  mutate(calificacion = ifelse(is.infinite(calificacion), NA, calificacion))

# 3. Transformar a formato ancho
datos_ancho <- datos_consolidados %>%
  pivot_wider(
    names_from = item,
    values_from = calificacion,
    id_cols = c(nombre, id),
    names_sort = TRUE
  )

# 4. Eliminar columnas no deseadas en formato ancho
datos_ancho <- datos_ancho %>%
  select(-matches("(?i)pre-test|prueba de diagnóstico inicial"))

# Verificar columnas finales
cat("Columnas finales:\n")
print(colnames(datos_ancho))



# 1. Primero identificamos y renombramos las columnas según su unidad
datos_renombrados <- datos_ancho %>%
  rename_with(~ {
    case_when(
      str_detect(.x, "Recapitulación.*Sesión 1") ~ "U1_Cuestionario de Recapitulación: Sesión 1",
      str_detect(.x, "Ejercicio.*Archivos fasta") ~ "U1_Ejercicio - Archivos fasta",
      str_detect(.x, "Actividad 1.1") ~ "U1_Actividad 1.1 Bases de datos",
      str_detect(.x, "Actividad 1.2") ~ "U1_Actividad 1.2 Formato fasta",
      str_detect(.x, "TAREA.*18 de FEBRERO.*Fasta-NCBI") ~ "U1_TAREA PARA MARTES 18 de FEBRERO. Act. 1.2 Ejercicio Fasta-NCBI",
      str_detect(.x, "Recapitulación.*Sesión 2") ~ "U1_Cuestionario de Recapitulación: Sesión 2",
      str_detect(.x, "TAREA.*25 DE FEBRERO.*Actividad 2.1") ~ "U2_TAREA PARA MARTES 25 DE FEBRERO. Actividad 2.1",
      str_detect(.x, "Actividad de sistemática molecular") ~ "U2_Actividad de sistemática molecular",
      str_detect(.x, "Evaluación de sistemática molecular") ~ "U2_Evaluación de sistemática molecular",
      str_detect(.x, "Cuestionario final U2") ~ "U2_Cuestionario final U2 - NGS y formatos",
      str_detect(.x, "actividades del 11 de marzo") ~ "U3_Cuestionario de actividades del 11 de marzo. PRESENTACIÓN + VIDEO",
      str_detect(.x, "Actividad 3.1.*PDB") ~ "U3_Actividad 3.1 (PDB 2025)",
      str_detect(.x, "Cuestionario de la actividad 3.1") ~ "U3_Cuestionario de la actividad 3.1",
      str_detect(.x, "Actividad 3.2.*GFP") ~ "U3_Actividad 3.2 (GFP 2025)",
      str_detect(.x, "Cuestionario de la actividad 3.2") ~ "U3_Cuestionario de la actividad 3.2",
      str_detect(.x, "Actividad 3.3.*Alineamiento") ~ "U3_Actividad 3.3 (Alineamiento estructural 2025)",
      str_detect(.x, "Cuestionario de la actividad 3.3") ~ "U3_Cuestionario de la actividad 3.3",
      str_detect(.x, "Cuestionario final proteínas") ~ "U3_Cuestionario final proteínas 2025",
      str_detect(.x, "ACTIVIDAD 4.1") ~ "U4_ACTIVIDAD 4.1 2025",
      str_detect(.x, "Evaluación Heatmap") ~ "U4_Evaluación Heatmap",
      str_detect(.x, "ACTIVIDAD 5.1 2025") ~ "U5_ACTIVIDAD 5.1 2025",
      str_detect(.x, "Cuestionario Evaluación U5") ~ "U5_Cuestionario Evaluación U5",
      .x %in% c("nombre", "id") ~ .x,  # Mantener nombre e id igual
      TRUE ~ .x  # Mantener otros nombres sin cambios
    )
  })

# 2. Eliminar columna no deseada
#datos_renombrados <- datos_renombrados %>%
#  select(-matches("Prueba de Diagnóstico Inicial"))

# 3. Definir el orden exacto deseado
orden_columnas <- c(
  "nombre", "id",
  "U1_Cuestionario de Recapitulación: Sesión 1",
  "U1_Ejercicio - Archivos fasta",
  "U1_Actividad 1.1 Bases de datos",
  "U1_Actividad 1.2 Formato fasta",
  "U1_TAREA PARA MARTES 18 de FEBRERO. Act. 1.2 Ejercicio Fasta-NCBI",
  "U1_Cuestionario de Recapitulación: Sesión 2",
  "U2_TAREA PARA MARTES 25 DE FEBRERO. Actividad 2.1",
  "U2_Actividad de sistemática molecular",
  "U2_Evaluación de sistemática molecular",
  "U2_Cuestionario final U2 - NGS y formatos",
  "U3_Cuestionario de actividades del 11 de marzo. PRESENTACIÓN + VIDEO",
  "U3_Actividad 3.1 (PDB 2025)",
  "U3_Cuestionario de la actividad 3.1",
  "U3_Actividad 3.2 (GFP 2025)",
  "U3_Cuestionario de la actividad 3.2",
  "U3_Actividad 3.3 (Alineamiento estructural 2025)",
  "U3_Cuestionario de la actividad 3.3",
  "U3_Cuestionario final proteínas 2025",
  "U4_ACTIVIDAD 4.1 2025",
  "U4_Evaluación Heatmap",
  "U5_ACTIVIDAD 5.1 2025",
  "U5_Cuestionario Evaluación U5"
)

# 4. Ordenar las columnas según el orden definido
datos_final <- datos_renombrados %>%
  select(any_of(orden_columnas))  # any_of() ignora columnas que no existan

# Columnas a convertir (ajusta según tus necesidades)
columnas_convertir <- c(
  "U1_Ejercicio - Archivos fasta",
  "U1_TAREA PARA MARTES 18 de FEBRERO. Act. 1.2 Ejercicio Fasta-NCBI"
)

# Conversión automática
datos_final <- datos_final %>%
  mutate(across(
    all_of(columnas_convertir),
    ~ ./10 %>% round(1)  # Divide entre 10 y redondea a 1 decimal
  ))

# 5. Verificar resultados
cat("Columnas finales:\n")
print(colnames(datos_final))




# 1. Convertir IDs a character y obtener listas únicas
ids_grupo2 <- grupo2 %>% 
  mutate(id_estudiante = as.character(`Número de ID`)) %>%
  pull(id_estudiante) %>% 
  unique()

ids_grupo3 <- grupo3 %>% 
  mutate(id_estudiante = as.character(`Número de ID`)) %>%
  pull(id_estudiante) %>% 
  unique()

# Asegurar que datos_final$id sea character (por si acaso)
datos_final <- datos_final %>% 
  mutate(id = as.character(id))

# 2. Separar los dataframes
datos_grupo2 <- datos_final %>% 
  filter(id %in% ids_grupo2) %>%
  mutate(grupo = "GRUPO 2")

datos_grupo3 <- datos_final %>% 
  filter(id %in% ids_grupo3) %>%
  mutate(grupo = "GRUPO 3")

# 3. Verificación adicional (detallada)
cat("\nVerificación detallada:\n")
cat("Total IDs únicos en Grupo 2:", length(ids_grupo2), "\n")
cat("Total IDs únicos en Grupo 3:", length(ids_grupo3), "\n")
cat("IDs en común entre grupos:", sum(ids_grupo2 %in% ids_grupo3), "\n\n")

cat("Muestra de IDs Grupo 2:", head(ids_grupo2), "\n")
cat("Muestra de IDs en datos_final:", head(datos_final$id), "\n")




##### Formarto para exportar en excel

library(openxlsx)
library(dplyr)

## 1. Función para preparar datos (NA -> 0) ----
preparar_para_formatos <- function(df) {
  numeric_cols <- names(df)[sapply(df, is.numeric) & !names(df) %in% c("id", "ID")]
  df %>%
    mutate(across(all_of(numeric_cols), ~ if_else(is.na(.), 0, .))) %>%
    mutate(Promedio = rowMeans(select(., all_of(numeric_cols)), na.rm = TRUE)) %>%
    mutate(Promedio = round(Promedio, 2))
}

## 2. Función para aplicar formatos ----
aplicar_formatos <- function(wb, sheet_name, df) {
  # Estilos
  style_red <- createStyle(fontColour = "#FFFFFF", bgFill = "#FF0000")   # Rojo para 0
  style_orange <- createStyle(fontColour = "#000000", bgFill = "#FFA500") # Naranja para 1-5
  
  # Identificar columnas numéricas (incluyendo "Promedio")
  numeric_cols <- names(df)[sapply(df, is.numeric)]
  
  for (col_name in numeric_cols) {
    col_num <- which(names(df) == col_name)
    col_letter <- int2col(col_num)  # Convertir número de columna a letra (A, B, C, etc.)
    
    # Para valores iguales a 0 (rojo)
    conditionalFormatting(
      wb, sheet_name,
      cols = col_num,
      rows = 2:(nrow(df) + 1),
      style = style_red,
      type = "expression",
      rule = paste0("$", col_letter, "2=0")
    )
    
    # Para valores entre 1 y 5 (naranja)
    conditionalFormatting(
      wb, sheet_name,
      cols = col_num,
      rows = 2:(nrow(df) + 1),
      style = style_orange,
      type = "expression",
      rule = paste0("AND($", col_letter, "2>=1, $", col_letter, "2<=5)")
    )
  }
}

# Función auxiliar para convertir números de columna a letras
int2col <- function(n) {
  if (n <= 26) return(LETTERS[n])
  paste0(LETTERS[(n - 1) %/% 26], LETTERS[(n - 1) %% 26 + 1])
}

cat("Validando datos antes de exportar...\n")
cat("Grupo 2:\n")
print(head(datos_grupo2))
cat("Grupo 3:\n")
print(head(datos_grupo3))


## 3. Exportar TODO en un solo archivo ----
exportar_todo_en_un_archivo <- function() {
  wb <- createWorkbook()
  
  # Hoja con datos originales del Grupo 2
  if(exists("datos_grupo2") && nrow(datos_grupo2) > 0) {
    addWorksheet(wb, "Grupo2_SinFormato")
    writeData(wb, "Grupo2_SinFormato", datos_grupo2)
    setColWidths(wb, "Grupo2_SinFormato", cols = 1:ncol(datos_grupo2), widths = "auto")
  }
  
  # Hoja con datos originales del Grupo 3
  if(exists("datos_grupo3") && nrow(datos_grupo3) > 0) {
    addWorksheet(wb, "Grupo3_SinFormato")
    writeData(wb, "Grupo3_SinFormato", datos_grupo3)
    setColWidths(wb, "Grupo3_SinFormato", cols = 1:ncol(datos_grupo3), widths = "auto")
  }
  
  # Hoja con datos formateados del Grupo 2
  if(exists("datos_grupo2") && nrow(datos_grupo2) > 0) {
    datos_g2_format <- preparar_para_formatos(datos_grupo2)
    addWorksheet(wb, "Grupo2_ConFormato")
    writeData(wb, "Grupo2_ConFormato", datos_g2_format)
    aplicar_formatos(wb, "Grupo2_ConFormato", datos_g2_format)
    setColWidths(wb, "Grupo2_ConFormato", cols = 1:ncol(datos_g2_format), widths = "auto")
  }
  
  # Hoja con datos formateados del Grupo 3
  if(exists("datos_grupo3") && nrow(datos_grupo3) > 0) {
    datos_g3_format <- preparar_para_formatos(datos_grupo3)
    addWorksheet(wb, "Grupo3_ConFormato")
    writeData(wb, "Grupo3_ConFormato", datos_g3_format)
    aplicar_formatos(wb, "Grupo3_ConFormato", datos_g3_format)
    setColWidths(wb, "Grupo3_ConFormato", cols = 1:ncol(datos_g3_format), widths = "auto")
  }
  
  # Guardar archivo
  if(!dir.exists("output")) dir.create("output")
  output_file <- "output/calificaciones_consolidado.xlsx"
  saveWorkbook(wb, output_file, overwrite = TRUE)
  message("Archivo exportado con éxito: ", normalizePath(output_file))
  
  # Mostrar resumen
  cat("\nHojas creadas:\n")
  if(exists("datos_grupo2")) cat("- Grupo2_SinFormato\n- Grupo2_ConFormato\n")
  if(exists("datos_grupo3")) cat("- Grupo3_SinFormato\n- Grupo3_ConFormato\n")
}

## 4. Ejecutar la exportación ----
exportar_todo_en_un_archivo()
