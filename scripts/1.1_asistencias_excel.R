# Instalar paquetes si no los tienes (solo una vez)
# install.packages("readxl")
# install.packages("dplyr")
# install.packages("stringr")

# Cargar los paquetes necesarios
library(readxl)
library(dplyr)
library(stringr)
library(openxlsx)
library(ggplot2)
# Cargar el archivo Excel saltando las primeras 2 filas
datos <- read_excel("data/2264_BIOINFV2B_Asistencias_20250419-2054.xlsx", 
                    skip = 2)  # <- Esta opción omite las primeras 2 filas

# Eliminar las últimas 7 columnas
num_columnas <- ncol(datos)
datos <- datos[, 1:(num_columnas - 7)]


# Función para limpiar los valores de asistencia
limpiar_asistencia <- function(x) {
  case_when(
    str_detect(x, "^R \\(\\d+/\\d+\\)") ~ "R",
    str_detect(x, "^P \\(\\d+/\\d+\\)") ~ "P",
    str_detect(x, "^FI \\(\\d+/\\d+\\)") ~ "FI",
    str_detect(x, "^FJ \\(\\d+/\\d+\\)") ~ "FJ",
    TRUE ~ x  # Mantener otros valores como están
  )
}

# Aplicar la función de limpieza desde la columna 8 hasta la última
datos_limpios <- datos %>%
  mutate(across(8:ncol(.), limpiar_asistencia))

# Verificar los resultados
head(datos_limpios)


# Separación directa por grupos (versión más simple)
grupo2 <- datos_limpios %>% 
  filter(`Grupos` == "GRUPO 2")  # Usando el nombre exacto de la columna

grupo3 <- datos_limpios %>% 
  filter(`Grupos` == "GRUPO 3")


# DF GRUPO 2. Identificar y eliminar columnas que mencionan "GRUPO 3" en su nombre
columnas_a_eliminar <- names(grupo2)[str_detect(names(grupo2), "GRUPO 3")]

if(length(columnas_a_eliminar) > 0) {
  message("Eliminando las siguientes columnas de GRUPO 2: ", 
          paste(columnas_a_eliminar, collapse = ", "))
  
  grupo2 <- grupo2 %>%
    select(-all_of(columnas_a_eliminar))
} else {
  message("No se encontraron columnas con 'GRUPO 3' en sus nombres")
}

# Hacer lo mismo para GRUPO 3 (por consistencia)
# Eliminar columnas que mencionan "GRUPO 2"
columnas_a_eliminar_g3 <- names(grupo3)[str_detect(names(grupo3), "GRUPO 2")]

if(length(columnas_a_eliminar_g3) > 0) {
  message("Eliminando las siguientes columnas de GRUPO 3: ", 
          paste(columnas_a_eliminar_g3, collapse = ", "))
  
  grupo3 <- grupo3 %>%
    select(-all_of(columnas_a_eliminar_g3))
}

# Verificación final
message("\nEstructura de GRUPO 2 después de limpieza:")
glimpse(grupo2)

message("\nEstructura de GRUPO 3 después de limpieza:")
glimpse(grupo3)


## Reemplazar "?" por "FI" en ambos data frames ----
grupo2 <- grupo2 %>%
  mutate(across(everything(), ~ifelse(. == "?", "FI", .)))

grupo3 <- grupo3 %>%
  mutate(across(everything(), ~ifelse(. == "?", "FI", .)))
         
## Agregar formato condicional con colores ----
         
# Función para crear estilos
estilos <- createStyle(
           fontColour = "white",
           textDecoration = "bold")
         
estilo_FI <- createStyle(
           bgFill = "#FF6B6B",  # Rojo pastel
           fontColour = "white")
         
estilo_R <- createStyle(
           bgFill = "#4ECDC4",  # Verde azulado pastel
           fontColour = "white")
         
## Preparar el archivo Excel ----
wb <- createWorkbook()
         
# Función para añadir hojas con formato
agregar_hoja_con_formato <- function(wb, df, nombre_hoja) {
           addWorksheet(wb, nombre_hoja)
           writeData(wb, sheet = nombre_hoja, df)
           
           # Aplicar formato condicional
           for(col in 1:ncol(df)) {
             # Para valores FI
             conditionalFormatting(wb, nombre_hoja,
                                   cols = col,
                                   rows = 2:(nrow(df)+1),
                                   style = estilo_FI,
                                   rule = '=="FI"')
             
             # Para valores R
             conditionalFormatting(wb, nombre_hoja,
                                   cols = col,
                                   rows = 2:(nrow(df)+1),
                                   style = estilo_R,
                                   rule = '=="R"')
           }
           
# Autoajustar columnas
           setColWidths(wb, nombre_hoja, cols = 1:ncol(df), widths = "auto")
         }
         
# Añadir hojas
agregar_hoja_con_formato(wb, grupo2, "GRUPO_2")
agregar_hoja_con_formato(wb, grupo3, "GRUPO_3")
         
## Guardar el archivo Excel ----
saveWorkbook(wb, "output/Asistencias_por_Grupo.xlsx", overwrite = TRUE)
         
message("Archivo Excel exportado exitosamente con dos hojas: GRUPO_2 y GRUPO_3")

