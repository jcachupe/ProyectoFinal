# Instalar y cargar paquetes necesarios
install.packages("readxl")
install.packages("ggplot2")
install.packages("forecast")
install.packages("dplyr")
install.packages("openxlsx")
install.packages("lubridate")
install.packages("tidyr")
install.packages("scales")


# Cargar paquetes
library(readxl)
library(ggplot2)
library(forecast)
library(dplyr)
library(openxlsx)
library(lubridate)
library(tidyr)
library(scales)

# Fijar el directorio de trabajo
setwd("C:/Users/ccachupe/Documents/BOOTCAMP/ProyectoR")

# Importar el dataset desde el archivo Excel subido en Posit Cloud
df <- read_excel("DataSet_ProductosR2 1.xlsx")

# Ver las primeras filas del dataset
head(df)

# Ver el dataframe completo en una ventana interactiva
View(df)

# Resumen de los datos
summary(df)

data_frame <- df %>%
  select(-Fecha_Apertura) %>%  # Eliminar la columna Fecha_Apertura
  group_by(Fecha, Producto) %>%  # Agrupar por Fecha y Producto
  summarise(
    Ventas = sum(Ventas, na.rm = TRUE),
    Intereses = sum(Intereses, na.rm = TRUE),
    Comisiones = sum(Comisiones, na.rm = TRUE)
  )

  # Ver el dataframe agrupado
head(data_frame)
View(data_frame)

# Convertir la columna de fecha a tipo Date
data_frame$Fecha <- as.Date(data_frame$Fecha, format="%Y-%m-%d")

# Verificar si la conversión de fechas fue exitosa
str(data_frame)

# Análisis Exploratorio de Datos
ggplot(data_frame, aes(x = Fecha, y = Ventas, color = Producto)) +
  geom_line() +
  labs(title = "Ventas Totales por Fecha", x = "Fecha", y = "Ventas")+
  scale_y_continuous(labels = scales::label_number(scale = 1e-6, suffix = "M")) +
  theme_minimal()

ggplot(data_frame, aes(x = Fecha, y = Ventas, color = Producto)) +
  geom_line() +
  facet_wrap(~ Producto) +
  labs(title = "Ventas por Producto", x = "Fecha", y = "Ventas")+
  scale_y_continuous(labels = scales::label_number(scale = 1e-6, suffix = "M")) +
  theme_minimal()

ggplot(data_frame, aes(x = Fecha, y = Ventas)) +
  geom_line() +
  facet_wrap(~ Producto, scales = "free_y") +  # Separar por producto
  labs(title = "Ventas por Producto", x = "Fecha", y = "Ventas") +
  scale_y_continuous(labels = scales::label_number(scale = 1e-6, suffix = "M")) +
  theme_minimal()

# ==================================================================================

# Crear dataframe para proyecciones
proyecciones <- data.frame(Fecha = seq.Date(from = as.Date("2023-01-01"), to = as.Date("2025-12-31"), by = "month"))

# Modelo Predictivo - Regresión Lineal
productos <- unique(data_frame$Producto)
for (producto in productos) {
  producto_data <- data_frame %>% filter(Producto == producto)
  modelo <- lm(Ventas ~ Fecha, data = producto_data)
  pred <- predict(modelo, newdata = data.frame(Fecha = proyecciones$Fecha))
  proyecciones[[producto]] <- pred
}

View(proyecciones)

# Modelo Predictivo - Series de Tiempo
proyecciones_arima <- data.frame(Fecha = proyecciones$Fecha)
for (producto in productos) {
  producto_data <- data_frame %>% filter(Producto == producto)
  ts_data <- ts(producto_data$Ventas, start = c(year(min(producto_data$Fecha)), month(min(producto_data$Fecha))), frequency = 12)
  fit <- auto.arima(ts_data)
  pred_arima <- forecast(fit, h = length(proyecciones$Fecha))
  proyecciones_arima[[producto]] <- as.numeric(pred_arima$mean)
}

View(proyecciones_arima)

# Visualización de Proyecciones
historico_proyeccion <- data_frame %>%
  mutate(Tipo = "Histórico") %>%
  bind_rows(
    proyecciones %>%
      mutate(Tipo = "Regresión Lineal", Fecha = as.Date(Fecha)) %>%
      pivot_longer(-c(Fecha, Tipo), names_to = "Producto", values_to = "Ventas")
  ) %>%
  bind_rows(
    proyecciones_arima %>%
      mutate(Tipo = "ARIMA", Fecha = as.Date(Fecha)) %>%
      pivot_longer(-c(Fecha, Tipo), names_to = "Producto", values_to = "Ventas")
  )

View(historico_proyeccion)

# Graficar las Proyecciones
ggplot(historico_proyeccion, aes(x = Fecha, y = Ventas, color = Tipo)) +
  geom_line() +
  facet_wrap(~ Producto, scales = "free_y") +
  labs(title = "Proyección de Ventas por Producto", x = "Fecha", y = "Ventas")+
  scale_y_continuous(labels = scales::label_number(scale = 1e-6, suffix = "M")) +
  theme_minimal()

# =============================================================================
wb <- createWorkbook()
addWorksheet(wb, "Datos Históricos")
writeData(wb, "Datos Históricos", data_frame)
addWorksheet(wb, "Proyecciones Regresión Lineal")
writeData(wb, "Proyecciones Regresión Lineal", proyecciones)
addWorksheet(wb, "Proyecciones ARIMA")
writeData(wb, "Proyecciones ARIMA", proyecciones_arima)
saveWorkbook(wb, "Proyecciones_Ventas.xlsx", overwrite = TRUE)

# Confirmar la ruta del archivo guardado
ruta_actual <- getwd()
archivo_excel <- file.path(ruta_actual, "Proyecciones_Ventas.xlsx")
cat("El archivo Excel se ha guardado en:", archivo_excel, "\n")
