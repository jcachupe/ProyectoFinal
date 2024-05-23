# Instalar y cargar paquetes necesarios
install.packages("readxl")
install.packages("ggplot2")
install.packages("forecast")
install.packages("dplyr")
install.packages("openxlsx")
install.packages("lubridate")

# Cargar paquetes
library(readxl)
library(ggplot2)
library(forecast)
library(dplyr)
library(openxlsx)
library(lubridate)

# Importar el dataset desde el archivo Excel subido en Posit Cloud
data_frame <- read_excel("DataSet_ProductosR.xlsx")

# Verificar si el archivo se cargó correctamente
if (!exists("data_frame")) {
  stop("El archivo no se ha cargado correctamente.")
}

# Verificar la estructura del objeto 'data_frame'
str(data_frame)

# Verificar la clase del objeto 'data_frame'
class(data_frame)

# Asegurarse de que 'data_frame' es un dataframe
if (!is.data.frame(data_frame)) {
  stop("El objeto 'data_frame' no es un dataframe.")
}

# Ver las primeras filas del dataset
head(data_frame)

# Resumen de los datos
summary(data_frame)

# Convertir la columna de fecha a tipo Date
data_frame$Fecha <- as.Date(data_frame$Fecha, format="%Y-%m-%d")

# Verificar si la conversión de fechas fue exitosa
str(data_frame)

# Análisis Exploratorio de Datos
ggplot(data_frame, aes(x = Fecha, y = Ventas, color = Producto)) +
  geom_line() +
  labs(title = "Ventas Totales por Fecha", x = "Fecha", y = "Ventas")

ggplot(data_frame, aes(x = Fecha, y = Ventas, color = Producto)) +
  geom_line() +
  facet_wrap(~ Producto) +
  labs(title = "Ventas por Producto", x = "Fecha", y = "Ventas")

# Modelo Predictivo - Regresión Lineal
proyecciones <- data.frame(Fecha = seq.Date(from = as.Date("2023-01-01"), to = as.Date("2025-12-31"), by = "month"))

productos <- unique(data_frame$Producto)
for (producto in productos) {
  producto_data <- data_frame %>% filter(Producto == producto)
  modelo <- lm(Ventas ~ Fecha, data = producto_data)
  pred <- predict(modelo, newdata = data.frame(Fecha = proyecciones$Fecha))
  proyecciones[[producto]] <- pred
}

# Modelo Predictivo - Series de Tiempo
proyecciones_arima <- proyecciones
for (producto in productos) {
  producto_data <- data_frame %>% filter(Producto == producto)
  ts_data <- ts(producto_data$Ventas, start = c(year(min(producto_data$Fecha)), month(min(producto_data$Fecha))), frequency = 12)
  fit <- auto.arima(ts_data)
  pred_arima <- forecast(fit, h = 36)
  proyecciones_arima[[producto]] <- c(rep(NA, length(ts_data)), pred_arima$mean)
}

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

# Graficar las Proyecciones
ggplot(historico_proyeccion, aes(x = Fecha, y = Ventas, color = Tipo)) +
  geom_line() +
  facet_wrap(~ Producto) +
  labs(title = "Proyección de Ventas por Producto", x = "Fecha", y = "Ventas") +
  theme_minimal()

# Análisis Estadístico
resultados_regresion <- lapply(productos, function(producto) {
  producto_data <- data_frame %>% filter(Producto == producto)
  modelo <- lm(Ventas ~ Fecha, data = producto_data)
  summary(modelo)
})

resultados_arima <- lapply(productos, function(producto) {
  producto_data <- data_frame %>% filter(Producto == producto)
  ts_data <- ts(producto_data$Ventas, start = c(year(min(producto_data$Fecha)), month(min(producto_data$Fecha))), frequency = 12)
  fit <- auto.arima(ts_data)
  summary(fit)
})

# Exportar Resultados a Excel
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
