airquality=read.xlsx("G:/Mi unidad/Docencia/2025-2026/Master PGPE/Plantillas/serie_temporal_plantilla.xlsx")

# Imputacion datos faltantes
library(openxlsx)
library(imputeTS)
plot.ts(airquality$Temp)
plantilla=loadWorkbook(file="G:/Mi unidad/Docencia/2025-2026/Master PGPE/Plantillas/serie_temporal_plantilla.xlsx")
Temp=na_mean(airquality$Temp)
plot.ts(Temp)
airquality$Temp=Temp
writeData(wb=plantilla, sheet=1, x=airquality, startCol=1, startRow=1)
saveWorkbook(wb=plantilla, file = "G:/Mi unidad/Docencia/2025-2026/Master PGPE/Plantillas/serie_temporal_plantilla_edit.xlsx", overwrite = TRUE)



# Prediccion
attach(airquality)
arima_wind=auto.arima(Wind)
arima_temp=auto.arima(Temp)
predicts_wind=predict(arima_wind,3)
predicts_temp=predict(arima_temp,3)
