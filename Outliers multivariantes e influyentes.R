# Carga de datos y ajuste de modelo de regresion lineal multiple
data("stackloss")
X <- stackloss[, c("Air.Flow", "Water.Temp", "Acid.Conc.")]
modelo_rlm <- lm(stack.loss ~ Air.Flow + Water.Temp + Acid.Conc., data = stackloss)
summary(modelo_rlm)

# Distancia de Mahalanobis
md_distancia <- mahalanobis(
  x = X,
  center = colMeans(X),
  cov = cov(X)
)

# Umbral chi cuadrado
p <- ncol(X)
alfa <- 0.025 
umbral_chi2 <- qchisq(1 - alfa, df = p)

# Identificacion de outliers multivariantes
stackloss$md_distancia <- md_distancia
stackloss$Outlier_MD <- ifelse(stackloss$md_distancia > umbral_chi2, "SI", "NO")

# Distancia de Cook
distancia_cook <- cooks.distance(modelo_rlm)
stackloss$Cooks_D <- distancia_cook
n <- nrow(stackloss) 
umbral_cook <- 4 / n
stackloss$Influyente <- ifelse(stackloss$Cooks_D > umbral_cook, "SI", "NO")

# Generar Excel

library(openxlsx)
libro <- createWorkbook() 
class(libro) 
addWorksheet(wb=libro, sheetName="OutliersInf")
writeData(wb=libro, sheet="OutliersInf", x=stackloss, startCol=1, startRow=1)
conditionalFormatting(wb=libro, sheet="OutliersInf", cols=5, rows=2:(nrow(stackloss)+1), style=c("blue","white","red"), type="colourScale")
conditionalFormatting(wb=libro, sheet="OutliersInf", cols=7, rows=2:(nrow(stackloss)+1), style=c("blue","white","red"), type="colourScale")
saveWorkbook(wb=libro, file="G:/Mi unidad/Docencia/2025-2026/Master PGPE/Data/stackloss.xlsx", overwrite = TRUE)
