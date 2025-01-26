
#  title: "EM7136 Forecasting Techniques (2024-2025 Fall) Final Project"
#  author: "Hajar Faiyad 924824912"
#  date: "2025-01-14"

# Packages used: fpp2 for forecasting methods and writexl to export data frames to excel files.
# Data: split into two main excel files, one for training and one for test.
library(fpp2)
library(writexl)

# Plotting the Monthly Data to understand the sequence
# Plotting the Autocorrelation Function to find any trend or seasonality
plot.ts(Train_Data$Y13)
acf(Train_Data$Y13, lag.max = 48)

# From ACF and Time Series Plot we can infer:
#    Monthly time series data has trend and additive seasonality (s=12).

# Assign the time series values to a variable Y13
Y13 <- ts(Train_Data$Y13, frequency = 12)

#     ********* Applying 5 Forecasting Methods ********* 

# 1. Holt-Winters' Additive Method
holt_winters_add <- hw(Y13, seasonal = "additive", h = 12)
summary(holt_winters_add)
holt_winters_add$mean

accuracy_holt_winters_add <- accuracy(holt_winters_add$mean, Test_Data$TestY13)
accuracy_holt_winters_add


# 2. Holt-Winters' Multiplicative Method
holt_winters_mul <- hw(Y13, seasonal = "multiplicative", h = 12)
summary(holt_winters_mul)
holt_winters_mul$mean

accuracy_holt_winters_mul <- accuracy(holt_winters_mul$mean, Test_Data$TestY13)
accuracy_holt_winters_mul


# 3. Additive Classical Decomposition
decom_add <- decompose(Y13, type = "additive")
autoplot(decom_add)

forecast_trend_decom_add <- holt(seasadj(decom_add), h=12)
forecast_trend_decom_add$mean
decom_add$figure

#autoplot(forecast_trend_decom_add)

forecast_decom_add <- 0 + decom_add$figure + forecast_trend_decom_add$mean
forecast_decom_add

#autoplot(forecast_decom_add)

accuracy_decom_add <- accuracy(forecast_decom_add, Test_Data$TestY13)
accuracy_decom_add

# 4. Multiplicative Classical Decomposition
decom_mul <- decompose(Y13, type = "multiplicative")
autoplot(decom_mul)

forecast_trend_decom_mul <- holt(seasadj(decom_mul), h=12)
forecast_trend_decom_mul$mean
decom_mul$figure

#autoplot(forecast_trend_decom_mul)

forecast_decom_mul <- 1 * decom_mul$figure * forecast_trend_decom_mul$mean
forecast_decom_mul

#autoplot(forecast_decom_mul)

accuracy_decom_mul <- accuracy(forecast_decom_mul, Test_Data$TestY13)
accuracy_decom_mul

# 5. Regression
Y13_reg <- Train_Data_Reg$Y13
t <- Train_Data_Reg$t
tsq <- Train_Data_Reg$tsq

# using factor() to introduce categorical variables
D1 <- factor(Train_Data_Reg$D1)
D2 <- factor(Train_Data_Reg$D2)
D3 <- factor(Train_Data_Reg$D3)
D4 <- factor(Train_Data_Reg$D4)
D5 <- factor(Train_Data_Reg$D5)
D6 <- factor(Train_Data_Reg$D6)
D7 <- factor(Train_Data_Reg$D7)
D8 <- factor(Train_Data_Reg$D8)
D9 <- factor(Train_Data_Reg$D9)
D10 <- factor(Train_Data_Reg$D10)
D11 <- factor(Train_Data_Reg$D11)

# fitting linear variables
reg_model <- lm(Y13_reg~t+tsq+D1+D2+D3+D4+D5+D6+D7+D8+D9+D10+D11, data = Train_Data_Reg)
summary(reg_model)
accuracy(reg_model)

checkresiduals(reg_model)

# normality check
ks.test(reg_model$residuals, "pnorm", mean(reg_model$residuals), sd(reg_model$residuals))
shapiro.test(reg_model$residuals)
plot(reg_model, 1)

forecast_reg <- predict(reg_model, newdata = Train_Data_Reg_Forecast)
forecast_reg

accuracy_reg <- accuracy(forecast_reg, Test_Data$TestY13)
accuracy_reg


#     ********* Gathering All Forecast Values and Accuracy Measures and Saving Them to Excel Files *********

# Forecast
Method <- c("Holt-Winters Additive", "Holt-Winters Multiplicative", "Additive Decomposition", "Multiplicative Decomposition", "Regression")
M1 <- c(holt_winters_add$mean[1], holt_winters_mul$mean[1], forecast_decom_add[1], forecast_decom_mul[1], forecast_reg[1])
M2 <- c(holt_winters_add$mean[2], holt_winters_mul$mean[2], forecast_decom_add[2], forecast_decom_mul[2], forecast_reg[2])
M3 <- c(holt_winters_add$mean[3], holt_winters_mul$mean[3], forecast_decom_add[3], forecast_decom_mul[3], forecast_reg[3])
M4 <- c(holt_winters_add$mean[4], holt_winters_mul$mean[4], forecast_decom_add[4], forecast_decom_mul[4], forecast_reg[4])
M5 <- c(holt_winters_add$mean[5], holt_winters_mul$mean[5], forecast_decom_add[5], forecast_decom_mul[5], forecast_reg[5])
M6 <- c(holt_winters_add$mean[6], holt_winters_mul$mean[6], forecast_decom_add[6], forecast_decom_mul[6], forecast_reg[6])
M7 <- c(holt_winters_add$mean[7], holt_winters_mul$mean[7], forecast_decom_add[7], forecast_decom_mul[7], forecast_reg[7])
M8 <- c(holt_winters_add$mean[8], holt_winters_mul$mean[8], forecast_decom_add[8], forecast_decom_mul[8], forecast_reg[8])
M9 <- c(holt_winters_add$mean[9], holt_winters_mul$mean[9], forecast_decom_add[9], forecast_decom_mul[9], forecast_reg[9])
M10 <- c(holt_winters_add$mean[10], holt_winters_mul$mean[10], forecast_decom_add[10], forecast_decom_mul[10], forecast_reg[10])
M11 <- c(holt_winters_add$mean[11], holt_winters_mul$mean[11], forecast_decom_add[11], forecast_decom_mul[11], forecast_reg[11])
M12 <- c(holt_winters_add$mean[12], holt_winters_mul$mean[12], forecast_decom_add[12], forecast_decom_mul[12], forecast_reg[12])

df1 <- data.frame(Method, M1, M2, M3, M4, M5, M6, M7, M8, M9, M10, M11, M12)
df1
write_xlsx(df1, "C:\\Hajar\\Masters Degree (Marmara Uni)\\Semester 1\\Forecasting Techniques\\Final Project\\Submission\\All_Forecasts.xlsx")


# Accuracy Measures 
RMSE <- c(accuracy_holt_winters_add[2], accuracy_holt_winters_mul[2], accuracy_decom_add[2], accuracy_decom_mul[2], accuracy_reg[2])
MAE <- c(accuracy_holt_winters_add[3], accuracy_holt_winters_mul[3], accuracy_decom_add[3], accuracy_decom_mul[3], accuracy_reg[3])
MAPE <- c(accuracy_holt_winters_add[5], accuracy_holt_winters_mul[5], accuracy_decom_add[5], accuracy_decom_mul[5], accuracy_reg[5])

df2 <- data.frame(Method, RMSE, MAE, MAPE)
df2
write_xlsx(df2, "C:\\Hajar\\Masters Degree (Marmara Uni)\\Semester 1\\Forecasting Techniques\\Final Project\\Submission\\Accuracy_Measures.xlsx")


# In conclusion, Multiplicative Decomposition method is the best model to forecast this data.

