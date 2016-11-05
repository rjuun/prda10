# 3 Requirements
# Create a Excel with a worksheet which suits as a place holder for storage and graphical repre- sentation of the BTC / USD price history.
# Create a worksheet which acts as a configuration area. The user should be able to change the following settings:
#   – Number of days of the BTC/EUR price history (e.g. 365 days)
#   – Number of days used for the moving average (e.g. 21 days)
#   – Numbe rof days used for the Relative Strength Index (e.g. 21 days)
#  Create a VBA procedure, which is able to download the price chart from the website http: //www.quandl.com in a specific format (e.g. CSV). The output should be stored in the already mentioned worksheet.
#  A VBA procedure should create separate column which shows the Moving Average (based on the configuration settings mentioned above)
#  A VBA procedure should create separate column which shows the Relative Strength Index (based on the configuration settings mentioned above)
#  Based on the three time series a chart should be generated with VBA. Please make sure of using a different axis for the RSI indicator.

# install the package
install.packages("Quandl")

# load the library and pass on the api_key
library(Quandl)
Quandl.api_key("P6QLMV4K3VrkT3jsdCP1")

# Variables
# set system date so that we can download only the last 365 days
startDate <- as.Date(Sys.Date() - 365)
k_set <- 21

# downloads bitcoin dataset and stores in a dataframe
bitcoin <- Quandl("BITCOIN/BTCEUSD", meta = TRUE, start_date = "1-1-2014")
bitcoin <- arrange(bitcoin, bitcoin$Date)

# rollmean function in R calculates the moving average automatically, but a number of
# definitions must follow:
# number of days for the moving average is defined by k_set
# since a number of observations are used to calculate the first moving average point
# it is necessary to calculate it apart, and add as parametr for the "fill"
# Also, 
first_avg <- mean(bitcoin$Close[1:k_set])
mov_average <- rollmean(bitcoin$Close, k = k_set, fill = NA, align = "left")

# once the moving average series is calculated, it must be binded together to the
# original dataframe
bitcoin <- cbind(bitcoin, mov_average)

# loads Technical Analysis package for R
library("TTR")
rel_strength <- RSI(bitcoin$Close, n = k_set)
bitcoin <- cbind(bitcoin, rel_strength)

# loads ggplot library to run graphics in R
library(ggplot2)

# 
ggplot(bitcoin, aes(Date, y = Price, color = variable)) + 
  geom_point(aes(y = Close, col = "Close Price")) + 
  geom_line(aes(y = mov_average, col = "Mov Average")) +
  geom_line(aes(y = rel_strength, col = "RSI"))