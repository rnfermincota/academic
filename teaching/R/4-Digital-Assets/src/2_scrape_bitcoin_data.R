
#--------------------------------------------------------------------------------
#--------------------------------------------------------------------------------
# Scrape Bitcoin Data
#--------------------------------------------------------------------------------
#--------------------------------------------------------------------------------

## Quandl Functions
#--------------------------------------------------------------------------------

# The quandl_tidy function is a wrapper around the Quandl function that returns a 
# cleaner tibble. Namely, we only want to get the respective `code`, its `date` 
# and `value`.

Quandl.api_key("D6jFSzVwLvbEE_6LvyhE") 

quandl_tidy <- function(code, name) { 
  df <- Quandl(code) %>% 
    mutate(code = code, name = name) %>% 
    rename(date = Date, value = Value) %>% 
    arrange(date) %>% 
    as_tibble()
  return(df)
}

#--------------------------------------------------------------------------------
## Bitcoin Exchange Rate Data
#--------------------------------------------------------------------------------

# We will use the quandl_tidy function to obtain USD/Bitcoin exchange data from Bitstamp, 
# a trading platform for Bitcoin.

if (!refresh_flag){
  load(file.path(path_data, "bitcoin_price.RData")) #
} else {
  bitcoin_price <- Quandl("BCHARTS/BITSTAMPUSD") %>%
    arrange(Date) %>%
    as_tibble()
  colnames(bitcoin_price) <- c(
    "date", "open", "high", "low", 
    "close", "volume_btc", 
    "volume_currency", "weighted_price"
  )
  save(bitcoin_price, file=file.path(path_data, "bitcoin_price.RData"))
}

#--------------------------------------------------------------------------------
## Bitcoin Indicators
#--------------------------------------------------------------------------------

# This section allows us to pull relevant BTC information with `quandl_tidy`. 
# Relevant features such as BTC Market Capitalisation, Hash Rate and BTC Days 
# Destroyed that measures the transaction volume of BTC.


if (!refresh_flag){
  load(file.path(path_data, "bitcoin_data.RData")) #
} else {
  code_list <- list(
    c("BCHAIN/TOTBC", "Total Bitcoins"), 
    c("BCHAIN/MKTCP", "Bitcoin Market Capitalization"), 
    c("BCHAIN/NADDU", "Bitcoin Number of Unique Addresses Used"), 
    c("BCHAIN/ETRAV", "Bitcoin Estimated Transaction Volume BTC"), 
    c("BCHAIN/ETRVU", "Bitcoin Estimated Transaction Volume USD"), 
    c("BCHAIN/TRVOU", "Bitcoin USD Exchange Trade Volume"), 
    c("BCHAIN/NTRAN", "Bitcoin Number of Transactions"), 
    c("BCHAIN/NTRAT", "Bitcoin Total Number of Transactions"), 
    c("BCHAIN/NTREP", "Bitcoin Number of Transactions Excluding Popular Addresses"), 
    c("BCHAIN/NTRBL", "Bitcoin Number of Tansaction per Block"), 
    c("BCHAIN/ATRCT", "Bitcoin Median Transaction Confirmation Time"), 
    c("BCHAIN/TRFEE", "Bitcoin Total Transaction Fees"), 
    c("BCHAIN/TRFUS", "Bitcoin Total Transaction Fees USD"), 
    c("BCHAIN/CPTRA", "Bitcoin Cost Per Transaction"), 
    c("BCHAIN/CPTRV", "Bitcoin Cost % of Transaction Volume"), 
    c("BCHAIN/BLCHS", "Bitcoin api.blockchain Size"), 
    c("BCHAIN/AVBLS", "Bitcoin Average Block Size"), 
    c("BCHAIN/TOUTV", "Bitcoin Total Output Volume"), 
    c("BCHAIN/HRATE", "Bitcoin Hash Rate"), 
    c("BCHAIN/BCDDE", "Bitcoin Days Destroyed"), 
    c("BCHAIN/BCDDW", "Bitcoin Days Destroyed Minimum Age 1 Week"), 
    c("BCHAIN/BCDDM", "Bitcoin Days Destroyed Minimum Age 1 Month"), 
    c("BCHAIN/BCDDY", "Bitcoin Days Destroyed Minimum Age 1 Year") ,
    c("BCHAIN/BCDDC", "Bitcoin Days Destroyed Cumulative")
  )
  
  bitcoin_data <- tibble()
  
  for (i in seq_along(code_list)) { 
    print(str_c("Downloading data for ", code_list[[i]][1], "."))
    bitcoin_data <- bind_rows(
      bitcoin_data, 
      quandl_tidy(
        code_list[[i]][1], 
        code_list[[i]][2])
      )
  }
  
  save(bitcoin_data, file=file.path(path_data, "bitcoin_data.RData"))
}


bitcoin_data <- bitcoin_data %>%
  select(-name) %>%
  spread(code, value)

colnames(bitcoin_data) <- make.names(colnames(bitcoin_data))
