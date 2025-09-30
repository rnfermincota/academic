#--------------------------------------------------------------------------------
#--------------------------------------------------------------------------------
# Feature Engineering
#--------------------------------------------------------------------------------
#--------------------------------------------------------------------------------
# Using the `tidymodel` package, the framework for machine learning has 
# been made easier. Removing or adding features only requires **2 steps**:
# -   Loading data of the selected feature\
# -   Joining to the rest of the features by `date`

# Thus, `Section 3` is non-exhaustive, and features can be freely added and 
# removed. In choosing which feature to include, we mostly included technical 
# indicators and metrics we believe influence BTC's prices.

#--------------------------------------------------------------------------------
## Cleaning Bitcoin Prices
#--------------------------------------------------------------------------------

# Data errors are cleaned by using `na.locf` (last observations carried forward).
bitcoin_price[bitcoin_price == 0] <- NA
bitcoin_price <- bitcoin_price %>% map_df(na.locf)


#--------------------------------------------------------------------------------
## Define Target
#--------------------------------------------------------------------------------
# Our target will be using a `binary classifcation`, where if the future return
# is positive, it will be indicated with `1`. If the future return is negative, 
# it will be indicated with a `0`. `tq_mutate()` introduces an easy method to 
# calculate periodic returns.

bitcoin_model <- bitcoin_price %>%
  tq_mutate(
    select = close,
    mutate_fun = periodReturn,
    period = 'daily',
    type = 'arithmetic',
    col_rename = 'future_return'
  ) %>%
  mutate(
    future_return_sign = as.factor(ifelse(future_return > 0, 1, 0)),
    close = lag(close, 1),
    date = date - days(1)
  ) %>%
  select(date, close, future_return, future_return_sign) 

bitcoin_model <- bitcoin_model[-1, ]


#--------------------------------------------------------------------------------
## VIX
#--------------------------------------------------------------------------------

# **VIX** is a ticker symbol that represents the CBOE Volatility Index which 
# measures the stock market's expectation of volatility based on S&P 500 options. 
# Given that BTC is highly volatile, and given its **growing perception** as 
# digital gold, we want to investigate if there is any relationship between the 
# fear index of traditional markets to BTC.

get_yahoo <- function(ticker) {
  df <- getSymbols(
    ticker, 
    src = 'yahoo', 
    auto.assign = FALSE, 
    from = '1900-01-01'
  )
  df <- df %>%
    as_tibble() %>%
    mutate(date = index(df))
  colnames(df) <- c(
    "open", "high", "low", "close", 
    "volume", "adjusted_close", 
    "date", "ticker"
  )
  return(df)
}

if (!refresh_flag){
  load(file.path(path_data, "vix.RData")) #
} else {
  vix <- get_yahoo('^VIX') %>%
    select(date, vix = adjusted_close) %>% 
    as_tibble() %>% 
    map_df(na.locf0)
  save(vix, file=file.path(path_data, "vix.RData"))
}

bitcoin_model <- bitcoin_model %>% 
  left_join(vix) %>%
  map_df(na.locf)


#--------------------------------------------------------------------------------
## Bitcoin Price Change
#--------------------------------------------------------------------------------

# We are trying to simulate the rate of change of prices by finding the price 
# change of Bitcoin over several days. We have set the maximum change to be 90 
# days as we believe that momentum only has a 90 day effecive period.


price_change <- function(
  close, n = 1, 
  type = change_type, Ticker = NA
){
  
  close_change <-  matrix(
    nrow = length(close), 
    ncol = n
  ) %>% 
    as.data.frame()
  
  for(i in 1:n){
    close_change[,i] <- ROC(close, n = i, type = type)
    if(is.na(Ticker)){
      names(close_change)[i] <- paste0(
        'close_change_', type, "_", 
        as.character(i), 'd'
      )
    }else{
      names(close_change)[i] <- paste0(
        Ticker, '_close_change_', 
        type, "_", as.character(i), 'd'
      )
    }
  }
  return(close_change)
}

# Continuous price change
close_change_continuous <- price_change(
  bitcoin_model$close, 
  n = 90, 
  type = "continuous"
)

# Discrete price change
close_change_discrete <- price_change(
  bitcoin_model$close, 
  n = 90, 
  type = "discrete"
)

bitcoin_model <- bitcoin_model %>%
  cbind(close_change_continuous) %>%
  cbind(close_change_discrete)




#--------------------------------------------------------------------------------
## Momentum of Price Movement
#--------------------------------------------------------------------------------

# Based on the rate of change from the close price and the price one day before,
# we are then able to get the momentum of the price movements.

# this is finding the derivative of the ROC above
momentum <- function(
  days_lag, 
  type = change_type
){
  
  col <- paste0(
    "close_change_", 
    type, 
    "_", 
    as.character(days_lag),
    "d"
  )
  
  df <- bitcoin_model %>%
    select("close", col) %>% 
    mutate(lag_change = lag(bitcoin_model[[col]])) %>%
    mutate(momentum = (bitcoin_model[[col]] - lag_change) / days_lag)
  
  return(df$momentum)
  
}

v <- 1 

# Generate momentum for continuous ROC
momentum_continuous <- momentum(
  days_lag = v, 
  type = "continuous"
) %>% as.data.frame()
names(momentum_continuous) <- paste0(
  "momentum_continuous_", as.character(v), "d"
)

# Generate momentum for discrete ROC
momentum_discrete <- momentum(
  days_lag = v, type = "discrete"
) %>% as.data.frame()
names(momentum_discrete) <- paste0(
  "momentum_discrete_", as.character(v), "d"
)

bitcoin_model <- bitcoin_model %>% 
  cbind(momentum_continuous) %>%
  cbind(momentum_discrete)


#--------------------------------------------------------------------------------
## Alternative Asset Class Data
#--------------------------------------------------------------------------------

# Since BTC exploded in the early 2010s, the role of it has always been widely 
# debated. BTC can be viewed as a safe haven asset class, a risk-on investment 
# or a hedge for the dollar in different periods of its rise. As such, we will 
# use 3 asset classes that represent each potential role that BTC can play. 
# The gold, for safe haven, the S&P500 as a risk on investment and the dollar 
# itself.

# We have also decided to include the USDCNY currency pair in light of the 
# recent digital yuan drive in China.

# We first scrape the internet for the relevant data.

if (!refresh_flag){
  load(file.path(path_data, "Gold.RData")) #
  load(file.path(path_data, "DXY.RData")) #
  load(file.path(path_data, "SP500.RData")) #
  load(file.path(path_data, "USDCNY.RData")) #
} else {
  ## Gold
  Gold <- Quandl("LBMA/GOLD") %>% select(1:2)
  names(Gold)[names(Gold) == "USD (AM)"] <- "close"
  names(Gold)[names(Gold) == "Date"] <- "date"
  save(Gold, file=file.path(path_data, "Gold.RData"))
  
  ## DXY 
  DXY <- get_yahoo('DX-Y.NYB') %>%
    select(date, close = adjusted_close)
  save(DXY, file=file.path(path_data, "DXY.RData"))
  
  ## SP500
  SP500 <- get_yahoo('^GSPC') %>%
    select(date, close = adjusted_close)
  save(SP500, file=file.path(path_data, "SP500.RData"))
  
  ## USD/Yuan
  USDCNY <- get_yahoo('USDCNY=X') %>%
    select(date, close = adjusted_close)
  save(USDCNY, file=file.path(path_data, "USDCNY.RData"))
}


# Then proceed to find their daily returns and lag them:
  
lagging_periods <- 90

## DXY
DXY[DXY == 0] <- NA
DXY <- DXY %>% 
  map_df(na.locf)

## Find Daily Returns
DXY <- DXY %>%
  tq_mutate(
    select = close,
    mutate_fun = periodReturn,
    period = 'daily',
    type = 'arithmetic',
    col_rename = 'future_return'
  ) 

DXY <- DXY[-1, ]

## Lagging
Ticker_name <- deparse(substitute(DXY))

close_change <- price_change(
  DXY$close, 
  n = lagging_periods, 
  type = 'discrete', 
  Ticker = Ticker_name
) # Can be changed to continuous but we will 
# stick to discrete for this one

DXY <- DXY %>%
  cbind(close_change) %>% 
  select(-c(2:3))

## Gold
Gold[Gold == 0] <- NA
Gold <- Gold %>%
  map_df(na.locf)

## Find Daily Returns
Gold <- Gold %>%
  tq_mutate(
    select = close,
    mutate_fun = periodReturn,
    period = 'daily',
    type = 'arithmetic',
    col_rename = 'future_return'
  )
Gold <- Gold[-1, ]

## Lagging
Ticker_name <- deparse(substitute(Gold))
close_change <- price_change(
  Gold$close, 
  n = lagging_periods, 
  type = 'discrete', 
  Ticker = Ticker_name
) # Can be changed to continuous but we will 
# stick to discrete for this one

Gold <- Gold %>%
  cbind(close_change) %>% 
  select(-c(2:3))

## DXY
USDCNY[USDCNY == 0] <- NA
USDCNY <- USDCNY %>%
  map_df(na.locf)

## Find Daily Returns
USDCNY <- USDCNY %>%
  tq_mutate(
    select = close,
    mutate_fun = periodReturn,
    period = 'daily',
    type = 'arithmetic',
    col_rename = 'future_return'
  ) 
USDCNY <- USDCNY[-1, ]

## Lagging
Ticker_name <- deparse(substitute(USDCNY))
close_change <- price_change(
  USDCNY$close, 
  n = lagging_periods,
  type = 'discrete',
  Ticker = Ticker_name
) # Can be changed to continuous but we will 
# stick to discrete for this one
USDCNY <- USDCNY %>%
  cbind(close_change) %>% select(-c(2:3))

## DXY
SP500[SP500 == 0] <- NA
SP500 <- SP500 %>%
  map_df(na.locf)

## Find Daily Returns
SP500 <- SP500 %>%
  tq_mutate(
    select = close,
    mutate_fun = periodReturn,
    period = 'daily',
    type = 'arithmetic',
    col_rename = 'future_return'
  )
SP500 <- SP500[-1,]

## Lagging
Ticker_name <- deparse(substitute(SP500))
close_change <- price_change(
  SP500$close,
  n = lagging_periods,
  type = 'discrete',
  Ticker = Ticker_name
) # Can be changed to continuous but we will stick 
# to discrete for this one

SP500 <- SP500 %>%
  cbind(close_change) %>% 
  select(-c(2:3))


# Finally we join them to the model:

  bitcoin_model <-left_join(
    bitcoin_model, 
    Gold, 
    by = 'date',
    all.x = F
  ) %>%
  left_join(DXY, by = 'date') %>% 
  left_join(USDCNY, by = 'date') %>% 
  left_join(SP500, by = 'date')

# Last observation carry forward for weekend data
bitcoin_model <- bitcoin_model %>% 
  mutate_if(is.numeric, na.locf0)


#--------------------------------------------------------------------------------
### Stock-to-Flow (S2F) Multiple
#--------------------------------------------------------------------------------

# The S2F ratio measures the **scarcity** of the asset, taking the **current 
# size of BTC** in the market as a ratio to the **yearly production of BTC**. 
# The S2F model developed by [PlanB](https://twitter.com/100trillionUSD?ref_src=twsrc%5Egoogle%7Ctwcamp%5Eserp%7Ctwgr%5Eauthor). 
# Although there have been rising doubts around this model for predicting 
# BTC prices, the model has been historically successful in the past few years. 
# The pattern of the spot price approaching the fair value projected with S2F 
# has been respected thus far.

# The multiple takes the ratio of the spot price against the fair value as 
# projected by S2F, i.e. a multiple nearer to 1 would indicate that BTC prices 
# are aligned to price projections due to scarcity; a multiple higher than 1 
# would likely indicate that BTC is currently overvalued based on its scarcity.

# With growing interest from financial institutions, but not yet from the mass 
# public (as seen from Google Trends), the scarcity of BTC is only going to be
# greater. Hence, we would like to include this as a feature to measure how 
# scarcity affects the spot price.


SF_date <- jsonlite::fromJSON(
  file.path(path_data, "stock-to-flow-ratio.json")
) %>% 
  as.data.frame() %>% 
  mutate(
    "Date1" = as.POSIXct(
      t, format="%Y-%m-%dT%H:%M:%S", tz = "UTC"), 
    "date" = as.Date(
      Date1, format = "%Y-%m-%d")
  ) %>% select(date)

SF_values <- jsonlite::fromJSON(
  file.path(path_data, "stock-to-flow-ratio.json")
) %>% 
  as.data.frame() %>% select(o)

SF_values <- unnest(cols=SF_values$o)

S2F_data <- cbind(SF_date, SF_values) %>%
  select(-daysTillHalving) %>% 
  rename("S2Fratio" = "ratio")

bitcoin_model <- bitcoin_model %>% 
  left_join(S2F_data, by = "date") %>% 
  mutate(S2F_multiple = close/S2Fratio) %>% 
  filter(!is.na(S2F_multiple)) %>% 
  select(-S2Fratio)
rm(SF_date, SF_values)




#--------------------------------------------------------------------------------
### Bitcoin Miner Revenue
#--------------------------------------------------------------------------------

# In this section, we will focus more on miners' revenue. Miners' revenue are 
# heavily correlated with BTC price, however, we also want to see if the rate 
# of change of revenue and momentum plays a part in determining BTC price.

# We also want to see if the moving average (MA) of miners' revenue would introduce 
# any pattern that can be picked up by the model. `frollmean()` is another useful 
# function that makes calculating a set of `1:X`-MA (where X is a positive integer) 
# easier.


if (!refresh_flag){
  load(file.path(path_data, "bc_mine_rev.RData")) #
} else {
  # bitcoin mining revenue
  bc_mine_rev <- Quandl("BCHAIN/MIREV") %>% 
    arrange(Date) %>%
    rename(miner_rev = Value)
  save(bc_mine_rev, file=file.path(path_data, "bc_mine_rev.RData"))
}

# mining revenue 1-90 day lag
bc_mine_rev_lag <- shift(
  bc_mine_rev$miner_rev, 
  n = 1:90, 
  type = 'lag', 
  give.names = TRUE
) %>% 
  as.data.frame() %>% 
  rename_all(
    funs(paste0(gsub('V1_lag_', 'miner_rev_lag_', x = .), 'd'))
  )

# mining revenue ROC
bc_mine_rev_change <- bc_mine_rev %>%
  cbind(bc_mine_rev_lag) %>%
  rename_all(funs(gsub('lag', 'change', x = .))) %>%
  mutate_if(
    grepl('miner_rev_change', names(.)), 
    ~ ifelse(. == 0, 0, (miner_rev - .)/ .)
  )

# mining revenue 2nd derive
bc_mine_rev_change2 <- bc_mine_rev_change %>%
  rename_all(funs(gsub('change', 'change2', x = .))) %>%
  mutate_if(
    grepl('miner_rev_change2', names(.)), 
    ~ ifelse(
      lag(., n = 1L) == 0, 0, 
      (. - lag(., n = 1L)) / lag(., n = 1L))
  )

# mining revenue ma 5/10/15/20/25/30/35/40/45/50
bc_mine_rev_ma <- frollmean(
  bc_mine_rev$miner_rev, 
  seq(5, 50, 5), 
  align = 'right'
) %>%
  as.data.frame()

names(bc_mine_rev_ma) <- paste0(
  sprintf(
    'bc_mine_rev_ma%s', 
    seq(5, 50, 5)), 'd')

# combine all features
bc_mine_rev <- bc_mine_rev %>% 
  cbind(bc_mine_rev_lag, 
        bc_mine_rev_ma) %>%
  inner_join(bc_mine_rev_change, by = c('Date', 'miner_rev')) %>%
  inner_join(bc_mine_rev_change2, by = c('Date', 'miner_rev'))

# mining revenue drawdown
bc_mine_rev <- bc_mine_rev %>%
  mutate(
    miner_rev_drawdown = -1 * (1 - miner_rev / cummax(miner_rev))
  )

bitcoin_model <- bitcoin_model %>%
  left_join(bc_mine_rev, by = c('date' = 'Date'))






#--------------------------------------------------------------------------------
### Techfactors
#--------------------------------------------------------------------------------

# The package `techfactor` calculates technical factors for investing instruments 
# in an efficient and appropriate manner.

# devtools::install_github("shrektan/techfactor")
# https://github.com/shrektan/techfactor/blob/master/man/tf_quote.R
# https://github.com/shrektan/techfactor
# what each alpha means can be found here: 
# https://arxiv.org/ftp/arxiv/papers/1601/1601.00991.pdf

library(techfactor)

#define function to get alpha
get_alpha <- function(df_1, df_2, n ='all', baseline = 'SP500'){
  #Techfactor only work with uppercase
  df_1 <- df_1 %>% rename_all(toupper)
  df_2 <- df_2 %>% rename_all(toupper)
  
  # Identify key as date
  df_1 <- data.table(df_1,key='DATE')
  df_2 <- data.table(df_2,key='DATE')
  
  tf_quote <- df_1[df_2] %>% 
    rename(
      BMK_CLOSE = i.CLOSE,
      BMK_OPEN = i.OPEN
    )
  # \code{BMK_OPEN} is the close and open price data of the index
  tf_quote <- tf_quote[-1,]
  
  #change 0 in pclose to na for TF to work
  tf_quote[tf_quote == 0] <- NA
  
  #get alpha
  from_to <- range(tf_quote$DATE)
  alpha_df <- tf_quote %>% 
    column_to_rownames(var = 'DATE') %>% 
    select(CLOSE)
  factors <- tf_reg_factors()
  
  if(n != 'all'){
    for(i in n:n){
      normal_factor <- attr(factors, "normal")[i]
      qt <- tf_quote_xptr(tf_quote)
      alpha <- tf_qt_cal(qt, normal_factor, from_to) %>% as_tibble()
      alpha_df <- alpha_df %>%cbind(alpha)
      # print(paste('Working on', names(alpha)))
    }
  }else{
    for(i in 1:128){
      #128 is ALL 191 factors
      normal_factor <- attr(factors, "normal")[i]
      qt <- tf_quote_xptr(tf_quote)
      alpha <- tf_qt_cal(qt, normal_factor, from_to) %>% as_tibble()
      alpha_df <- alpha_df %>%cbind(alpha)
      # print(paste('Working on', names(alpha)))
    }
  }
  
  #Add in alpha_df as new features
  alpha_df <- alpha_df %>% select(-CLOSE) %>% 
  rownames_to_column(var = 'date') %>% 
  mutate(date = ymd(date)) %>% 
    rename_if(
      is.numeric, 
      funs(paste0(baseline , gsub('alpha', '_alpha_', x = .))))

  return(alpha_df)
}


# BTC price information
bitcoin_p <- bitcoin_price %>% select(-volume_currency) %>% 
  mutate(pclose = lag(close),
         amount = weighted_price*volume_btc) %>%
  rename(vwap = weighted_price,
         volume = volume_btc) %>% 
  select(date,pclose,open,high,low,close,vwap,volume,amount)
  
# Remove first row
bitcoin_p <- bitcoin_p[-1,]

if (!refresh_flag){
  load(file.path(path_data, "SP500_.RData"))
} else {
  # Index information
  SP500 <- tq_get(
    '^GSPC',
    from = "2011-09-13",
    to = Sys.Date(),
    get = "stock.prices"
  )  
  save(SP500, file=file.path(path_data, "SP500_.RData"))
}

SP500 <- SP500 %>% 
  select(date,close,open)

alpha_df <- get_alpha(
  bitcoin_p,
  SP500, 
  n = 'all', 
  baseline = 'SP500')

bitcoin_model <- bitcoin_model %>% 
  left_join(alpha_df, by = c("date"), all.x=F)




#--------------------------------------------------------------------------------
## Bitcoin Drawdown
#--------------------------------------------------------------------------------

# We also added the drawdown of Bitcoin price, as it would potentially be 
# a valid indicator.


bitcoin_model <- bitcoin_model %>%
  mutate(close_drawdown = -1 * (1 - close / cummax(close)))


#--------------------------------------------------------------------------------
## Rolling Daily Return Volatility
#--------------------------------------------------------------------------------

# The standard deviation of Bitcoin prices could potentially be good 
# indicators as well. We added 1 day to 90 day rolling average of 
# standard deviation.


close_sd <- frollapply(
  bitcoin_model$close_change_continuous_1d, 
  1:90, sd
) %>% #We can use either the continuous or discrete
  as.data.frame()

names(close_sd) <- paste0(
  sprintf('close_sd_%s', seq(1:90)), 'd'
)

bitcoin_model <- bitcoin_model %>%
  cbind(close_sd) %>%
  mutate(close_sd_1d = 0)


#--------------------------------------------------------------------------------
## Number of Positive Days
#--------------------------------------------------------------------------------

# The number of positive days could be a good indicator of a bull or bear sum.

bitcoin_model <- bitcoin_model %>%
  mutate(
    close_positive = ifelse(close_change_continuous_1d > 0, 1, 0), 
    close_negative = ifelse(close_change_continuous_1d <= 0, 1, 0)
  )

close_positive <- frollsum(
  bitcoin_model$close_positive, 1:90, align = 'right'
) %>% as.data.frame()

names(close_positive) <- paste0(
  sprintf('close_positive_%s', seq(1:90)), 'd'
)

bitcoin_model <- bitcoin_model %>%
  cbind(close_positive)


#--------------------------------------------------------------------------------
## Number of Consecutive Positive and Negative Days
#--------------------------------------------------------------------------------

# Likewise, the number of consecutive positive and negative days could 
# also indicate bull and bear runs.


bitcoin_model <- bitcoin_model %>% 
  mutate(
    close_positive_streak = 
      close_positive * unlist(
        map(rle(close_positive)[["lengths"]], seq_len)
      ), 
    close_negative_streak = 
      close_negative * unlist(
        map(rle(close_negative)[["lengths"]], seq_len)
      )
  )


#--------------------------------------------------------------------------------
## Time Series Features
#--------------------------------------------------------------------------------

# We have added the following time series features:

# -   **entropy**: Measures the "forecastability" of a series - low values = 
#     high sig-to-noise, large vals = difficult to forecast
# -   **stability**: Means/variances are computed for all tiled windows - stability 
#     is the variance of the means
# -   **lumpiness**: Lumpiness is the variance of the variances
# -   **max_level_shift**: Finds the largest mean shift between two consecutive 
#     windows (returns two values, size of shift and time index of shift)
# -   **max_var_shift**: the max variance shift between two consecutive 
#     windows (returns two values, size of shift and time index of shift)
# -   **max_kl_shift**: Finds the largest shift in the Kulback-Leibler 
#     divergence between two consecutive windows (returns two values, size of 
#     shift and time index of shift)
# -   **crossing_points**: Number of times a series crosses the mean line


ts_feature_set <- c(
    "entropy", 
    "stability", 
    "lumpiness",     
    "max_level_shift",
    "max_var_shift",  
    "max_kl_shift", 
    "crossing_points" 
)

bitcoin_model <- bitcoin_model %>%
  mutate(ts_features = slide(
    .x = future_return,
    .f = ~ tsfeatures(.x, features = ts_feature_set),
    .before = 90,
    .complete = TRUE)) %>% 
  unnest(ts_features)


#--------------------------------------------------------------------------------
## Technical Indicators
#--------------------------------------------------------------------------------

# We will be highlighting the below indicators that form our trading strategy:

# -   20, 50 & 200 Day Exponential Moving Average (EMA): an indicator for 
#     entry/exit when yesterday's and today's candle are above/below EMA.\

# -   Golden Cross: flashes a **buy** signal when the current 20-day EMA crosses 
#     above the current 50-day EMA and a **sell** signal when the current 50-day EMA crosses under the current 20-day EMA.\

# -   Moving Average Convergence Divergence (MACD): a signal cross over after a 
#     day's candle to indicate trend-following momentum.\

# -   Stochastic Momentum Index (SMI): another typical indicator for momentum, 
#     using the 20-day EMA.

# These indicators will indicate a binary output which will be used as features 
# for our final model.

bitcoin_TI <- bitcoin_price %>% 
  select(date, close)


#--------------------------------------------------------------------------------
### Exponential Moving Average (EMA)
#--------------------------------------------------------------------------------

# 20-day EMA
bitcoin_TI <- bitcoin_TI %>% 
  mutate(
    ema20 = EMA(close, n = 20)
  ) %>% 
  mutate(
    ema_20P = Lag(
      # previous day ema > close, yesterday close > ema and 
      # today close > ema
      ifelse(
        Lag(close, 2) < Lag(ema20, 2) & 
          Lag(close) > Lag(ema20) & 
          close > ema20, 1, 
        # previous day close > ema, yesterday close < ema and 
        # today close < ema
        ifelse(
          Lag(close, 2) > Lag(ema20, 2) & 
            Lag(close) < Lag(ema20) & 
            close < ema20, 0, NA))
    )) %>% 
  fill(ema_20P, .direction = "down") %>% 
  mutate(
    ema_20P = ifelse(is.na(ema_20P) == T, 0, ema_20P)
  )

# 50-day EMA
bitcoin_TI <- bitcoin_TI %>% 
  mutate(
    ema50 = EMA(close, n = 50)
  ) %>% 
  mutate(
    ema_50P = Lag(
      ifelse(
        Lag(close, 2) < Lag(ema50, 2) & 
          Lag(close) > Lag(ema50) & close > ema50, 1, 
        ifelse(
          Lag(close, 2) > Lag(ema50, 2) & 
            Lag(close) < Lag(ema50) & close < ema50, 0, 
          NA)))
  ) %>% 
  fill(ema_50P, .direction = "down") %>% 
  mutate(ema_50P = ifelse(is.na(ema_50P) == T, 0, ema_50P))

# 200-day EMA
bitcoin_TI <- bitcoin_TI %>% 
  mutate(
    ema200 = EMA(close, n=200)
  ) %>% 
  mutate(
    ema_200P = Lag(
      ifelse(
        Lag(close, 2) < Lag(ema200, 2) & 
          Lag(close) > Lag(ema200) & 
          close > ema200, 1, 
        ifelse(
          Lag(close, 2) > Lag(ema200, 2) &
            Lag(close) < Lag(ema200) & close < ema200, 
          0, NA))
    )
  ) %>% 
  fill(ema_200P, .direction = "down") %>% 
  mutate(ema_200P = ifelse(is.na(ema_200P) == T, 0, ema_200P))


#--------------------------------------------------------------------------------
### Golden Cross
#--------------------------------------------------------------------------------


# Cross-over of ema 20 and ema 50
bitcoin_TI <- bitcoin_TI %>% 
  mutate(
    ema_CO_2050 = Lag(
      ifelse(Lag(ema20) < Lag(ema50) & ema20 > ema50, 1,
             ifelse(
               Lag(ema20) > Lag(ema50) & ema20 < ema50,
               0, NA
             )
      )
    )
  ) %>% 
  fill(ema_CO_2050, .direction = "down") %>% 
  mutate(
    ema_CO_2050 = ifelse(is.na(ema_CO_2050) == T, 0, ema_CO_2050)
  ) 


#--------------------------------------------------------------------------------
### Moving Average Convergence Divergence (MACD)
#--------------------------------------------------------------------------------


# * MACD (12,26,9) ----
macd <- MACD(
  Cl(bitcoin_price), 
  nFast = 12, 
  nSlow = 26, 
  nSig = 9
)

bitcoin_TI <- bitcoin_TI %>% 
  cbind(Lag(
    ifelse(
      Lag(macd[, 1]) < Lag(macd[, 2]) & macd[, 1] > macd[, 2],
      1,
      ifelse(
        Lag(macd[, 1]) > Lag(macd[, 2]) & macd[, 1] < macd[, 2], 
        0,NA
      )        
    )      
  ) %>% 
    as.data.frame() %>% 
    fill(Lag.1, .direction = "down")) %>% 
  rename("macd_CO" = "Lag.1") %>% 
  mutate(macd_CO = ifelse(is.na(macd_CO)== T, 0, macd_CO))


rm(macd)


#--------------------------------------------------------------------------------
### Stochastic Momentum Index (SMI)
#--------------------------------------------------------------------------------

# Using 20-day EMA as a baseline:
  

smi <- SMI(
  cbind(
    Hi(bitcoin_price), 
    Lo(bitcoin_price), 
    Cl(bitcoin_price)
  ),
  n = 13,
  nFast = 2,
  nSlow = 25,
  nSig = 9
)

bitcoin_TI <- bitcoin_TI %>% 
  mutate(smi_value = smi[,1]) %>% 
  mutate(
    SMI = Lag(
      ifelse(Lag(close) < Lag(ema20) & close > ema20 & 
               smi_value < -40, 1,
             ifelse(Lag(close) > Lag(ema20) & close < ema20 & 
                      smi_value > 40, 0, NA)
      )
    )
  ) %>% 
  fill(SMI, .direction = "down") %>% 
  mutate(SMI = ifelse(is.na(SMI) == T, 0, SMI))


rm(smi)


#--------------------------------------------------------------------------------
## Sentiment Analysis
#--------------------------------------------------------------------------------

# Sentiment analysis is quite critical for cryptocurrency prices. For a form of 
# technology that garners huge amounts of skeptism, the overall sentiment of 
# Bitcoin, in particular, can influence its prices greatly.

get_news <- function(api_key, to.date=Sys.time()){
  to.time <-  as.POSIXct(to.date) %>% as.integer()
  data <- jsonlite::fromJSON(
    paste0(
      "https://min-api.cryptocompare.com/data/v2/news/?lang=EN&api_key={",
      api_key,
      '}&feeds=cryptocompare,cointelegraph,coindesk&lTs=',
      to.time,
      sep=''))
  
  news <- data$Data %>% 
    select(id, published_on,source, title, body, url) %>% 
    mutate(date = as.POSIXct(published_on, origin="1970-01-01")) %>% 
    select(-published_on)
  
  return(news)
}

news <- readRDS(file.path(path_data, 'crypto-news.rds'))
if (refresh_flag){
  # Update with new recent articles
  api_key <- "USE-YOUR-OWN-API-KEY"
  news_temp <- get_news(api_key) # From Current -- required before loop
  date <- as.Date(min(news_temp$date))
  last_date <- as.Date(max(news$date))
  
  for(i in 1:1e7){
    #Get Data
    data <- get_news(api_key, date)
    news_temp <- rbind(news_temp, data)
    
    #Define last date for next loop 
    date <- min(data$date)
    if(as.Date(date) <= last_date){
      news <- union(news_temp, news) 
      saveRDS(news, file = file.path(path_data, 'crypto-news.rds'))
      break 
    }
    print(paste0('Iteration for news before:', date))
  }
  saveRDS(news, file=file.path(path_data, 'crypto-news.rds'))
}

news_text <- news %>% 
  select(id, date, body) %>% 
  mutate(date = as.Date(date)) %>%
  arrange(date) %>% 
  filter(date >= '2017-01-01') %>% 
  group_by(date,id) %>%  
  unnest_tokens(word, body) 

# Collate words of article into a single body
news_text_new <- news_text %>% 
  group_by(date) %>%
  summarise(TextClean = str_c(word, collapse = " "))

news_text_1 <- news_text_new %>% 
  unnest_tokens(word,TextClean) %>%
  mutate(word = gsub("[^A-Za-z ]", "", word)) %>%
  filter(word != "")

sentiment <- news_text_1 %>%
  anti_join(stop_words) %>% 
  inner_join(get_sentiments("bing")) %>%
  group_by(date) %>% 
  count(word, sentiment) %>%
  arrange(desc(date), sentiment) %>% 
  mutate(
    sentiment_sc = ifelse(
      sentiment == 'negative', -1, 1
    ),
    word_score = sentiment_sc*n
  )

# Sentiment Score will be used for Modelling
sentiment_score <- sentiment %>%  
  group_by(date) %>% 
  summarise_at(vars(word_score),sum) %>% 
  rename(sentiment_score_d = word_score)

# tail(sentiment_score, 5)


#--------------------------------------------------------------------------------
## Cleaning Data
#--------------------------------------------------------------------------------

# Note that our `technical indicator` features are binary, and thus they will 
# be treated as categorical variables.

# We will also include typical `bitcoin indicators` as listed in [Section 2.3]

bitcoin_model <- bitcoin_model %>% left_join(bitcoin_TI, by = c("date", "close"))
bitcoin_model <- bitcoin_model %>% left_join(bitcoin_data, by = "date")
bitcoin_model <- bitcoin_model %>% left_join(sentiment_score, by = 'date')

bitcoin_model <- bitcoin_model %>% map_df(~na.locf0(.))
bitcoin_model <- bitcoin_model %>% filter(date >= '2018-01-01') #first mon of 2017

# format.dt.f(bitcoin_model)
# rmarkdown::paged_table(bitcoin_model)
save(bitcoin_model, file=file.path(path_data, "bitcoin_model.RData"))

#--------------------------------------------------------------------------------
rm(bitcoin_model, alpha_df)
rm(bc_mine_rev, bc_mine_rev_change, bc_mine_rev_change2, bc_mine_rev_lag, bc_mine_rev_ma)
rm(bitcoin_data, bitcoin_p, bitcoin_price, bitcoin_TI) # bitcoin_return
rm(close_change, close_change_continuous, close_change_discrete, close_positive, close_sd)
# rm(code_list)
rm(DXY, Gold, SP500, USDCNY, vix)
rm(momentum_continuous, momentum_discrete)
rm(news, news_text, news_text_1, news_text_new)
rm(S2F_data)
rm(sentiment, sentiment_score)
rm(v, lagging_periods) # i
rm(Ticker_name)
rm(ts_feature_set)
rm(get_alpha, get_news, get_yahoo, momentum, price_change, quandl_tidy)
