remotes::install_github("curso-r/treesnip@catboost")
remotes::install_github("tidymodels/stacks", ref = "main")
devtools::install_github('catboost/catboost', subdir = 'catboost/R-package')

install.packages("doParallel")
install.packages("tidymodels")
install.packages("OneR")
install.packages("timetk")
install.packages("caret")
install.packages("mlr")

library(doParallel)
library(tidymodels)
library(readr)
library(OneR)
library(timetk)
library(caret)
library(mlr)
library(TTR)
library(foreach)
library(magrittr)
library(quantmod)
library(treesnip)
library(stacks)
library(catboost)

all_cores <- parallel::detectCores(logical = F)
registerDoParallel(cores = all_cores)

rm(list = ls())

#### DATA IMPORT ----
path_root = "."
path_data=file.path(path_root, "data")

#Load files into global environment with name of file as object name
output_links=file.path(path_data, list.files(path=path_data, pattern="*.csv"))

for(file in output_links){
  name <- stringr::str_extract_all(file, "[^/]+") %>%
    pluck(1) %>% last() %>%
    stringr::str_remove(".csv")
  
  suppressWarnings(
    assign(name, readr::read_csv(file, col_types = cols()),envir=.GlobalEnv)
  )
}


etfs_prices %>% glimpse()
stock_prices %>% glimpse()

asset_prices <- rbind(etfs_prices,stock_prices)

#----------------------------------------------------------------------------
## Data Preparationg
#----------------------------------------------------------------------------
prep_price = asset_prices %>%
  select(
    ticker,
    ref.date,
    price.open,
    price.high,
    price.low,
    price.close,
    price.adjusted,
    volume
  ) %>%
  set_names(c("id",
              "date",
              "open",
              "high",
              "low",
              "close",
              "price.adj",
              "volume")
  )

lag_period = 9
period = "week"

prep_price_ret = prep_price %>%
  group_by(id) %>%
  # Convert to daily scale
  timetk::summarize_by_time(
    .date_var   = date,
    .by         = period,
    value       = last(close),
    .type       = "ceiling"
  ) %>%
  # Log returns over n lag period
  mutate(return = log(dplyr::lead((value / dplyr::lag(value, n = lag_period)),n = lag_period)),
         boxcox = dplyr::lead(timetk::box_cox_vec(value / timetk::lag_vec(value, lag = lag_period), lambda = "auto", silent = TRUE), n = lag_period)) %>%
  ungroup() %>%
  select(-value)

x <- case_when(
  period == 'day' ~ 1,
  period == 'week' ~ 2,
  period == 'month' ~ 3
)

prep_price_ret <- switch(x,prep_price_ret,
                         prep_price_ret %>%
                           mutate(date = date - 2)) #Adjust to Fridays,
out <- prep_price_ret %>%
  dplyr::left_join(prep_price, by = c('id','date'))

#----------------------------------------------------------------------------
## Industry Mapping
#----------------------------------------------------------------------------

asset_universe <- M6_Universe %>%
  set_names(c('id','class','ticker','name','sector_etftype','industry_etfsubtype','sector')) %>%
  select(c('ticker','sector'))

out <- left_join(out,asset_universe,by = c("id" = "ticker"))

#### DATA ENGINEERING ----

asset_features <- function(data, n = 10, na.rm = FALSE, timebased.feat = TRUE){
  
  # Function to add technical indicators
  In<-function(data, p = n){
    # replace NA with 0
    data <- data %>%
      mutate(open = na.locf0(open),
             high = na.locf0(high),
             low = na.locf0(low),
             close = na.locf0(close),
             price.adj = na.locf0(price.adj),
             volume = na.locf0(volume),
             med = na.locf0(med))
    
    # calculate input matrix
    adx  <- ADX(HLC(data), n = p)
    ar   <- aroon(data[ ,c('high', 'low')], n = p)[ ,'oscillator']
    cci  <- CCI(HLC(data), n = p)
    chv  <- chaikinVolatility(HLC(data), n = p)
    cmo  <- CMO(data[ ,'med'], n = p)
    macd <- MACD(data[ ,'med'], 12, 26, 9)[ ,'macd']
    rsi  <- RSI(data[ ,'med'], n = p)
    stoh <- stoch(HLC(data),14, 3, 3)
    vol  <- volatility(OHLC(data), n = p, calc = "yang.zhang", N = 96)
    vwap <- VWAP(Cl(data),data[,c('volume')],p)
    In   <- cbind(adx, ar, cci, chv, cmo, macd, rsi, stoh, vol, vwap)
    
    return(In)
  }
  
  # 1. Addition of basic features
  data_eng_tbl = data %>%
    group_by(id) %>%
    mutate(
      CO    = close - open,              #' CO: Difference of Close and Open (Close−Open)
      HO    = high - open,               #' HO: Difference of High and Open (High−Open)
      LO    = low - open,                #' LO: Difference of Low and Open (Low−Open)
      HL    = high - low,                #' HL: Difference of High and Low (High−Low)
      dH    = c(NA, diff(high)),         #' dH: High of the previous 15min candle (Lag(High))
      dL    = c(NA, diff(low)),          #' dL: Low of the previous 15min candle (Lag(Low))
      dC    = c(NA, diff(close)),        #' dC: Close of the previous 15min candle (Lag(Close))
      med   = (high + close)/2,          #' dC: Close of the previous 15min candle (High + Close)
      HL_2  = (high + low)/2,            #' HL_2: Average of the High and Low (High+Low)/2
      HLC_3 = (high + low + close)/3,    #' HLC_3: Average of the High, Low and Close (High+Low+Close)/3
      Wg    = (high + low + 2 * close)/4 #' Wg: Weighted Average of the High, Low, Close by 0.25,0.25,0.5 (High+Low+2(Close))/4
    ) %>%
    ungroup()
  
  # 2. Addition of technical indicators (Refer to the functions.R code)
  # Apply function across all asset ids
  asset_name = unique(data_eng_tbl$id)

  ind <- foreach(i = 1:length(asset_name),.packages = "dplyr",.combine = "rbind")%dopar%{
    In(data = data_eng_tbl %>% filter(id == asset_name[i]))}

  final_data_eng_tbl <- data_eng_tbl %>%
    cbind(ind)
  
  if(na.rm){final_data_eng_tbl %<>% drop_na()}
  
  # 3. Addition of time based features
  if(timebased.feat){final_data_eng_tbl %<>% tk_augment_timeseries_signature(.date_var = date) %>%
      tk_augment_holiday_signature(.date_var = date,.holiday_pattern = "^$", .locale_set = "all", .exchange_set = "all") %>%
      select(-c(index.num,diff,hour,minute,second,hour12,am.pm,wday.lbl,month.lbl,mday,ends_with(".iso"),ends_with(".xts"))) %>% arrange(id,date)}
  
  return(final_data_eng_tbl)
}


out <- asset_features(out)

#### TARGET VARIABLE ----

create_target <- function(data, bin = 5, nlabel = c("rank_1","rank_2","rank_3","rank_4","rank_5"), method = "content"){
  data$return = data$return %>% replace(is.na(.), 0)
  target_df <- data %>% group_by(date) %>%
    mutate(Ranking = OneR::bin(return, nbins = bin,
                               labels = nlabel,
                               method = method)) %>%
    ungroup()
  
  return(target_df)
}

asset_tbl <- create_target(out, bin = 5, 
                           nlabel = c("rank_1","rank_2","rank_3","rank_4","rank_5"),
                           method = "content")

#### DATA TRAINING ----

asset_dataset <- asset_tbl %>%
  select(-c('return','boxcox'))

#----------------------------------------------------------------------------
## One Hot Encoding
#----------------------------------------------------------------------------
asset_numeric <- recipe(Ranking ~ ., data = asset_dataset) %>%
  update_role(c('date','id'), new_role = "id") %>%
  update_role(all_of('Ranking'), new_role = "outcome") %>%
  step_dummy(all_nominal_predictors(), one_hot = T) %>%
  prep() %>% juice() %>% arrange(id, date)


#----------------------------------------------------------------------------
## Feature Selection
#----------------------------------------------------------------------------

remove_corr <- function(data, target,corr_cut = 0.7){
  x    <- data %>% dplyr::select(-c("id","date",target))
  y_id <- data[,c("id","date",target)]
  
  descCor  <- suppressWarnings(cor(x))
  descCor[is.na(descCor)] = 0.99
  highCor  <- caret::findCorrelation(descCor, cutoff = corr_cut)
  x.f      <- x[ ,-highCor]
  data_out <- cbind(y_id,x.f) %>% as_tibble()
  num = length(data) - length(data_out)
  message(paste('Remove corr:',num,'variables removed'))
  return(data_out)
}

remove_constants <- function(data,target){
  x    <- data %>% dplyr::select(-c("id","date",target))
  y_id <- data[,c("id","date",target)]
  
  x.f <- suppressMessages(mlr::removeConstantFeatures(x, perc=.10, na.ignore = TRUE))
  data_out <- cbind(y_id, x.f) %>% as_tibble()
  num = length(data) - length(data_out)
  message(paste('Remove constant:',num,'variables removed'))
  return(data_out)
}

remove_duplicates <- function(data, target){
  x    <- data %>% dplyr::select(-c("id","date",target))
  y_id <- data[,c("id","date",target)]
  
  x.f <- x[!duplicated(as.list(x))]
  data_out <- cbind(y_id, x.f) %>% as_tibble()
  
  num = length(data) - length(data_out)
  message(paste('Remove duplicates:',num,'variables removed'))
  return(data_out)
}

feature_selection <- function(data, target, corr = 0.7){
  id_col <- data %>% select(c("id", "date", target))
  x <- data %>% select(-c("id", "date", target))
  x_rm <- foreach(i=1:ncol(x),.packages = c("dplyr","zoo","tidyr","purrr"),.combine = cbind) %dopar% {
    x[i] %>% pluck(1) %>% na.locf0() %>% replace_na(0)
  }
  colnames(x_rm) <- colnames(x) 
  data <- cbind(id_col, x_rm)
  
  out <- data %>%
    remove_corr(target = target,corr_cut = corr) %>%
    remove_constants(target = target) %>%
    remove_duplicates(target = target)
  return(out)
}

asset_select <- feature_selection(asset_numeric, target = 'Ranking', corr = 0.7)

#### PREDICTIONS ----
# Prediction will be considered under two core framework
# 1. Prediction - ML core to predict the performance of the asset
# 2. Investment Decision

#----------------------------------------------------------------------------
## Preparation for Model Training
#----------------------------------------------------------------------------

splits = time_series_split(asset_select, date_var = date, assess = '9 week', cumulative = TRUE)

data_prepared = training(splits) %>% mutate(Ranking = droplevels(Ranking))
data_test = testing(splits)

recipe_stacks <- recipe(Ranking ~ ., data = data_prepared) %>%
  update_role(c('date','id'), new_role = "id") %>%
  update_role(all_of('Ranking'), new_role = "outcome")

resamples_stacks <- recipe_stacks %>% 
  prep() %>% 
  juice() %>% 
  vfold_cv(v = 5)

#----------------------------------------------------------------------------
## Tuning Control
#----------------------------------------------------------------------------

stack_control <- control_grid(save_pred = TRUE, save_workflow = TRUE)
stack_metrics <- metric_set(mn_log_loss,roc_auc)

#----------------------------------------------------------------------------
## Defining Models
#----------------------------------------------------------------------------
set_dependency("boost_tree", eng = "catboost", "catboost")
set_dependency("boost_tree", eng = "catboost", "treesnip")

cat_model <- boost_tree(
  learn_rate = 0.01,
  tree_depth = 1,
  min_n = 1,
  mtry = 500,
  trees = tune(),
  stop_iter = 50
) %>%
  set_engine('catboost') %>%
  set_mode('classification')

log_model <- multinom_reg(penalty = tune(), mixture = 1) %>% 
  set_engine("glmnet")

cat_params <- grid_regular(
  parameters(cat_model), levels = 6, filter = c(trees > 1))

log_params <- grid_regular(
  parameters(log_model), 
  levels = 6, 
  filter = c(penalty < 1)
)
#----------------------------------------------------------------------------
## Defining Workflows
#----------------------------------------------------------------------------

cat_wflw_stack <- workflow() %>% 
  add_model(cat_model) %>% 
  add_recipe(recipe_stacks)

log_wflw_stack <- workflow() %>% 
  add_model(log_model) %>% 
  add_recipe(recipe_stacks)

#----------------------------------------------------------------------------
## Tuning Model with Resamples
#----------------------------------------------------------------------------

cat_res <- tune_grid(
  cat_wflw_stack,
  resamples = resamples_stacks,
  grid = cat_params,
  metrics = stack_metrics,
  control = stack_control
)

log_res <- tune_grid(
  log_wflw_stack,
  resamples = resamples_stacks,
  grid = log_params,
  metrics = stack_metrics,
  control = stack_control
)

#----------------------------------------------------------------------------
## Stacking Models
#----------------------------------------------------------------------------

stack_model <- stacks() %>% 
  add_candidates(cat_res) %>% 
  add_candidates(log_res) 

stack_weights <- stack_model %>% 
  blend_predictions(non_negative = FALSE, penalty = 0.00001)

final_stack <- stack_weights %>% fit_members()

#----------------------------------------------------------------------------
## Model Prediction
#----------------------------------------------------------------------------

stack_pred <- predict(final_stack, new_data = data_test, type = 'prob') %>%
  bind_cols(data_test) %>% select(contains(c(".pred","id","date")))

final_preds <- stack_pred %>% subset(date == "2022-09-16") # This is actually predicting 9 weeks later

#----------------------------------------------------------------------------
## Investment Decision
#----------------------------------------------------------------------------

invest <- foreach(i=1:nrow(final_preds),.packages = c("dplyr","zoo","tidyr","purrr"),.combine = rbind) %dopar% {
  final_preds[i,1] %>% pluck(1) * - 0.075 + final_preds[i,2] %>% pluck(1) * - 0.025 + final_preds[i,4] %>% pluck(1) * 0.025 + final_preds[i,5] %>% pluck(1) * 0.075
}

final_preds <- final_preds %>% mutate(Decision = invest[,1]) %>% select(c('id',contains(".pred"),'Decision')) %>%
                set_names(c('ID','Rank1','Rank2','Rank3','Rank4','Rank5','Decision'))

#----------------------------------------------------------------------------
## Saving it as submission
#----------------------------------------------------------------------------

write.csv(final_preds,"sample_submission.csv", row.names = FALSE)