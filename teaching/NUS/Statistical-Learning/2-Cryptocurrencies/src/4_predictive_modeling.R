# source(file.path(path_src, "1_load_pkgs.R"))

library(doParallel)
all_cores <- parallel::detectCores(logical = F)
registerDoParallel(cores = all_cores)

#----------------------------------------------------------------------------
# Predictive Modeling
#----------------------------------------------------------------------------
# After we have joined the features into our `bitcoin_model` dataframe, we 
# can now begin to build our predictive models. We will first proceed to 
# introduce the models that will be used, namely:
# -   LightGBM
# -   Catboost
# -   Logistic Regression

# The `tidymodels` package allows us to create our analysis based on a series 
# of steps.

# We first have to replace the NAs within our dataset, which are from features 
# (such as **VIX**) that are closed on weekends. We will be using the `na.locf`
# to deal with NAs.

load(file.path(path_data, "bitcoin_model.RData"))

bitcoin_return <- bitcoin_model %>%
  select(date, close, future_return, future_return_sign)

bitcoin_model <- bitcoin_model %>% 
  select(-future_return) %>%
  map_df(~na.locf0(.))


#----------------------------------------------------------------------------
## Preparation for Modelling
#----------------------------------------------------------------------------
### Initial Data Split
#----------------------------------------------------------------------------

# We first need to split the data into 2 parts, the training set and the 
# test set. The purpose of the training set is to provide models with data
# to train on when predicting whether the price of Bitcoin will increase or 
# not in the next day. For the test set, it is used as a new dataset to compare
# how well the trained model is at actually predicting, since it is technically 
# 'unseen' data.

# Since we are dealing with a time series, we will split the time according to 
# the dates, using `initial_time_split()`. We also chose to go with an 80% 
# proportion for the train test split.

data_split <- initial_time_split(bitcoin_model, prop = 0.8)
train <- training(data_split)
test <- testing(data_split)


#----------------------------------------------------------------------------
### Recipe
#----------------------------------------------------------------------------
# The first step would be to create a `recipe`, which acts as a pipeline 
# which the models we choose to use can be built upon by specifying in 
# `set_engine()`.

recipe_spec <- recipe(future_return_sign ~ ., data = train) %>%
  update_role(date, close, new_role = "ID")

# rmarkdown::paged_table(recipe_spec %>% prep() %>% juice() %>% head())

# The chosen models can then be easily cross validated by piping `time_series_cv()` 
# and `tuning` can also be easily piped with `tuneGrid()`.

#----------------------------------------------------------------------------
### Cross Validation
#----------------------------------------------------------------------------

# For time series cross validation, we recognize that they are 2 methods that we 
# can do. The expanding window method and the sliding window method. Conveniently,
# using the `modeltime` package we can easily switch between the 2 methods using 
# `cumulative = F/T`.

resamples_cv_expanding <- recipe_spec %>%
  prep() %>%
  juice() %>%
  time_series_cv(
    date_var = date,
    initial = '3 month',
    assess = '3 month',
    skip = '3 month',
    cumulative = TRUE
  )

# After comparing the results from both methods, the **expanding** method yielded 
# better results.

#----------------------------------------------------------------------------
## LightGBM
#----------------------------------------------------------------------------

# LightGBM is short for "Light Gradient Boosting Machine", a popular predictive 
# model used in the data science field. Gradient boosting produces a 
# prediction model in the form of an ensemble of weak prediction models, 
# normally decision trees. For LightGBM, this technique split leaf-wise
# rather than tree-depth wise or level-wise. Hence, it has much better accuracy
# compared to other boosting algorithm. However, due to its leaf-wise split, 
# it can be prone to overfitting.

#----------------------------------------------------------------------------
### Defining Workflow
#----------------------------------------------------------------------------

gbm_model <- boost_tree(learn_rate = 0.01,
                        tree_depth = 1,
                        min_n = 1,
                        mtry = 500,
                        trees = tune(),
                        stop_iter = 50) %>%
  set_engine('lightgbm') %>%
  set_mode('classification')

gbm_model

gbm_wflw <- workflow() %>%
  add_recipe(recipe_spec) %>%
  add_model(gbm_model)

gbm_wflw


#----------------------------------------------------------------------------
### Training Model
#----------------------------------------------------------------------------

# In this case, we are cross validated too determine the optimal number of 
# trees to build the model. Trees for a gradient boosting model is how many 
# of the weak trees to build.

gbm_params <- grid_regular(
  parameters(gbm_model), 
  levels = 6, 
  filter = c(trees > 1)
)

if (!refresh_flag){
  gbm_model_trained <- readRDS(file.path(path_output, "gbm_model_trained.rds"))
} else {
  gbm_model_trained <- tune_grid(
    gbm_wflw, 
    grid = gbm_params,
    metrics = metric_set(mn_log_loss),
    resamples = resamples_cv_expanding, #expanding is used instead of sliding
    control = control_resamples(verbose = FALSE,
                                save_pred = TRUE,
                                allow_par = TRUE))
  
  
  saveRDS(gbm_model_trained, file=file.path(path_output, "gbm_model_trained.rds"))
}

gbm_model_trained %>% collect_metrics()


#----------------------------------------------------------------------------
### Making Predictions
#----------------------------------------------------------------------------

# After training the model, we can use `select_best` to choose our best model 
# and `finalize` based on a specified metric. The team has decided to specify 
# with **mean logarithmic loss**, to see which model yields the best predictions.

# The reason that the team has decide to use the **mean logarithmic loss** is 
# because we wish to go beyond just the final class prediction and evaluate 
# the absolute probabilistic difference of each prediction. The more certain 
# our model is that an observation is 1 for example, which it is in fact, the 
# lower the error. Conversely, it penalizes very heavily when the model is 
# very certain about an outcome that is untrue.

# Once we have finalized the model, we can test our predictions using the 
# test set. We can simply `bake` the recipes and indicate a `new_data` to 
# implement the finalised model on different train/test sets.

best_params <- gbm_model_trained %>%
  select_best('mn_log_loss', maximise = FALSE)

gbm_best_params <- best_params

gbm_model_best <- gbm_model %>%
  finalize_model(best_params)

gbm_wflw_best <- workflow() %>%
  add_recipe(recipe_spec) %>%
  add_model(gbm_model_best)


#----------------------------------------------------------------------------
### Evaluating Performance
#----------------------------------------------------------------------------

# Performance on Training 
train_processed <- bake(recipe_spec %>% prep(),  new_data = train)

train_prediction_gbm <- gbm_model_best %>%
  fit(
    formula = future_return_sign ~ .,
    data = train_processed
  ) %>% 
  predict(new_data = train_processed,type = 'prob') %>% 
  bind_cols(train) %>% 
  mutate(.pred_class =ifelse(.pred_0 > 0.5, 0, 1))

metrics_list <- metric_set(mn_log_loss)

gbm_score_train <- train_prediction_gbm %>% 
  metrics_list(truth=future_return_sign,estimate=.pred_class,.pred_0)

knitr::kable(gbm_score_train)

# Performance on Testing
test_processed <- bake(recipe_spec %>% prep(),  new_data = test)

test_prediction_gbm <- gbm_model_best %>%
  fit(
    formula = future_return_sign ~ .,
    data = train_processed
  ) %>% 
  predict(new_data = test_processed,type='prob') %>% 
  bind_cols(test) %>% mutate(.pred_class=ifelse(.pred_0>0.5,0,1))

gbm_score_test <- test_prediction_gbm %>% 
  metrics_list(future_return_sign,estimate=.pred_class,.pred_0)

knitr::kable(gbm_score_test)


#----------------------------------------------------------------------------
## CatBoost
#----------------------------------------------------------------------------
# https://catboost.ai/en/docs/installation/r-installation-local-copy-installation
# install.packages('devtools')
# devtools::install_url('https://github.com/catboost/catboost/releases/download/v0.24.3/catboost-R-Windows-0.24.3.tgz', INSTALL_opts = c("--no-multiarch"))

# CatBoost is a variant of the boosting method, and is claimed to be one of 
# the models that is revolutionising the machine learning game. You can read 
# more about the nuances of CatBoost here.
# (https://hanishrohit.medium.com/whats-so-special-about-catboost-335d64d754ae)

set_dependency("boost_tree", eng = "catboost", "catboost")
set_dependency("boost_tree", eng = "catboost", "treesnip")


#----------------------------------------------------------------------------
### Defining Model
#----------------------------------------------------------------------------
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

cat_model

cat_wflw <- workflow() %>%
  add_recipe(recipe_spec) %>%
  add_model(cat_model)
cat_wflw


#----------------------------------------------------------------------------
### Training Model
#----------------------------------------------------------------------------
# Check
cat_params <- grid_regular(
  parameters(cat_model), levels = 6, filter = c(trees > 1)
)
if (!refresh_flag){
  cat_model_trained <- readRDS(file.path(path_output, "cat_model_trained.rds"))
} else {
  cat_model_trained <- tune_grid(
    cat_wflw, 
    grid = cat_params,
    metrics = metric_set(mn_log_loss), 
    resamples = resamples_cv_expanding,
    control = control_resamples(verbose = FALSE,
                                save_pred = TRUE,
                                allow_par = TRUE))
  
  saveRDS(cat_model_trained, file=file.path(path_output, "cat_model_trained.rds"))
}
cat_model_trained %>% collect_metrics()


#----------------------------------------------------------------------------
### Making Predictions
#----------------------------------------------------------------------------

best_params <- cat_model_trained %>%
  select_best('mn_log_loss', maximise = FALSE)

cat_model_best <- cat_model %>%
  finalize_model(best_params)

cat_wflw_best <- workflow() %>%
  add_recipe(recipe_spec) %>%
  add_model(cat_model_best)


#----------------------------------------------------------------------------
### Evaluating Performance
#----------------------------------------------------------------------------

# Performance on Training 
train_processed <- bake(recipe_spec %>% prep(),  new_data = train)

train_prediction_cat <- cat_model_best %>%
  fit(
    formula = future_return_sign ~ .,
    data = train_processed %>% select(-date)
  ) %>% 
  predict(new_data = train_processed,type='prob') %>% 
  bind_cols(train) %>% 
  mutate(.pred_class=ifelse(.pred_0>0.5,0,1))

metrics_list <- metric_set(mn_log_loss)

cat_score_train <- train_prediction_cat %>% 
  metrics_list(truth=future_return_sign, estimate=.pred_class, .pred_0)

knitr::kable(cat_score_train)

# Performance on Testing
test_processed <- bake(recipe_spec %>% prep(),  new_data = test)

test_prediction_cat <- cat_model_best %>%
  fit(
    formula = future_return_sign ~ .,
    data = train_processed %>% select(-date)
  ) %>% 
  predict(new_data = test_processed,type='prob') %>% 
  bind_cols(test) %>% 
  mutate(.pred_class=ifelse(.pred_0>0.5,0,1))

cat_score_test <- test_prediction_cat %>% 
  metrics_list(future_return_sign, estimate=.pred_class, .pred_0)

knitr::kable(cat_score_test)


#----------------------------------------------------------------------------
## Logistic Regression
#----------------------------------------------------------------------------

# Logistic regression is a statistical model that uses a logit function to 
# model a binary dependent variable. Essentially, you could think of it as 
# something like the linear regression of classification models.

### Defining Recipe

# By setting `mixture = 1`, it helps to remove irrelevant predictors and 
# choose a simpler model.

log_model <- logistic_reg(penalty = tune(), mixture = 1) %>% 
  set_engine("glmnet")

log_wflw <- workflow() %>% add_model(log_model) %>% 
  add_recipe(recipe_spec)


#----------------------------------------------------------------------------
### Training Model
#----------------------------------------------------------------------------
log_params <- grid_regular(
  parameters(log_model), 
  levels = 6, 
  filter = c(penalty < 1)
)

if (!refresh_flag){
  log_model_trained <- readRDS(file.path(path_output, "log_model_trained.rds"))
} else {
  
  log_model_trained <- tune_grid(
    log_wflw, 
    grid = log_params,
    metrics = metric_set(mn_log_loss),
    resamples = resamples_cv_expanding,
    control = control_resamples(verbose = T,
                                save_pred = TRUE,
                                allow_par = TRUE))
  
  saveRDS(log_model_trained, file=file.path(path_output, "log_model_trained.rds"))
}


log_model_trained %>% collect_metrics()

#----------------------------------------------------------------------------
### Making Predictions
#----------------------------------------------------------------------------

best_params <- log_model_trained %>%
  select_best('mn_log_loss', maximise = FALSE)

log_model_best <- log_model %>%
  finalize_model(best_params)

log_wflw_best <- workflow() %>%
  add_recipe(recipe_spec) %>%
  add_model(log_model_best)

#----------------------------------------------------------------------------
### Evaluating Performance
#----------------------------------------------------------------------------
# Performance on Training 
train_processed <- bake(recipe_spec %>% prep(),  new_data = train)

train_prediction_log <- log_model_best %>%
  fit(
    formula = future_return_sign ~ .,
    data = train_processed %>% select(-date)
  ) %>% 
  predict(new_data = train_processed,type='prob') %>% 
  bind_cols(train) %>% mutate(.pred_class=ifelse(.pred_0>0.5,0,1))

metrics_list <- metric_set(mn_log_loss)

log_score_train <- train_prediction_log %>% 
  metrics_list(truth=future_return_sign,estimate=.pred_class,.pred_0)

knitr::kable(log_score_train)

# Performance on Testing
test_processed <- bake(recipe_spec %>% prep(),  new_data = test)

test_prediction_log <- log_model_best %>%
  fit(
    formula = future_return_sign ~ .,
    data = train_processed %>% select(-date)
  ) %>% 
  predict(new_data = test_processed,type='prob') %>% 
  bind_cols(test) %>% mutate(.pred_class=ifelse(.pred_0>0.5,0,1))

log_score_test <- test_prediction_log %>% 
  metrics_list(future_return_sign, estimate=.pred_class, .pred_0)

knitr::kable(log_score_test)


#----------------------------------------------------------------------------
## Model Summaries (on Train and Test Set)
#----------------------------------------------------------------------------

model_summaries <- rbind(
  gbm_score_train, 
  gbm_score_test, 
  cat_score_train, 
  cat_score_test, 
  log_score_train, 
  log_score_test) 

model_summaries$Model <- c(
  rep("LightGBM", 2), rep("Catboost", 2), rep("Log Reg", 2)
)
model_summaries$Data <- rep(c("Train", "Test") , 3) 

model_summaries <- model_summaries %>% 
  select(Model, Data, .metric, .estimate) %>% 
  arrange(.estimate)

knitr::kable(model_summaries)
saveRDS(model_summaries, file=file.path(path_output, "model_summaries.rds"))


#----------------------------------------------------------------------------
## Model Performances (on Test Set)
#----------------------------------------------------------------------------

stratcomp_gbm <- test_prediction_gbm %>% 
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class),
         `Buy & Hold` = ROC(close, n = 1) %>% lead(), 
         `GBM Strategy` = .pred_class * `Buy & Hold`) %>%
  select(date, `Buy & Hold`, `GBM Strategy`)

stratcomp_cat <- test_prediction_cat %>% 
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class),
         `Buy & Hold` = ROC(close, n = 1) %>% lead(), 
         `Cat Strategy` = .pred_class*`Buy & Hold`) %>%
  select(date, `Buy & Hold`, `Cat Strategy`)

stratcomp_log <- test_prediction_log %>% 
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class),
         `Buy & Hold` = ROC(close, n = 1) %>% lead(), 
         `Log Strategy` = .pred_class*`Buy & Hold`) %>%
  select(date, `Buy & Hold`, `Log Strategy`)

stratcomp <- stratcomp_gbm %>%
  left_join(stratcomp_cat, by = c("date", "Buy & Hold")) %>%
  left_join(stratcomp_log, by = c("date", "Buy & Hold")) %>%
  column_to_rownames(var = "date") %>% 
  as.xts()

table.AnnualizedReturns(stratcomp)
charts.PerformanceSummary(stratcomp, main = "Strategy Performance")
saveRDS(stratcomp, file=file.path(path_output, "stratcomp.rds"))


# As we can see from the graph, we can identify that only the catboost models 
# is doing better than the 'Buy & Hold' strategy. 

# We should also be considering the transaction costs of buying and selling 
# Bitcoin, when determining the best strategy. CatBoost could be the best 
# performer, but with too many buying and selling transactions, the transaction 
# cost might deflate its performance significantly.

#----------------------------------------------------------------------------
### Model Performances (on Test Set) with Transaction Costs
#----------------------------------------------------------------------------

# Factoring a 0.3% transaction cost:
  
transaction_cost <- 0.003

stratcomp_gbm_cost <- test_prediction_gbm %>% 
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class),
         `Buy & Hold` = ROC(close, n = 1) %>% lead(), 
         `GBM Strategy` = .pred_class*`Buy & Hold`,
         trading_cost = (abs(.pred_class - lag(.pred_class, n = 1, default = 0))*transaction_cost),
         `GBM Strategy C` = `GBM Strategy` - trading_cost) %>%
  select(date, `Buy & Hold`, `GBM Strategy C`)

stratcomp_cat_cost <- test_prediction_cat %>% 
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class),
         `Buy & Hold` = ROC(close, n = 1) %>% lead(), 
         `Cat Strategy` = .pred_class*`Buy & Hold`,
         trading_cost = (abs(.pred_class - lag(.pred_class, n = 1, default = 0))*transaction_cost),
         `Cat Strategy C` = `Cat Strategy` - trading_cost) %>%
  select(date, `Buy & Hold`, `Cat Strategy C`)

stratcomp_log_cost <- test_prediction_log %>% 
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class),
         `Buy & Hold` = ROC(close, n = 1) %>% lead(), 
         `Log Strategy` = .pred_class*`Buy & Hold`,
         trading_cost = (abs(.pred_class - lag(.pred_class, n = 1, default = 0))*transaction_cost),
         `Log Strategy C` = `Log Strategy` - trading_cost) %>%
  select(date, `Buy & Hold`, `Log Strategy C`)

stratcomp_cost <- stratcomp_gbm_cost %>%
  left_join(stratcomp_cat_cost, by = c("date", "Buy & Hold")) %>%
  left_join(stratcomp_log_cost, by = c("date", "Buy & Hold")) %>%
  column_to_rownames(var = "date") %>% as.xts()

table.AnnualizedReturns(stratcomp_cost)
# After including the transaction cost, CatBoost is the not so effective anymore.
# charts.PerformanceSummary(stratcomp_cost, main = "Strategy (with Cost) Performance")
saveRDS(stratcomp_cost, file=file.path(path_output, "stratcomp_cost.rds"))

#----------------------------------------------------------------------------
save(
  gbm_model, gbm_params,
  cat_model, cat_params, 
  log_model, log_params,
  file=file.path(path_output, "predictive_modeling_all.RData"))
#----------------------------------------------------------------------------
#----------------------------------------------------------------------------
rm(data_split)
rm(recipe_spec, resamples_cv_expanding)
rm(bitcoin_model)

rm(gbm_model, gbm_model_trained)
rm(gbm_score_test, gbm_score_train)
rm(gbm_wflw, gbm_wflw_best)
rm(gbm_params)

rm(cat_model, cat_model_best, cat_model_trained)
rm(cat_score_test, cat_score_train)
rm(cat_wflw, cat_wflw_best)
rm(cat_params)

rm(log_model, log_model_best, log_model_trained)
rm(log_score_test, log_score_train)
rm(log_wflw, log_wflw_best)
rm(log_params)

rm(stratcomp, stratcomp_cost)
rm(stratcomp_gbm, stratcomp_gbm_cost)
rm(stratcomp_cat, stratcomp_cat_cost)
rm(stratcomp_log, stratcomp_log_cost)

rm(train, train_processed)
rm(train_prediction_cat, train_prediction_log, train_prediction_gbm)

rm(test, test_processed) 
rm(test_prediction_cat, test_prediction_log, test_prediction_gbm)

rm(all_cores, metrics_list)
rm(transaction_cost)
rm(best_params)
rm(model_summaries)
