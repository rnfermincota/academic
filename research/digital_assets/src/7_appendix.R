## 8.1. Appendix A (LGBM Workflow)

In LGBM, hyperparameters for `no. of leaves`, `min_data_in_leaf` and `max_depth` are the most important features. As of now, we were unable to obtain number of leaves, thus, we will tune for `min_data_in_leaf (min_n)` and `max_depth (tree_depth)`.

Implementing LGBM into [Section 4](#Section4):

# install.packages("lightgbm")
library(lightgbm)

### DEFINING MODEL ###
set_dependency("boost_tree", eng = "lightgbm", pkg = "treesnip")
lgbm_model <- boost_tree(learn_rate = 0.01,
                         tree_depth = tune(),
                         min_n = tune(),
                         mtry = 500,
                         trees = 500,
                         stop_iter = 50) %>%
  set_engine('lightgbm') %>%
  set_mode('classification')

### DEFINING WORKFLOW ###
lgbm_wflw <- workflow() %>%
  add_recipe(recipe_spec) %>%
  add_model(lgbm_model)

### DEFINING PARAMETERS ###
lgbm_params <- grid_max_entropy(parameters(min_n(), tree_depth()), size = 10)

### SETTING UP PARALLEL PROCESSING ###
all_cores <- parallel::detectCores(logical=F)
registerDoParallel(cores = all_cores)

### TUNING PARAMETERS ###
lgbm_model_trained <- tune_grid(
  lgbm_wflw, 
  grid = lgbm_params,
  metrics = metric_set(mn_log_loss),
  resamples = resamples_cv_expanding,
  control = control_resamples(verbose = T,
                              save_pred = TRUE,
                              allow_par = TRUE))

### SELECTING BEST MODEL ###
best_params <- lgbm_model_trained %>%
  select_best('mn_log_loss', maximise = FALSE)
lgbm_model_best <- lgbm_model %>%
  finalize_model(best_params)

lgbm_wflw_best <- workflow() %>%
  add_recipe(recipe_spec) %>%
  add_model(lgbm_model_best)

### TRAIN MODEL PERFORMANCE ###
train_processed <- bake(recipe_spec %>% prep(),  new_data = train)

train_prediction_lgb <- lgbm_model_best %>%
  fit(
    formula = future_return_sign ~ .,
    data = train_processed %>% select(-date)
  ) %>% 
  predict(new_data = train_processed,type='prob') %>% 
  bind_cols(train) %>% mutate(.pred_class=ifelse(.pred_0>0.5,0,1))

metrics_list <- metric_set(roc_auc, mn_log_loss, accuracy)

lgbm_score_train <- train_prediction_lgb %>% 
  metrics_list(truth=future_return_sign,estimate=.pred_class,.pred_0)

knitr::kable(lgbm_score_train)

### TEST MODEL PERFORMANCE ###
test_processed <- bake(recipe_spec %>% prep(),  new_data = test)

test_prediction_lgb <- lgbm_model_best %>%
  fit(
    formula = future_return_sign ~ .,
    data = train_processed %>% select(-date)
  ) %>% 
  predict(new_data = test_processed,type='prob') %>% 
  bind_cols(test) %>% mutate(.pred_class=ifelse(.pred_0>0.5,0,1))

lgbm_score_test <- test_prediction_lgb %>% metrics_list(future_return_sign,estimate=.pred_class,.pred_0)

knitr::kable(lgbm_score_test)
