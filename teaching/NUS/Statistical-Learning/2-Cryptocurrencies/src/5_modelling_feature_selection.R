# source(file.path(path_src, "1_load_pkgs.R"))

library(doParallel)
all_cores <- parallel::detectCores(logical = F)
registerDoParallel(cores = all_cores)

#----------------------------------------------------------------------------

load(file.path(path_data, "bitcoin_model.RData"))

bitcoin_return <- bitcoin_model %>%
  select(date, close, future_return, future_return_sign)

bitcoin_model <- bitcoin_model %>% 
  select(-future_return) %>%
  map_df(~na.locf0(.))

load(file.path(path_output, "predictive_modeling.RData"))

#----------------------------------------------------------------------------
#----------------------------------------------------------------------------
# Modelling with Feature Selection
#----------------------------------------------------------------------------
#----------------------------------------------------------------------------
# We recognise the overfitting problem without reducing the number of 
# features in the dataset, as we can clearly see in the results, the 
# CV and Test result differ greatly, showing degrees of overfitting. 
# Thus we will proceed by using feature selection methods to mitigate 
# the problem of overfitting.

# With `tidymodels`, we can easily create a base recipe to build these 
# different feature selection steps upon.

recipe_select_base <- recipe(
  future_return_sign ~ ., 
  data = bitcoin_model
) %>%
  update_role(date, close, new_role = "ID") 

#----------------------------------------------------------------------------
## Through Principal Component Analysis (PCA)
#----------------------------------------------------------------------------

# Given our large dataset, Principal Component Analysis (PCA) allows us to 
# reduce the dimensionality of our dataset, while retaining majority of the 
# information provided by our features. This is done by reducing the 
# features into its **principal components** that capture the maximal 
# amount of variance in BTC prices.

# PCA entails several steps, firstly needing to standardise the range of 
# continuous features to equalise their contributions. This can be easily 
# done by first normalising our numeric columns with `step_normalize()`. 
# The next step in PCA would be then to obtain the covariance matrix, 
# which allows us to identify correlated features. Lastly, eigenvectors 
# and eigenvalues are computed from the covariance matrix to determine 
# the principal components. PCA then tries to allocate maximum possible 
# information to the first component, then the maximum **remaining** 
# information to the second component and so on.

recipe_select_pca <- recipe_select_base %>% 
  step_nzv(all_predictors()) %>% # removed near zero variance variables
  step_normalize(all_numeric(), - all_outcomes()) %>% 
  step_pca(all_predictors(), num_comp = 10, id = "pca")

pca_estimates <- prep(recipe_select_pca)

pca_info <- tidy(
  pca_estimates, 
  id = "pca", 
  type = "variance"
)

ggplot(
  pca_info %>% filter(terms == "variance"), 
  aes(x=component, y = value)
) + 
  geom_line(stat="identity") + geom_point() + 
  labs(
    x = "Principal Component", 
    y = "Variance", 
    title = "Variance Explained by Each Component"
  )

# From the above plot, we see that majority of our features do not have 
# significant impact in explaining the variances.

ggplot(
  pca_info %>% filter(terms == "variance"), 
  aes(x = component, y = value)
) + 
  geom_line(stat = "identity") + 
  geom_point() + 
  labs(x = "Principal Component", 
       y = "Variance", 
       title = "Variance Explained by Each Component") + 
  xlim(c(0,30))

# And we can see that the impact on variance begins to tail off after 
# the 10th principal component.

ggplot(
  pca_info %>% 
    filter(terms == "cumulative percent variance") %>% 
    top_n(-10, value), 
  aes(x=component, y = value)
) + 
  geom_line(stat = "identity") + 
  geom_bar(stat="identity", aes(fill = value)) + 
  geom_point() + 
  geom_text(
    aes(x = component, y = value, label = paste0(round(value, 2), "%")), 
    vjust = -.5, size = 3
  ) + 
  scale_x_discrete(limits = seq(1, 10)) + 
  labs(x = "Principal Component", 
       y = "Variance", 
       title = "Cumulative Variance Explained with Each Component", 
       subtitle = "For first 10 principal components") 

# From the above screeplot, see that across the first 10 principal 
# components, the sum of these only explain up to
paste0(
  round(
    pca_info %>% 
      filter(terms == "cumulative percent variance") %>% 
      filter(component == 10) %>% 
      pull(value), 
    1
  ), "%") # of the model.

# Thus, we will be looking at the first 10 principal components.

# new recipe
recipe_select_pca <- recipe_select_base %>% 
  step_nzv(all_predictors()) %>% 
  step_normalize(all_numeric(), - all_outcomes()) %>% 
  step_pca(all_predictors(), num_comp = 10, id = "pca")

# dataframe
pca_prep <- juice(prep(recipe_select_pca))

# Train test split pca_prep
if (!refresh_flag){
  pca_split <- readRDS(file.path(path_output, "pca_split.rds"))
} else {
  pca_split <- initial_time_split(pca_prep, prop = 0.8)
  saveRDS(pca_split, file=file.path(path_output, "pca_split.rds"))
}

train <- training(pca_split)
test <- testing(pca_split)
rm(pca_estimates, pca_info, recipe_select_pca)

# new recipe
recipe_final_pca <- recipe(future_return_sign ~., data = train) %>% 
  update_role(date, close, new_role = "ID")

# getting resamples
resamples_pca_expanding <- recipe_final_pca %>% 
  prep() %>% juice() %>% 
  time_series_cv(datevar = date, 
                 nitial = '3 month', 
                 assess = '3 month', 
                 skip = '3 month', 
                 cumulative = T)

# Getting the gbm_model


## Model of choice
selected_model <- gbm_model_best
selected_params <- gbm_best_params

pca_wflw <- workflow() %>% 
  add_recipe(recipe_final_pca) %>% 
  add_model(selected_model)

if (!refresh_flag){
  pca_trained <- readRDS(file.path(path_output, "pca_trained.rds"))
} else {
  pca_trained <- tune_grid(
    pca_wflw,
    grid = selected_params, 
    metrics = metric_set(mn_log_loss), 
    resamples = resamples_pca_expanding,
    control = control_resamples(verbose = F, save_pred = T, allow_par = T)
  )
  saveRDS(pca_trained, file=file.path(path_output, "pca_trained.rds"))
}

# After which we will evaluate the performance
best_pca_params <- pca_trained %>%
  select_best('mn_log_loss', maximise = FALSE)

pca_model_best <- selected_model %>%
  finalize_model(best_pca_params)

pca_wflw_best <- workflow() %>%
  add_recipe(recipe_final_pca) %>%
  add_model(pca_model_best)

train_processed <- bake(recipe_final_pca %>% prep(),  new_data = train)
train_prediction_pca <- pca_model_best %>%
  fit(
    formula = future_return_sign ~ .,
    data = train_processed %>% select(-date)
  ) %>% 
  predict(new_data = train_processed,type='prob') %>% 
  bind_cols(train) %>% mutate(.pred_class=ifelse(.pred_0>0.5,0,1)) %>%
  mutate(.pred_class = as.factor(.pred_class))

metrics_list <- metric_set(mn_log_loss)

pca_score_train <- train_prediction_pca %>% 
  metrics_list(truth=future_return_sign,estimate=.pred_class,.pred_0)

knitr::kable(pca_score_train)

# Performance on Testing
test_processed <- bake(recipe_final_pca %>% prep(),  new_data = test)

test_prediction_pca <- pca_model_best %>%
  fit(
    formula = future_return_sign ~ .,
    data = train_processed %>% select(-date)
  ) %>% 
  predict(new_data = test_processed,type='prob') %>% 
  bind_cols(test) %>% mutate(.pred_class=ifelse(.pred_0>0.5,0,1)) %>%
  mutate(.pred_class = as.factor(.pred_class))

pca_score_test <- test_prediction_pca %>% 
  metrics_list(future_return_sign,estimate=.pred_class,.pred_0)

knitr::kable(pca_score_test)

# However, PCA cannot capture our categorical `technical indicators` 
# features. Multiple Correspondence Analysis (MCA) or Categorical PCA 
# (CATPCA) could be used instead, but it would ignore the rest of our 
# numerical features.

#----------------------------------------------------------------------------
## Using recipeselectors 
#----------------------------------------------------------------------------
library(recipeselectors) 
# devtools::install_github("stevenpawley/recipeselectors")
# install.packages("FSelectorRcpp") # this package is needed as well

# The package `recipeselectors` is designed to enhance the `tidymodels` 
# recipe package by adding supervised feature selection steps.

# The package currently has 7 methods of feature selection listed below:
  
# - step_select_infgain which selects feature through Information Gain.
# - step_select_mrmr which selects features based on Maximum Relevancy Minimum Redundancy.
# - step_select_roc which selects Receiver Operating Curve (ROC)-based feature selection based on each predictors' relationship with the response outcome measured using a ROC curve.
# - step_select_xtab which provides feature selection through statistical association, typically for nominal variables.
# - step_select_vip which uses model-based selection using feature importance scores.
# - step_select_boruta which introduces a Boruta feature selection step.
# - step_select_carscore which provides a CAR score to select features, mainly in regression models.

# For this project, we will focus on exploring `Information Gain`, 
# `Maximum Relevancy Minimum Redundancy` and `Boruta` feature selection 
# steps.

#----------------------------------------------------------------------------
### Information Gain
#----------------------------------------------------------------------------
# Information gain calculates the reduction in entropy from transforming a
# dataset in some way. In this project, we are using information gain to 
# select the variables that maximise the information gain for the model, 
# which in turn minimizes entropy and best splits the dataset into groups 
# for effective classification.

# select features first
recipe_select_infgain <- recipe_select_base %>% 
  step_select_infgain(
    all_predictors(), 
    outcome = "future_return_sign", 
    threshold = .99
  ) 

# new dataframe
infgain_prep <- juice(prep(recipe_select_infgain))

# Train test split infgain_prep
if (!refresh_flag){
  infgain_split <- readRDS(file.path(path_output, "infgain_split.rds"))
} else {
  infgain_split <- rsample::initial_time_split(infgain_prep, prop = 0.8)
  saveRDS(infgain_split, file=file.path(path_output, "infgain_split.rds"))
}

train <- training(infgain_split)
test <- testing(infgain_split)
rm(recipe_select_infgain)

# new recipe
recipe_final_infgain <- recipe(
  future_return_sign ~., data = infgain_prep
) %>% 
  update_role(date, close, new_role = "ID")

# getting resamples
resamples_infgain_expanding <- recipe_final_infgain %>% 
  prep() %>% juice() %>% 
  time_series_cv(datevar = date, 
                 initial = '3 month', 
                 assess = '3 month', 
                 skip = '3 month', 
                 cumulative = T)

# choice of model - cat
infgain_wflw <- workflow() %>% 
  add_recipe(recipe_final_infgain) %>% 
  add_model(selected_model)

if (!refresh_flag){
  infgain_trained <- readRDS(file.path(path_output, "infgain_trained.rds"))
} else {
  infgain_trained <- tune_grid(
    infgain_wflw,
    grid = selected_params, 
    metrics = metric_set(mn_log_loss), 
    resamples = resamples_infgain_expanding,
    control = control_resamples(verbose = F, save_pred = T, allow_par = T)
  )
  saveRDS(infgain_trained, file=file.path(path_output, "infgain_trained.rds"))
}

# After which we will evaluate the performance
best_inf_params <- infgain_trained %>%
  select_best('mn_log_loss', maximise = FALSE)

inf_model_best <- selected_model %>%
  finalize_model(best_inf_params)

inf_wflw_best <- workflow() %>%
  add_recipe(recipe_final_infgain) %>%
  add_model(inf_model_best)

train_processed <- bake(
  recipe_final_infgain %>% prep(), 
  new_data = train
)

train_prediction_inf <- inf_model_best %>%
  fit(
      formula = future_return_sign ~ .,
      data = train_processed %>% select(-date)
  ) %>% 
  predict(new_data = train_processed,type='prob') %>% 
  bind_cols(train) %>% mutate(.pred_class=ifelse(.pred_0>0.5, 0, 1)) %>%
  mutate(.pred_class = as.factor(.pred_class))

metrics_list <- metric_set(mn_log_loss)

inf_score_train <- train_prediction_inf %>% 
  metrics_list(truth=future_return_sign,estimate=.pred_class,.pred_0)

knitr::kable(inf_score_train)

# Performance on Testing
test_processed <- bake(recipe_final_infgain %>% prep(),  new_data = test)

test_prediction_inf <- inf_model_best %>%
  fit(
      formula = future_return_sign ~ .,
      data = train_processed %>% select(-date)
  ) %>% 
  predict(new_data = test_processed,type='prob') %>% 
  bind_cols(test) %>% mutate(.pred_class=ifelse(.pred_0>0.5,0,1)) %>%
  mutate(.pred_class = as.factor(.pred_class))

inf_score_test <- test_prediction_inf %>% metrics_list(future_return_sign,estimate=.pred_class,.pred_0)

knitr::kable(inf_score_test)


#----------------------------------------------------------------------------
### Maximum Relevancy Minimum Redundancy
#----------------------------------------------------------------------------
# install.packages("praznik") # this package is needed as well

# Maximum relevancy minimum redundancy (mRMR) attempts to remove redundant 
# subsets of data when feature selection is done. There are many ways of 
# doing mRMR, one is by selecting features that correlate strongest to the 
# classification variable, this being "Maximum Relevancy".

# select features first
recipe_select_mrmr <- recipe_select_base %>% 
  step_select_mrmr(
    all_predictors(), outcome = "future_return_sign", threshold = .99
  )

# new dataframe
mrmr_prep <- juice(prep(recipe_select_mrmr))

# Train test split 
if (!refresh_flag){
  mrmr_split <- readRDS(file.path(path_output, "mrmr_split.rds"))
} else {
  mrmr_split <- rsample::initial_time_split(mrmr_prep, prop=0.8)
  saveRDS(mrmr_split, file=file.path(path_output, "mrmr_split.rds"))
}

train <- training(mrmr_split)
test <- testing(mrmr_split)

rm(recipe_select_mrmr)

# new recipe
recipe_final_mrmr <- recipe(future_return_sign ~., data = mrmr_prep) %>% 
  update_role(date, close, new_role = "ID")

# resamples
resamples_mrmr_expanding <- recipe_final_mrmr %>% 
  prep() %>% juice() %>% 
  time_series_cv(datevar = date, 
                 initial = '3 month', 
                 assess = '3 month', 
                 skip = '3 month', 
                 cumulative = T)

# choice of model - cat
mrmr_wflw <- workflow() %>% 
  add_recipe(recipe_final_mrmr) %>% 
  add_model(selected_model)

if (!refresh_flag){
  mrmr_trained <- readRDS(file.path(path_output, "mrmr_trained.rds"))
} else {
  mrmr_trained <- tune_grid(
    mrmr_wflw,
    grid = selected_params, 
    metrics = metric_set(mn_log_loss), 
    resamples = resamples_mrmr_expanding,
    control = control_resamples(verbose = F, save_pred = T, allow_par = T)
  )
  saveRDS(mrmr_trained, file=file.path(path_output, "mrmr_trained.rds"))
}

# After which we will evaluate the performance
best_mrmr_params <- mrmr_trained %>%
  select_best('mn_log_loss', maximise = FALSE)

mrmr_model_best <- selected_model %>%
  finalize_model(best_mrmr_params)

mrmr_wflw_best <- workflow() %>%
  add_recipe(recipe_final_mrmr) %>%
  add_model(mrmr_model_best)

train_processed <- bake(recipe_final_mrmr %>% prep(),  new_data = train)
train_prediction_mrmr <- mrmr_model_best %>%
  fit(
      formula = future_return_sign ~ .,
      data = train_processed %>% select(-date)
  ) %>% 
  predict(new_data = train_processed,type='prob') %>% 
  bind_cols(train) %>% mutate(.pred_class=ifelse(.pred_0>0.5,0,1)) %>% 
  mutate(.pred_class = as.factor(.pred_class))

metrics_list <- metric_set(mn_log_loss)

mrmr_score_train <- train_prediction_mrmr %>% 
  metrics_list(truth=future_return_sign,estimate=.pred_class,.pred_0)

knitr::kable(mrmr_score_train)

# Performance on Testing
test_processed <- bake(recipe_final_mrmr %>% prep(),  new_data = test)

test_prediction_mrmr <- mrmr_model_best %>%
  fit(
      formula = future_return_sign ~ .,
      data = train_processed %>% select(-date)
  ) %>% 
  predict(new_data = test_processed,type='prob') %>% 
  bind_cols(test) %>% mutate(.pred_class=ifelse(.pred_0>0.5,0,1)) %>%
  mutate(.pred_class = as.factor(.pred_class))

mrmr_score_test <- test_prediction_mrmr %>% metrics_list(future_return_sign,estimate=.pred_class,.pred_0)

knitr::kable(mrmr_score_test)


#----------------------------------------------------------------------------
### Boruta
#----------------------------------------------------------------------------
# install.packages("Boruta") # this package is needed as well

# The boruta algorithm is build around a random forest classification 
# algorithm. It tries to capture all the important, interesting features 
# you might have in a dataset. It creates "shadow variables" and calculates 
# the importance of all features. Features that have a higher importance 
# than the "shadow variables" are selected.

# select features
recipe_select_boruta <- recipe_select_base %>% 
  step_select_boruta(all_predictors(), outcome = "future_return_sign") 

# Train test split
if (!refresh_flag){
  boruta_prep <- readRDS(file.path(path_output, "boruta_prep.rds"))
  boruta_split <- readRDS(file.path(path_output, "boruta_split.rds"))
} else {
  boruta_prep <- juice(prep(recipe_select_boruta))
  boruta_split <- rsample::initial_time_split(boruta_prep, prop = 0.8)
  saveRDS(boruta_prep, file=file.path(path_output, "boruta_prep.rds"))
  saveRDS(boruta_split, file=file.path(path_output, "boruta_split.rds"))
}
rm(recipe_select_boruta)

train <- training(boruta_split)
test <- testing(boruta_split)

# new recipe
recipe_final_boruta <- recipe(
  future_return_sign ~ ., data = boruta_prep
) %>% 
  update_role(date, close, new_role = "ID")

# resamples
resamples_boruta_expanding <- recipe_final_boruta %>% 
  prep() %>% juice() %>% 
  time_series_cv(datevar = date, 
                 initial = '3 month', 
                 assess = '3 month', 
                 skip = '3 month', 
                 cumulative = T)

# model of choice - cat

boruta_wflw <- workflow() %>% 
  add_recipe(recipe_final_boruta) %>% 
  add_model(selected_model)

if (!refresh_flag){
  boruta_trained <- readRDS(file.path(path_output, "boruta_trained.rds"))
} else {
  boruta_trained <- tune_grid(
    boruta_wflw,
    grid = selected_params, 
    metrics = metric_set(mn_log_loss), 
    resamples = resamples_boruta_expanding,
    control = control_resamples(verbose = F, save_pred = T, allow_par = T)
  )
  saveRDS(boruta_trained, file=file.path(path_output, "boruta_trained.rds"))
}

# After which we will evaluate the performance
best_boruta_params <- boruta_trained %>%
  select_best('mn_log_loss', maximise = FALSE)

boruta_model_best <- selected_model %>%
  finalize_model(best_boruta_params)

boruta_wflw_best <- workflow() %>%
  add_recipe(recipe_final_boruta) %>%
  add_model(boruta_model_best)

train_processed <- bake(recipe_final_boruta %>% prep(),  new_data = train)
train_prediction_boruta <- boruta_model_best %>%
  fit(
      formula = future_return_sign ~ .,
      data = train_processed %>% select(-date)
  ) %>% 
  predict(new_data = train_processed,type='prob') %>% 
  bind_cols(train) %>% mutate(.pred_class=ifelse(.pred_0>0.5,0,1)) %>%
  mutate(.pred_class = as.factor(.pred_class))

metrics_list <- metric_set(mn_log_loss)

boruta_score_train <- train_prediction_boruta %>% 
  metrics_list(truth=future_return_sign,estimate=.pred_class,.pred_0)

knitr::kable(boruta_score_train)

# Performance on Testing
test_processed <- bake(recipe_final_boruta %>% prep(),  new_data = test)

test_prediction_boruta <- boruta_model_best %>%
  fit(
      formula = future_return_sign ~ .,
      data = train_processed %>% select(-date)
  ) %>% 
  predict(new_data = test_processed,type='prob') %>% 
  bind_cols(test) %>% mutate(.pred_class=ifelse(.pred_0>0.5,0,1)) %>%
  mutate(.pred_class = as.factor(.pred_class))

boruta_score_test <- test_prediction_boruta %>% metrics_list(future_return_sign,estimate=.pred_class,.pred_0)

knitr::kable(boruta_score_test)


#----------------------------------------------------------------------------
#----------------------------------------------------------------------------
## Feature Summaries
#----------------------------------------------------------------------------
#----------------------------------------------------------------------------
feat_summaries <- rbind(pca_score_train, 
                        pca_score_test, 
                        inf_score_train, 
                        inf_score_test, 
                        mrmr_score_train, 
                        mrmr_score_test, 
                        boruta_score_train, 
                        boruta_score_test) 
feat_summaries$Method <- c(rep("PCA", 2), rep("Inf. Gain", 2), rep("MRMR", 2), rep("Boruta", 2))
feat_summaries$Data <- rep(c("Train", "Test"), 4)

feat_summaries <- feat_summaries %>% 
  select(Method, Data, .metric, .estimate) %>% arrange(.estimate)
saveRDS(feat_summaries, file=file.path(path_output, "feat_summaries.rds"))

knitr::kable(feat_summaries)

#----------------------------------------------------------------------------
## Feature-Selected Model Performance
#----------------------------------------------------------------------------

stratcomp_inf <- test_prediction_inf %>% 
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class) - 1,
         `Buy & Hold` = ROC(close, n = 1) %>% lead(), 
         `Inf. CatBoost` = .pred_class*`Buy & Hold`) %>%
  select(date, `Buy & Hold`, `Inf. CatBoost`)

stratcomp_pca <- test_prediction_pca %>%
  select(-close) %>% #close in pca is normalise, so we remove it
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class) - 1,
         `Buy & Hold` = ROC(close, n = 1) %>% lead(),
         `PCA CatBoost` = .pred_class*`Buy & Hold`) %>%
  select(date, `Buy & Hold`, `PCA CatBoost`)

stratcomp_mrmr <- test_prediction_mrmr %>% 
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class) - 1,
         `Buy & Hold` = ROC(close, n = 1) %>% lead(), 
         `MRMR CatBoost` = .pred_class*`Buy & Hold`) %>%
  select(date, `Buy & Hold`, `MRMR CatBoost`)

stratcomp_boruta <- test_prediction_boruta %>% 
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class) - 1,
         `Buy & Hold` = ROC(close, n = 1) %>% lead(), 
         `Boruta CatBoost` = .pred_class*`Buy & Hold`) %>%
  select(date, `Buy & Hold`, `Boruta CatBoost`)

stratcomp_feat_select <- stratcomp_inf %>%
  left_join(stratcomp_pca, by = c("date", "Buy & Hold")) %>%
  left_join(stratcomp_mrmr, by = c("date", "Buy & Hold")) %>%
  left_join(stratcomp_boruta, by = c("date", "Buy & Hold")) %>%
  column_to_rownames(var = "date") %>% as.xts()

table.AnnualizedReturns(stratcomp_feat_select)
saveRDS(stratcomp_feat_select, file=file.path(path_output, "stratcomp_feat_select.rds"))
# charts.PerformanceSummary(stratcomp_feat_select, main = "Strategy Performance")

# From the 4 feature selection models we built, we notice that Information 
# Gain performed. Let's see the performance with trading costs.

#----------------------------------------------------------------------------
### Feature-Selected Model Performance (with Cost)
#----------------------------------------------------------------------------


transaction_cost <- 0.003

stratcomp_inf_cost <- test_prediction_inf %>% 
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class) - 1,
         `Buy & Hold` = ROC(close, n = 1) %>% lead(), 
         `Inf. CatBoost` = .pred_class*`Buy & Hold`,
         trading_cost = (abs(.pred_class - lag(.pred_class, n = 1, default = 0))*transaction_cost),
         `Inf. CatBoost C` = `Inf. CatBoost` - trading_cost) %>%
  select(date, `Buy & Hold`, `Inf. CatBoost C`)

stratcomp_pca_cost <- test_prediction_pca %>%
  select(-close) %>% #close in pca is normalise, so we remove it
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class) - 1,
         `Buy & Hold` = ROC(close, n = 1) %>% lead(),
         `PCA CatBoost` = .pred_class*`Buy & Hold`,
         trading_cost = (abs(.pred_class - lag(.pred_class, n = 1, default = 0))*transaction_cost),
         `PCA CatBoost C` = `PCA CatBoost` - trading_cost) %>%
  select(date, `Buy & Hold`, `PCA CatBoost C`)

stratcomp_mrmr_cost <- test_prediction_mrmr %>% 
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class) - 1,
         `Buy & Hold` = ROC(close, n = 1) %>% lead(), 
         `MRMR CatBoost` = .pred_class*`Buy & Hold`,
         trading_cost = (abs(.pred_class - lag(.pred_class, n = 1, default = 0))*transaction_cost),
         `MRMR CatBoost C` = `MRMR CatBoost` - trading_cost) %>%
  select(date, `Buy & Hold`, `MRMR CatBoost C`)

stratcomp_boruta_cost <- test_prediction_boruta %>% 
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class) - 1,
         `Buy & Hold` = ROC(close, n = 1) %>% lead(), 
         `Boruta CatBoost` = .pred_class*`Buy & Hold`,
         trading_cost = (abs(.pred_class - lag(.pred_class, n = 1, default = 0))*transaction_cost),
         `Boruta CatBoost C` = `Boruta CatBoost` - trading_cost) %>%
  select(date, `Buy & Hold`, `Boruta CatBoost C`)

stratcomp_cost_feat_select <- stratcomp_inf_cost %>%
  left_join(stratcomp_pca_cost, by = c("date", "Buy & Hold")) %>%
  left_join(stratcomp_mrmr_cost, by = c("date", "Buy & Hold")) %>%
  left_join(stratcomp_boruta_cost, by = c("date", "Buy & Hold")) %>%
  column_to_rownames(var = "date") %>% as.xts()

table.AnnualizedReturns(stratcomp_cost_feat_select)
saveRDS(stratcomp_cost_feat_select, file=file.path(path_output, "stratcomp_cost_feat_select.rds"))
# charts.PerformanceSummary(stratcomp_cost_feat_select, main = "Strategy (with Cost) Performance")
#----------------------------------------------------------------------------
#----------------------------------------------------------------------------
rm(cat_model, cat_params)
rm(selected_model, selected_params, recipe_select_base)
# rm(log_model, log_params)
rm(bitcoin_model, bitcoin_return)

rm(train, test)
rm(train_processed, test_processed)

rm(train_prediction_pca, test_prediction_pca)
rm(resamples_pca_expanding, recipe_final_pca)
rm(pca_model_best, pca_prep, pca_score_test, pca_score_train)
rm(pca_split, pca_trained, pca_wflw, pca_wflw_best)

rm(train_prediction_inf, test_prediction_inf)
rm(resamples_infgain_expanding, recipe_final_infgain)
rm(inf_model_best, inf_score_test, inf_score_train, inf_wflw_best)
rm(infgain_prep, infgain_split, infgain_trained, infgain_wflw)

rm(train_prediction_mrmr, test_prediction_mrmr)
rm(resamples_mrmr_expanding, recipe_final_mrmr)
rm(mrmr_model_best, mrmr_prep, mrmr_score_test, mrmr_score_train, mrmr_split)
rm(mrmr_score_test, mrmr_score_train, mrmr_model_best, mrmr_trained, mrmr_wflw, mrmr_wflw_best)

rm(train_prediction_boruta, test_prediction_boruta)
rm(resamples_boruta_expanding, recipe_final_boruta)
rm(best_boruta_params, best_inf_params, best_mrmr_params, best_pca_params)
rm(boruta_model_best, boruta_prep, boruta_score_test, boruta_score_train)
rm(boruta_score_test, boruta_score_train, boruta_model_best, boruta_split, boruta_trained, boruta_wflw, boruta_wflw_best)

rm(stratcomp_boruta, stratcomp_boruta_cost, stratcomp_cost_feat_select, stratcomp_feat_select)
rm(stratcomp_inf, stratcomp_inf_cost, stratcomp_mrmr, stratcomp_mrmr_cost)
rm(stratcomp_pca, stratcomp_pca_cost)

rm(all_cores)
rm(feat_summaries)
rm(transaction_cost)

