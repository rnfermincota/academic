if (F){
  rm(list = ls())
  # 1 FEATURE ENGINEERING
  library(tidyverse) # loading dplyr, tibble, ggplot2, .. dependencies
  library(timetk) # using timetk plotting, diagnostics and augment operations
  library(tsibble) # for month to Date conversion
  library(tsibbledata) # for aus_retail dataset
  library(fastDummies) # for dummyfying categorical variables
  library(skimr) # for quick statistics
  
  # 2 FEATURE ENGINEERING WITH RECIPES
  library(tidymodels)
  
  # 3 MACHINE LEARNING
  library(modeltime)
  library(tictoc)
  
  # 4 HYPERPARAMETER TUNING
  library(future)
  library(doFuture)
  library(plotly)
  
  path_root=".."
  path_features=file.path(path_root, "features")
  path_models=file.path(path_root, "models")
  
  path_tuned=file.path(path_models, "tuned")
  path_prophet=file.path(path_tuned, "prophet")

  artifacts <- read_rds(file.path(path_features, "feature_engineering_artifacts_list.rds"))
}

splits <- artifacts$splits

# k = 10 folds
set.seed(123)
resamples_kfold <- training(splits) %>%
  vfold_cv(v = 10)
# Registers the doFuture parallel processing
registerDoFuture()
n_cores <- parallel::detectCores()

#-------------------------------------------------------------------------
#-------------------------------------------------------------------------
# Connects to prophet::prophet() and xgboost::xgb.train()
model_spec_prophet_tune <- prophet_reg(
  mode = "regression",
  changepoint_num = tune(),
  seasonality_yearly = TRUE,
  seasonality_weekly = FALSE,
  seasonality_daily = FALSE,
  # changepoint_range = tune(),
  # prior_scale_changepoints = tune(),
  # prior_scale_seasonality = tune(),
  # prior_scale_holidays = tune()
) %>%
  set_engine("prophet")

wflw_spec_prophet_tune <- workflow() %>%
  add_model(model_spec_prophet_tune) %>%
  add_recipe(artifacts$recipes$recipe_spec)

#-------------------------------------------------------------------------
#-------------------------------------------------------------------------

set.seed(123)
grid_spec1 <- grid_latin_hypercube(
  hardhat::extract_parameter_set_dials(model_spec_prophet_tune),
  size = 20
)

# plan(strategy = sequential)
plan( # toggle on parallel processing
  strategy = cluster,
  workers = parallel::makeCluster(n_cores)
)
# tic()

tuned_results_prophet1 <- wflw_spec_prophet_tune %>%
  tune_grid(
    resamples = resamples_kfold,
    grid = grid_spec1,
    control = control_grid(verbose = TRUE, allow_par = TRUE)
  )

# toc()
tuned_results_prophet1 %>%
  write_rds(file.path(path_prophet, "tuned_results_prophet1.rds"))

# toggle off parallel processing
plan(strategy = sequential)



#-------------------------------------------------------------------------
#-------------------------------------------------------------------------

# plan(strategy = sequential)
plan( # toggle on parallel processing
  strategy = cluster,
  workers = parallel::makeCluster(n_cores)
)

# Fitting round 3 best RMSE model
set.seed(123)

wflw_fit_prophet_tuned <- wflw_spec_prophet_tune %>%
  finalize_workflow(
    select_best(tuned_results_prophet1, "rmse", n = 1)
  ) %>%
  fit(training(splits))

wflw_fit_prophet_tuned %>%
  write_rds(file.path(path_prophet, "wflw_fit_prophet_tuned.rds"))

# toggle off parallel processing
plan(strategy = sequential)


#-------------------------------------------------------------------------
# Fitting round 3 best RSQmodel
# plan(strategy = sequential)
plan( # toggle on parallel processing
  strategy = cluster,
  workers = parallel::makeCluster(n_cores)
)

set.seed(123)

wflw_fit_prophet_tuned_rsq <- wflw_spec_prophet_tune %>%
  finalize_workflow(
    select_best(tuned_results_prophet1, "rsq", n = 1)
  ) %>%
  fit(training(splits))

wflw_fit_prophet_tuned_rsq %>%
  write_rds(file.path(path_prophet, "wflw_fit_prophet_tuned_rsq.rds"))

# toggle off parallel processing
plan(strategy = sequential)

#-------------------------------------------------------------------------
#-------------------------------------------------------------------------
tuned_prophet <- list(
  # Workflow spec
  tuned_wkflw_spec = wflw_spec_prophet_tune,
  # Grid spec
  tune_grid_spec = list(
    round1 = grid_spec1
  ),
  # Tuning Results
  tuned_results = list(
    round1 = tuned_results_prophet1
  ),
  # Tuned Workflow Fit
  tuned_wflw_fit = wflw_fit_prophet_tuned,
  # from FE
  splits        = artifacts$splits,
  data          = artifacts$data,
  recipes       = artifacts$recipes,
  standardize   = artifacts$standardize,
  normalize     = artifacts$normalize
)

tuned_prophet %>%
  write_rds(file.path(path_prophet, "tuned_prophet.rds"))
#-------------------------------------------------------------------------
#-------------------------------------------------------------------------
rm(splits, resamples_kfold, n_cores)
rm(model_spec_prophet_tune, wflw_spec_prophet_tune)
rm(grid_spec1)
rm(tuned_results_prophet1)
rm(wflw_fit_prophet_tuned, wflw_fit_prophet_tuned_rsq)
rm(tuned_prophet)
