if (F){
  # 01 FEATURE ENGINEERING
  library(tidyverse)  # loading dplyr, tibble, ggplot2, .. dependencies
  library(timetk)  # using timetk plotting, diagnostics and augment operations
  library(tsibble)  # for month to Date conversion
  library(tsibbledata)  # for aus_retail dataset
  library(fastDummies)  # for dummyfying categorical variables
  library(skimr) # for quick statistics
  
  # 02 FEATURE ENGINEERING WITH RECIPES
  library(tidymodels) # with workflow dependency
  
  # 03 MACHINE LEARNING
  library(modeltime) # ML models specifications and engines
  library(tictoc) # measure training elapsed time
  
  # 04 HYPERPARAMETER TUNING
  library(future)
  library(doFuture)
  library(plotly)
  
  # 05 ENSEMBLES
  library(modeltime.ensemble)
  
  path_root="."
  # path_data=file.path(path_root, "data")
  path_models=file.path(path_root, "models")
  path_features=file.path(path_root, "features")
  path_tuned=file.path(path_models, "tuned")
  path_stacked=file.path(path_models, "stacked")
  
  artifacts <- read_rds(file.path(path_features, "feature_engineering_artifacts_list.rds"))
}
#------------
#------------

splits <- artifacts$splits

set.seed(123)
resamples_kfold <- training(splits) %>%
  drop_na() %>%
  vfold_cv(v = 10)

#------------
#------------

# Load all calibration tables (tuned & non-tuned models)
calibration_tbl <- read_rds(file.path(path_tuned, "workflows_NonandTuned_artifacts_list.rds"))
calibration_tbl <- calibration_tbl$calibration

submodels_resamples_kfold_tbl <- calibration_tbl %>%
  modeltime_fit_resamples(
    resamples = resamples_kfold,
    control = control_resamples(
      verbose   = TRUE,
      allow_par = TRUE,
    )
  )
submodels_resamples_kfold_tbl %>%
  write_rds(file.path(path_stacked, "submodels_resamples_kfold_tbl.rds"))
# submodels_resamples_kfold_tbl=read_rds(file.path(path_stacked, "submodels_resamples_kfold_tbl.rds"))

#------------
#------------

# Parallel Processing
registerDoFuture()
n_cores <- parallel::detectCores()

plan(
  strategy = cluster,
  workers = parallel::makeCluster(n_cores)
)

#------------
#------------

set.seed(123)
ensemble_fit_ranger_kfold <- submodels_resamples_kfold_tbl %>% # to save
  ensemble_model_spec( # with Metalearner Tuning
    model_spec = rand_forest(
      mode = "regression",
      trees = tune(),
      min_n = tune()
    ) %>%
      set_engine("ranger"),
    kfolds = 10,
    grid = 20,
    control = control_grid(verbose = TRUE, allow_par = TRUE)
  )
ensemble_fit_ranger_kfold %>%
  write_rds(file.path(path_stacked, "ensemble_fit_ranger_kfold.rds"))

set.seed(123)
ensemble_fit_xgboost_kfold <- submodels_resamples_kfold_tbl %>%
  ensemble_model_spec(
    model_spec = boost_tree(
      mode = "regression",
      trees = tune(),
      tree_depth = tune(),
      learn_rate = tune(),
      loss_reduction = tune()
    ) %>%
      set_engine("xgboost"),
    kfolds = 10,
    grid = 20,
    control = control_grid(verbose = TRUE, allow_par = TRUE)
  )
ensemble_fit_xgboost_kfold %>%
  write_rds(file.path(path_stacked,"ensemble_fit_xgboost_kfold.rds"))


set.seed(123)
ensemble_fit_svm_kfold <- submodels_resamples_kfold_tbl %>%
  ensemble_model_spec(
    model_spec = svm_rbf(
      mode = "regression",
      cost = tune(),
      rbf_sigma = tune(),
      margin = tune()
    ) %>%
      set_engine("kernlab"),
    kfold = 10,
    grid  = 20,
    control = control_grid(verbose = TRUE, allow_par = TRUE)
  )
ensemble_fit_svm_kfold %>%
  write_rds(file.path(path_stacked,"ensemble_fit_svm_kfold.rds"))


#------------
#------------

loadings_tbl <- modeltime_table(
  ensemble_fit_ranger_kfold,
  ensemble_fit_xgboost_kfold,
  ensemble_fit_svm_kfold
) %>%
  modeltime_calibrate(testing(splits)) %>%
  modeltime_accuracy() %>%
  mutate(rank = min_rank(-rmse)) %>%
  select(.model_id, rank)

stacking_fit_wt <- modeltime_table( # to save
  ensemble_fit_ranger_kfold,
  ensemble_fit_xgboost_kfold,
  ensemble_fit_svm_kfold
) %>%
  ensemble_weighted(loadings = loadings_tbl$rank)

stacking_fit_wt %>%
  write_rds(file.path(path_stacked,"stacking_fit_wt.rds"))

#------------
#------------

calibration_stacking <- stacking_fit_wt %>% # to save!
  modeltime_table() %>%
  modeltime_calibrate(testing(splits))

calibration_stacking %>%
  write_rds(file.path(path_stacked,"calibration_stacking.rds"))


#------------
#------------
# calibration_stacking=read_rds(file.path(path_stacked,"calibration_stacking.rds"))

# Toggle ON parallel processing
plan(
  strategy = cluster,
  workers  = parallel::makeCluster(n_cores)
)

set.seed(123)
refit_stacking_tbl <- calibration_stacking %>%
  modeltime_refit(
    data = artifacts$data$data_prepared_tbl,
    resamples = artifacts$data$data_prepared_tbl %>%
      drop_na() %>%
      vfold_cv(v = 10)
  )

# 12-month forecast calculations with future dataset
forecast_stacking_tbl <- refit_stacking_tbl %>%
  modeltime_forecast(
    new_data = artifacts$data$future_tbl,
    actual_data = artifacts$data$data_prepared_tbl %>%
      drop_na(),
    keep_data = TRUE
  )

# Toggle OFF parallel processing
plan(sequential)

Industries  <- artifacts$data$industries

lforecasts <- map(
  seq(length(Industries)),
  function(i){
    forecast_stacking_tbl %>%
      filter(Industry == Industries[i]) %>%
      # group_by(Industry) %>%
      mutate(
        across(
          .value:.conf_hi,
          .fns = ~standardize_inv_vec(
            x = ., 
            mean = artifacts$standardize$std_mean[i],
            sd   = artifacts$standardize$std_sd[i])
        )
      ) %>%
      mutate(across(.value:.conf_hi, .fns = ~expm1(x = .)))
  }
)

forecast_stacking_tbl <- bind_rows(lforecasts)

forecast_stacking_tbl %>%
  write_rds(file.path(path_stacked,"forecast_stacking_tbl.rds"))

#------------
#------------

rm(calibration_tbl, splits, resamples_kfold, n_cores)
rm(submodels_resamples_kfold_tbl, ensemble_fit_ranger_kfold, ensemble_fit_xgboost_kfold)
rm(ensemble_fit_svm_kfold)
rm(loadings_tbl, stacking_fit_wt)
rm(calibration_stacking)
rm(refit_stacking_tbl, forecast_stacking_tbl)
rm(lforecasts)
rm(artifacts, Industries)
