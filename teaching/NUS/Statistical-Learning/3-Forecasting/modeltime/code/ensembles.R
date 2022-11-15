if (F){

  # 01 FEATURE ENGINEERING
  library(tidyverse) # loading dplyr, tibble, ggplot2, .. dependencies
  library(timetk) # using timetk plotting, diagnostics and augment operations
  library(tsibble) # for month to Date conversion
  library(tsibbledata) # for aus_retail dataset
  library(fastDummies) # for dummyfying categorical variables
  library(skimr) # for quick statistics
  
  # 02 FEATURE ENGINEERING WITH RECIPES
  library(tidymodels) # with workflow dependency
  
  # 03 MACHINE LEARNING
  library(modeltime) # ML models specifications and engines
  library(tictoc) # measure training elapsed times
  
  # 04 HYPERPARAMETER TUNING
  library(future)
  library(doFuture)
  library(plotly)
  
  # 05 ENSEMBLES
  library(modeltime.ensemble)
  
  path_root=".."
  # path_data=file.path(path_root, "data")
  path_models=file.path(path_root, "models")
  path_features=file.path(path_root, "features")
  path_tuned=file.path(path_models, "tuned")
  path_ensembles=file.path(path_models, "ensembles")
  
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

workflows_NonandTuned_artifacts_list <- read_rds(
  file.path(path_tuned, "workflows_NonandTuned_artifacts_list.rds")
)

calibration_tbl <- workflows_NonandTuned_artifacts_list$calibration

submodels_tbl <- workflows_NonandTuned_artifacts_list$workflows
ensemble_fit_mean <- submodels_tbl %>% 
  ensemble_average(type = "mean")

ensemble_fit_median <- submodels_tbl %>%
  ensemble_average(type = "median")

loadings_tbl <- calibration_tbl %>%
  modeltime_accuracy() %>%
  mutate(rank = min_rank(-rmse))

loadings_tbl <- loadings_tbl %>%
  select(.model_id, rank)

ensemble_fit_wt <- submodels_tbl %>%
  ensemble_weighted(loadings = loadings_tbl$rank)

calibration_all_tbl1 <- modeltime_table( # save it
  ensemble_fit_mean,
  ensemble_fit_median,
  ensemble_fit_wt
) %>%
  modeltime_calibrate(testing(splits))

calibration_all_tbl1 %>%
  write_rds(file.path(path_ensembles, "calibration_all_tbl1.rds"))


calibration_all_tbl2 <- calibration_all_tbl1 %>%
  combine_modeltime_tables(calibration_tbl) %>%
  modeltime_calibrate(testing(splits))

calibration_all_tbl2 %>%
  write_rds(file.path(path_ensembles, "calibration_all_tbl2.rds"))

#------------
#------------
rm(splits, resamples_kfold)
rm(workflows_NonandTuned_artifacts_list)
rm(calibration_tbl, submodels_tbl)
rm(ensemble_fit_mean, ensemble_fit_median)
rm(loadings_tbl)
rm(ensemble_fit_wt)
rm(calibration_all_tbl1, calibration_all_tbl2)
