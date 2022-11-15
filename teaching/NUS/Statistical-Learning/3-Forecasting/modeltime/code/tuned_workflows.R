if (F){
  # 01 FEATURE ENGINEERING
  library(tidyverse)  # loading dplyr, tibble, ggplot2, .. dependencies
  library(timetk)  # using timetk plotting, diagnostics and augment operations
  library(fastDummies)  # for dummyfying categorical variables
  library(skimr) # for quick statistics
  
  # 02 FEATURE ENGINEERING WITH RECIPES
  library(tidymodels) # with workflow dependency
  
  # 03 MACHINE LEARNING
  library(modeltime) # ML models specifications and engines
  library(tictoc) # measure training elapsed time
  
  path_root=".."
  # path_data=file.path(path_root, "data")
  path_models=file.path(path_root, "models")
  path_features=file.path(path_root, "features")
  path_tuned=file.path(path_models, "tuned")
  
  artifacts <- read_rds(file.path(path_features, "feature_engineering_artifacts_list.rds"))
}
#------------
#------------
wflw_artifacts <- read_rds(file.path(path_models, "workflows_artifacts_list.rds")) # non-tuned

submodels_tbl <- modeltime_table(
  wflw_artifacts$workflows$wflw_random_forest,
  wflw_artifacts$workflows$wflw_xgboost,
  wflw_artifacts$workflows$wflw_prophet,
  wflw_artifacts$workflows$wflw_prophet_xgboost
)

tuned_random_forest=read_rds(file.path(path_random_forest, "tuned_random_forest.rds"))
tuned_xgboost=read_rds(file.path(path_xgboost, "tuned_xgboost.rds"))
tuned_prophet=read_rds(file.path(path_prophet, "tuned_prophet.rds"))
tuned_prophet_xgboost=read_rds(file.path(path_prophet_xgboost, "tuned_prophet_xgboost.rds"))

submodels_all_tbl <- modeltime_table(
  tuned_random_forest$tuned_wflw_fit,
  tuned_xgboost$tuned_wflw_fit,
  tuned_prophet$tuned_wflw_fit,
  tuned_prophet_xgboost$tuned_wflw_fit
) %>%
  update_model_description(1, "RANGER - Tuned") %>%
  update_model_description(2, "XGBOOST - Tuned") %>%
  update_model_description(3, "PROPHET W/ REGRESSORS - Tuned") %>%
  update_model_description(4, "PROPHET W/ XGBOOST ERRORS - Tuned") %>%
  combine_modeltime_tables(submodels_tbl)

splits      <- artifacts$splits

calibration_all_tbl <- submodels_all_tbl %>%
  modeltime_calibrate(testing(splits))

workflow_all_artifacts = list(
  workflows = submodels_all_tbl,
  calibration = calibration_all_tbl
)

workflow_all_artifacts %>% write_rds(file.path(path_tuned, "workflows_NonandTuned_artifacts_list.rds"))

#------------
#------------

rm(splits)
rm(tuned_random_forest, tuned_xgboost, tuned_prophet, tuned_prophet_xgboost)
rm(wflw_artifacts, submodels_tbl, submodels_all_tbl)
rm(calibration_all_tbl, workflow_all_artifacts)

