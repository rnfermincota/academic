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
  path_images=file.path(path_root, "images")
  
  artifacts <- read_rds(file.path(path_features, "feature_engineering_artifacts_list.rds"))
}
#------------
#------------

splits      <- artifacts$splits
recipe_spec <- artifacts$recipes$recipe_spec
Industries  <- artifacts$data$industries

#------------
#------------
set.seed(123)
wflw_fit_random_forest <- workflow() %>%
  add_model(
    spec = rand_forest(
      mode = "regression"
    ) %>%
      set_engine("ranger")
  ) %>%
  add_recipe(
    recipe_spec %>% 
      # update_role(Month, new_role = "indicator")
      # update_role() doesn't work with Data features, we must replace by step_rm()
      step_rm(Month)
  ) %>%
  fit(training(splits))

set.seed(123)
wflw_fit_xgboost <- workflow() %>%
  add_model(
    spec = boost_tree(
      mode = "regression"
    ) %>%
      set_engine("xgboost")
  ) %>%
  add_recipe(
    recipe_spec %>%
      #update_role(Month, new_role = "indicator")
      # update_role() doesn't work with Data features, we must replace by step_rm()
      step_rm(Month)
  ) %>%
  fit(training(splits))

set.seed(123)
wflw_fit_prophet <- workflow() %>%
  add_model(
    spec = prophet_reg(
      seasonality_daily = FALSE,
      seasonality_weekly = FALSE,
      seasonality_yearly = TRUE
    ) %>%
      set_engine("prophet")
  ) %>%
  add_recipe(recipe_spec) %>%
  fit(training(splits))

set.seed(123)
wflw_fit_prophet_xgboost <- workflow() %>%
  add_model(
    spec = prophet_boost(
      seasonality_daily  = FALSE,
      seasonality_weekly = FALSE,
      seasonality_yearly = FALSE
    ) %>%
      set_engine("prophet_xgboost")
  ) %>%
  add_recipe(recipe_spec) %>%
  fit(training(splits))

#------------
#------------

submodels_tbl <- modeltime_table(
  wflw_fit_random_forest,
  wflw_fit_xgboost,
  wflw_fit_prophet,
  wflw_fit_prophet_xgboost
)

#------------
#------------

calibrated_wflws_tbl <- submodels_tbl %>%
  modeltime_calibrate(new_data = testing(splits))

#------------
#------------

workflow_artifacts <- list(
  workflows = list(
    wflw_random_forest = wflw_fit_random_forest,
    wflw_xgboost = wflw_fit_xgboost,
    wflw_prophet = wflw_fit_prophet,
    wflw_prophet_xgboost = wflw_fit_prophet_xgboost
  ),
  calibration = list(calibration_tbl = calibrated_wflws_tbl)
)

workflow_artifacts %>%
  write_rds(file.path(path_models, "workflows_artifacts_list.rds"))
#------------
#------------

rm(splits, recipe_spec, Industries)
rm(wflw_fit_random_forest, wflw_fit_xgboost, wflw_fit_prophet, wflw_fit_prophet_xgboost)
rm(submodels_tbl, calibrated_wflws_tbl, workflow_artifacts)

