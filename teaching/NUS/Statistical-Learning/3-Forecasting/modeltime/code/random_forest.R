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
  path_random_forest=file.path(path_tuned, "random_forest")

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
# https://juliasilge.com/blog/sf-trees-random-tuning/
# Connects to prophet::prophet() and xgboost::xgb.train()
model_spec_random_forest_tune <- rand_forest(
  mode = "regression",
  mtry = tune(),
) %>%
  set_engine("ranger")

wflw_spec_random_forest_tune <- workflow() %>%
  add_model(model_spec_random_forest_tune) %>%
  add_recipe(
    artifacts$recipes$recipe_spec %>% 
      # update_role(Month, new_role = "indicator")
      # update_role() doesn't work with Data features, we must replace by step_rm()
      step_rm(Month)
  )

#-------------------------------------------------------------------------
#-------------------------------------------------------------------------

set.seed(123)

grid_spec1 <- grid_latin_hypercube(
  # parameters(model_spec_random_forest_tune) %>%
  hardhat::extract_parameter_set_dials(model_spec_random_forest_tune) %>%
      update(mtry = mtry(range = c(1, 49))),
  size=20
)


# plan(strategy = sequential)
plan( # toggle on parallel processing
  strategy = cluster,
  workers = parallel::makeCluster(n_cores)
)
# tic()

tuned_results_random_forest1 <- wflw_spec_random_forest_tune %>%
  tune_grid(
    resamples = resamples_kfold,
    grid = grid_spec1,
    control = control_grid(verbose = TRUE, allow_par = TRUE)
  )

# toc()
tuned_results_random_forest1 %>%
  write_rds(file.path(path_random_forest, "tuned_results_random_forest1.rds"))

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

wflw_fit_random_forest_tuned <- wflw_spec_random_forest_tune %>%
  finalize_workflow(
    select_best(tuned_results_random_forest1, "rmse", n = 1)
  ) %>%
  fit(training(splits))

wflw_fit_random_forest_tuned %>%
  write_rds(file.path(path_random_forest, "wflw_fit_random_forest_tuned.rds"))

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
wflw_fit_random_forest_tuned_rsq <- wflw_spec_random_forest_tune %>%
  finalize_workflow(
    select_best(tuned_results_random_forest1, "rsq", n = 1)
  ) %>%
  fit(training(splits))

wflw_fit_random_forest_tuned_rsq %>%
  write_rds(file.path(path_random_forest, "wflw_fit_random_forest_tuned_rsq.rds"))

# toggle off parallel processing
plan(strategy = sequential)



#-------------------------------------------------------------------------
#-------------------------------------------------------------------------
tuned_random_forest <- list(
  # Workflow spec
  tuned_wkflw_spec = wflw_spec_random_forest_tune,
  # Grid spec
  tune_grid_spec = list(
    round1 = grid_spec1
  ),
  # Tuning Results
  tuned_results = list(
    round1 = tuned_results_random_forest1
  ),
  # Tuned Workflow Fit
  tuned_wflw_fit = wflw_fit_random_forest_tuned,
  # from FE
  splits        = artifacts$splits,
  data          = artifacts$data,
  recipes       = artifacts$recipes,
  standardize   = artifacts$standardize,
  normalize     = artifacts$normalize
)

tuned_random_forest %>%
  write_rds(file.path(path_random_forest, "tuned_random_forest.rds"))
#-------------------------------------------------------------------------
#-------------------------------------------------------------------------
rm(splits, resamples_kfold, n_cores)
rm(model_spec_random_forest_tune, wflw_spec_random_forest_tune)
rm(grid_spec1)
rm(tuned_results_random_forest1)
rm(wflw_fit_random_forest_tuned, wflw_fit_random_forest_tuned_rsq)
rm(tuned_random_forest)
