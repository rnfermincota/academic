if (F){
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
  path_xgboost=file.path(path_tuned, "xgboost")

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
model_spec_xgboost_tune <- boost_tree(
  mode = "regression",
  mtry = tune(),
  trees = tune(),
  min_n = tune(),
  tree_depth = tune(),
  learn_rate = tune(),
  loss_reduction = tune(),
) %>%
  set_engine("xgboost")

wflw_spec_xgboost_tune <- workflow() %>%
  add_model(model_spec_xgboost_tune) %>%
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
  # parameters(model_spec_xgboost_tune) %>%
  hardhat::extract_parameter_set_dials(model_spec_xgboost_tune) %>%
      update(mtry = mtry(range = c(1, 49))),
  size=20
)


# plan(strategy = sequential)
plan( # toggle on parallel processing
  strategy = cluster,
  workers = parallel::makeCluster(n_cores)
)
# tic()

tuned_results_xgboost1 <- wflw_spec_xgboost_tune %>%
  tune_grid(
    resamples = resamples_kfold,
    grid = grid_spec1,
    control = control_grid(verbose = TRUE, allow_par = TRUE)
  )

# toc()
tuned_results_xgboost1 %>%
  write_rds(file.path(path_xgboost, "tuned_results_xgboost1.rds"))

# toggle off parallel processing
plan(strategy = sequential)

#-------------------------------------------------------------------------

# update or adjust the parameter range within the grid specification.
set.seed(123)
grid_spec2 <- grid_latin_hypercube(
  # parameters(model_spec_xgboost_tune) %>%
  hardhat::extract_parameter_set_dials(model_spec_xgboost_tune) %>%
  update(
    mtry = mtry(range = c(1, 49)),
    learn_rate = learn_rate(range = c(-2.0, -1.0))
  ),
  size = 20
)


# perform hyperparameter tuning with new grid specification
# plan(strategy = sequential)
plan( # toggle on parallel processing
  strategy = cluster,
  workers = parallel::makeCluster(n_cores)
)
# tic()

tuned_results_xgboost2 <- wflw_spec_xgboost_tune %>%
  tune_grid(
    resamples = resamples_kfold,
    grid = grid_spec2,
    control = control_grid(verbose = TRUE, allow_par = TRUE)
  )

# toc()

tuned_results_xgboost2 %>%
  write_rds(file.path(path_xgboost, "tuned_results_xgboost2.rds"))

# toggle off parallel processing
plan(strategy = sequential)

#-------------------------------------------------------------------------

set.seed(123)

grid_spec3 <- grid_latin_hypercube(
  # parameters(model_spec_xgboost_tune) %>%
  hardhat::extract_parameter_set_dials(model_spec_xgboost_tune) %>%
  update(
    mtry = mtry(range = c(1, 49)),
    learn_rate = learn_rate(range = c(-2.0, -1.0)),
    trees = trees(range = c(1283, 1906))
  ),
  size = 20
)

# plan(strategy = sequential)
plan( # toggle on parallel processing
  strategy = cluster,
  workers = parallel::makeCluster(n_cores)
)

# tic()

tuned_results_xgboost3 <- wflw_spec_xgboost_tune %>%
  tune_grid(
    resamples = resamples_kfold,
    grid      = grid_spec3,
    control   = control_grid(verbose = TRUE, allow_par = TRUE)
  )

# toc()

tuned_results_xgboost3 %>%
  write_rds(file.path(path_xgboost, "tuned_results_xgboost3.rds"))

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

wflw_fit_xgboost_tuned <- wflw_spec_xgboost_tune %>%
  finalize_workflow(
    select_best(tuned_results_xgboost3, "rmse", n = 1)
  ) %>%
  fit(training(splits))

wflw_fit_xgboost_tuned %>%
  write_rds(file.path(path_xgboost, "wflw_fit_xgboost_tuned.rds"))

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
wflw_fit_xgboost_tuned_rsq <- wflw_spec_xgboost_tune %>%
  finalize_workflow(
    select_best(tuned_results_xgboost3, "rsq", n = 1)
  ) %>%
  fit(training(splits))

wflw_fit_xgboost_tuned_rsq %>%
  write_rds(file.path(path_xgboost, "wflw_fit_xgboost_tuned_rsq.rds"))

# toggle off parallel processing
plan(strategy = sequential)



#-------------------------------------------------------------------------
#-------------------------------------------------------------------------
tuned_xgboost <- list(
  # Workflow spec
  tuned_wkflw_spec = wflw_spec_xgboost_tune,
  # Grid spec
  tune_grid_spec = list(
    round1 = grid_spec1,
    round2 = grid_spec2,
    round3 = grid_spec3
  ),
  # Tuning Results
  tuned_results = list(
    round1 = tuned_results_xgboost1,
    round2 = tuned_results_xgboost2,
    round3 = tuned_results_xgboost3
  ),
  # Tuned Workflow Fit
  tuned_wflw_fit = wflw_fit_xgboost_tuned,
  # from FE
  splits        = artifacts$splits,
  data          = artifacts$data,
  recipes       = artifacts$recipes,
  standardize   = artifacts$standardize,
  normalize     = artifacts$normalize
)

tuned_xgboost %>%
  write_rds(file.path(path_xgboost, "tuned_xgboost.rds"))

#-------------------------------------------------------------------------
#-------------------------------------------------------------------------
rm(splits, resamples_kfold, n_cores)
rm(model_spec_xgboost_tune, wflw_spec_xgboost_tune)
rm(grid_spec1, grid_spec2, grid_spec3)
rm(tuned_results_xgboost1, tuned_results_xgboost2, tuned_results_xgboost3)
rm(wflw_fit_xgboost_tuned, wflw_fit_xgboost_tuned_rsq)
rm(tuned_xgboost)