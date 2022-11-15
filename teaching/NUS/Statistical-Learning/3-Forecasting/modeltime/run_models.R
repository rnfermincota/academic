rm(list=ls()); graphics.off()
#------------
# RUNNING THIS SCRIPT WILL TAKE 1-2 HOURS IN TOTAL!!!
#------------

# 1 FEATURE ENGINEERING
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
library(dplyr)
library(future)
library(doFuture)
library(plotly)

# 05 ENSEMBLES
library(modeltime.ensemble)

#------------
#------------

path_root="."
path_data=file.path(path_root, "data")
path_features=file.path(path_root, "features")
path_code=file.path(path_root, "src")

path_models=file.path(path_root, "models")
unlink(file.path(path_models), recursive = T, force = T) # 9.55 GB
if (!dir.exists(path_models)) {dir.create(path_models)}

path_tuned=file.path(path_models, "tuned")
if (!dir.exists(path_tuned)) {dir.create(path_tuned)}

path_ensembles=file.path(path_models, "ensembles")
if (!dir.exists(path_ensembles)) {dir.create(path_ensembles)}

path_stacked=file.path(path_models, "stacked")
if (!dir.exists(path_stacked)) {dir.create(path_stacked)}

path_random_forest=file.path(path_tuned, "random_forest")
if (!dir.exists(path_random_forest)) {dir.create(path_random_forest)}

path_xgboost=file.path(path_tuned, "xgboost")
if (!dir.exists(path_xgboost)) {dir.create(path_xgboost)}

path_prophet=file.path(path_tuned, "prophet")
if (!dir.exists(path_prophet)) {dir.create(path_prophet)}

path_prophet_xgboost=file.path(path_tuned, "prophet_xgboost")
if (!dir.exists(path_prophet_xgboost)) {dir.create(path_prophet_xgboost)}

#------------
#------------
source(file.path(path_code, "feature_engineering.R")) 
artifacts <- read_rds(file.path(path_features, "feature_engineering_artifacts_list.rds"))

#------------
#------------

source(file.path(path_code, "non_tuned_workflows.R")) 

#------------
#------------
# Be Patient!

source(file.path(path_code, "prophet_xgboost.R"))
source(file.path(path_code, "random_forest.R"))
source(file.path(path_code, "prophet.R"))
source(file.path(path_code, "xgboost.R"))

source(file.path(path_code, "tuned_workflows.R")) 
source(file.path(path_code, "ensembles.R")) 
source(file.path(path_code, "stacked.R")) 


#------------
#------------
rm(artifacts)
rm(path_prophet_xgboost, path_random_forest, path_prophet, path_xgboost)
rm(path_root, path_features, path_models, path_tuned, path_stacked)
gc() # .rs.restartR(); memory.size(max=F)
