rm(list=ls())
graphics.off()
#--------------------------------------------------------------------------------
#--------------------------------------------------------------------------------
# Bitcoin Momentum Feature Based Forecasting
#--------------------------------------------------------------------------------
# This example serves to highlight the effectiveness of tidymodels and we 
# recognise that the models produced here are in no more the most optimal ones.

# Firstly, we did not tune all the hyperparameters of each model and only focus
# on tuning at most one hyperparameter. With enough computational power, we would
# have tune more hyperparameters and possible make the model more complex since
# we have taken steps to ensure that overfitting problems would be addressed 
# subsequently.

# Secondly, the number of models used. We only looked at 3 models, XGB, 
# Catboost and Logistic Regression, and this is by no means an indication that
# these models are the best and produces the best results. Other models that 
# we have not included could also possibly be used and be lead to a better 
# performance.
#--------------------------------------------------------------------------------

# This document describes the workflow of machine learning applied into a 
# Bitcoin trading strategy

path_root="."
path_src=file.path(path_root, "src")
path_data=file.path(path_root, "data")
path_output=file.path(path_root, "output")
refresh_flag=FALSE

# In this report, we will highlight the basics of implementing a machine 
# learning workflow to produce a trading strategy for Bitcoin (BTC) based 
# on historical data. tidymodels has simplified machine learning, by 
# creating a clear, logical and systematic approach. tidymodels makes 
# the process one that builds upon the previous step, which will be further 
# explained in detail in this report.

# This process entails:
#----------------------------------------------------------------------------
# Data Gathering/Combining Features to Include
# source(file.path(path_src, "1_load_pkgs.R"))
source(file.path(path_src, "2_scrape_bitcoin_data.R"))
source(file.path(path_src, "3_feature_engineering.R"))

#----------------------------------------------------------------------------
# Model Exploration, Training and Evaluation
source(file.path(path_src, "4_predictive_modeling.R"))

#----------------------------------------------------------------------------
# Feature Selection
source(file.path(path_src, "5_modelling_feature_selection.R"))
#----------------------------------------------------------------------------
# Ensembling Models
source(file.path(path_src, "6_stacking.R"))

#----------------------------------------------------------------------------
# XGB and Catboost models had the best performance, and these models are 
# typically prone to overfitting. By using an array of feature selection 
# methods, we have reduced the overall number of features significantly, 
# lowering the probability of overfitting. tidymodels has made it much 
# easier and efficient to experiment with different features. As BTC 
# continues to change and as the world gradually accepts its value, this 
# paradigm shift in the realm of cryptocurrency would also change the 
# features we explored as listed in Section 3. tidyverse made it easier 
# to simply join these features together.
  
# These features could also be easily explored and reduced using tidymodels 
# native functions such as step_pca(). Package extensions such as 
# recipeselectors explored in Section 5 also synergises with tidymodels. 
# By using inbuilt functions, it opens up the opportunity to explore several 
# feature reduction methods available, all building on the recipe set and 
# the workflow.
    
# Lastly, tidymodels has also made it easier to explore these features 
# across different models. With stacks(), we could have condensed Section 4
# on Predictive Model Exploration and Evaluation into just Section 6, and 
# have the confidence that stacking would pull the best out of each model 
# and combine it into a single model.
#----------------------------------------------------------------------------
rm(path_data, path_output, path_root, path_src, refresh_flag)
