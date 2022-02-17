# Matt to revise
# source(file.path(path_src, "1_load_pkgs.R"))

library(doParallel)
all_cores <- parallel::detectCores(logical = F)
registerDoParallel(cores = all_cores)

#----------------------------------------------------------------------------

load(file.path(path_data, "bitcoin_model.RData"))

#----------------------------------------------------------------------------
# Stacking
#----------------------------------------------------------------------------

# Instead of using `modeltime.ensemble`, given that our prediction model is
# classification-based, we will be using the `stacks` package which is an 
# extension of `tidymodels`. Model stacking is another ensembling method 
# that takes the outputs of multiple models and combines them to generate a
# new model, to generate predictions informed **by each model**.

library(stacks)
# remotes::install_github("tidymodels/stacks", ref = "main")

#----------------------------------------------------------------------------
## Defining Recipe
#----------------------------------------------------------------------------

# To successfully for an ensemble, each model must share the same resample. 
# We will use the plain vanilla models.

# train test split
data_split <- initial_time_split(bitcoin_model, prop = 0.8)
train <- training(data_split)
test <- testing(data_split)

# new recipe
recipe_stacks <- recipe(future_return_sign ~., data = train) %>% 
  update_role(date, new_role = "ID")

# getting resamples (stacks does not seem to work with time_series_cv, so we 
# will use vfold_cv() instead)
resamples_stacks <- recipe_stacks %>% 
  prep() %>% 
  juice() %>% 
  vfold_cv(v = 5)


#----------------------------------------------------------------------------
## Tuning Control
#----------------------------------------------------------------------------

Now that we have defined our base recipes and models, we have to first create a `stack_control` and we would have to save these predictions from each model as the ensemble will be using these predictions to generate the coefficients.


stack_control <- control_grid(save_pred = TRUE, save_workflow = TRUE)
stack_metrics <- metric_set(roc_auc, mn_log_loss, accuracy) # we added more metrics as stacks runs into errors with mn_log_loss


#----------------------------------------------------------------------------
## Defining Workflows
#----------------------------------------------------------------------------


xgb_wflw_stack <- workflow() %>% 
  add_model(xgb_model) %>% 
  add_recipe(recipe_stacks)

cat_wflw_stack <- workflow() %>% 
  add_model(cat_model) %>% 
  add_recipe(recipe_stacks)

log_wflw_stack <- workflow() %>% 
  add_model(log_model) %>% 
  add_recipe(recipe_stacks)


#----------------------------------------------------------------------------
## Tuning Model with Resamples
#----------------------------------------------------------------------------


xgb_res <- tune_grid(
  xgb_wflw_stack,
  resamples = resamples_stacks,
  grid = xgb_params,
  metrics = stack_metrics,
  control = stack_control
)

cat_res <- tune_grid(
  cat_wflw_stack,
  resamples = resamples_stacks,
  grid = cat_params,
  metrics = stack_metrics,
  control = stack_control
)

log_res <- tune_grid(
  log_wflw_stack,
  resamples = resamples_stacks,
  grid = log_params,
  metrics = stack_metrics,
  control = stack_control
)


#----------------------------------------------------------------------------
## Stacking Models
#----------------------------------------------------------------------------

Once we have defined the above, we can start by initiating `stacks()` and then to `add_candidates()` which would be the different models we have explored. We can then `blend_predictions()` to find out what is the combination of each model to be used in our final ensemble.

{r, eval=FALSE}
stack_model <- stacks() %>% 
  add_candidates(xgb_res) %>% 
  add_candidates(cat_res) %>% 
  add_candidates(log_res) 

stack_weights <- stack_model %>% 
  blend_predictions(non_negative = FALSE, penalty = 0.00001)


We can see the above weights that the ensemble recommends to put on each model.

`blend_predictions()` helps to remove predictors with no influence and once we pipe this through `fit_members()`, our model stack would be trained based on our input models and can predict on new data.

{r, eval=FALSE}
final_stack <- stack_weights %>% fit_members()


Essentially what `stacks()` has helped us to do, is to condense multiple lines of code which delves into:\
\* Finding the best parameters for each model and using that best model\
\* Doing a LASSO regression to find the coefficients of each model into the ensemble

We can see how efficient `stacks()` would be if there were more and more models to combine and to predict upon.

#----------------------------------------------------------------------------
## Testing Stack
#----------------------------------------------------------------------------

Once we get our final stack model, `final_stack`, we can use it to predict on our `test` data.

{r, eval=FALSE}
stack_pred <- predict(final_stack, new_data = test, type = 'prob') %>% 
  bind_cols(test) %>% mutate(.pred_class = ifelse(.pred_0 > 0.5, 0, 1)) %>% 
  mutate(.pred_class = as.factor(.pred_class)) 

stack_pred_perf <- stack_pred %>%
  metrics_list(future_return_sign,estimate = .pred_class,.pred_0)

stack_pred %>% yardstick::accuracy(truth = future_return_sign, estimate = .pred_class)

knitr::kable(stack_pred_perf)


{r, eval=FALSE}
stratcomp_stack <- stack_pred %>% 
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class) - 1,
         `Buy & Hold` = ROC(close, n = 1) %>% lead(), 
         Stack = .pred_class*`Buy & Hold`) %>%
  select(date, `Buy & Hold`, Stack) %>%
  column_to_rownames(var = "date") %>% 
  as.xts()

table.AnnualizedReturns(stratcomp_stack)
charts.PerformanceSummary(stratcomp_stack, main = "Stack Model Performance")


#----------------------------------------------------------------------------
### Stack Model Performance (with Cost)
#----------------------------------------------------------------------------

{r, eval=FALSE}
transaction_cost <- 0.003

stratcomp_stack_cost <- stack_pred %>% 
  left_join(bitcoin_return) %>%
  mutate(.pred_class = as.numeric(.pred_class) - 1,
         `Buy & Hold` = ROC(close, n = 1) %>% lead(), 
         Stack = .pred_class*`Buy & Hold`,
         trading_cost = (abs(.pred_class - lag(.pred_class, n = 1, default = 0))*transaction_cost),
         `Stack C` = Stack - trading_cost) %>%
  select(date, `Buy & Hold`, `Stack C`) %>%
  column_to_rownames(var = "date") %>% 
  as.xts()

table.AnnualizedReturns(stratcomp_stack_cost)
charts.PerformanceSummary(stratcomp_stack_cost, main = "Stack Model (with Cost) Performance")





