# BUSINESS SCIENCE UNIVERSITY ----
# LEARNING LAB 63: MODELTIME NESTED FORECASTING ----
# **** ----

# LIBRARIES & DATA ----

library(modeltime)
library(tidymodels)
library(tidyverse)
library(timetk)

sales_raw_tbl <- read_rds("data/walmart_item_sales.rds")

sample_12_tbl <- sales_raw_tbl %>%
    filter(as.numeric(item_id) %in% 1:12)

sample_12_tbl %>%
    group_by(item_id) %>%
    plot_time_series(date, value, .facet_ncol = 3, .smooth = FALSE)

# NESTED TIME SERIES ----

nested_data_tbl <- sales_raw_tbl %>%
    group_by(item_id) %>%
    extend_timeseries(
        .id_var = item_id,
        .date_var = date,
        .length_future = 90
    ) %>%
    nest_timeseries(
        .id_var = item_id,
        .length_future = 90
    ) %>%
    split_nested_timeseries(
        .length_test = 90
    )

nested_data_tbl %>% tail()


# MODELING ----

# * XGBoost Recipe ----

rec_xgb <- recipe(value ~ ., extract_nested_train_split(nested_data_tbl)) %>%
    step_timeseries_signature(date) %>%
    step_rm(date) %>%
    step_zv(all_predictors()) %>%
    step_dummy(all_nominal_predictors(), one_hot = TRUE)

bake(prep(rec_xgb), extract_nested_train_split(nested_data_tbl))

# * XGBoost Models ----

wflw_xgb_1 <- workflow() %>%
    add_model(boost_tree("regression", learn_rate = 0.35) %>% set_engine("xgboost")) %>%
    add_recipe(rec_xgb)

wflw_xgb_2 <- workflow() %>%
    add_model(boost_tree("regression", learn_rate = 0.50) %>% set_engine("xgboost")) %>%
    add_recipe(rec_xgb)

# * BONUS 1: New Algorithm: Temporal Hierachical Forecasting (THIEF) ----

wflw_thief <- workflow() %>%
    add_model(temporal_hierarchy() %>% set_engine("thief")) %>%
    add_recipe(recipe(value ~ ., extract_nested_train_split(nested_data_tbl)))

# 1.0 TRY 1 TIME SERIES ----
#   - Tells us if our models work at least once (before we scale)

try_sample_tbl <- nested_data_tbl %>%
    slice(1) %>%
    modeltime_nested_fit(

        model_list = list(
            wflw_xgb_1,
            wflw_xgb_2,
            wflw_thief
        ),

        control = control_nested_fit(
            verbose   = TRUE,
            allow_par = FALSE
        )
    )

try_sample_tbl

# * Check Errors ----

try_sample_tbl %>% extract_nested_error_report()


# 2.0 SCALE ----
#  - LONG RUNNING SCRIPT (2-4 MIN)

# Option 1 - Local CPUs
# parallel_start(6)

# Option 2 - Local Spark Session
library(sparklyr)
# sparklyr::spark_install()
sc <- spark_connect(master = "local[12]")
parallel_start(sc, .method = "spark")

# Takes about 2.4 min on 12-core laptop with Spark
nested_modeltime_tbl <- nested_data_tbl %>%
    # slice_tail(n = 6) %>%
    modeltime_nested_fit(

        model_list = list(
            wflw_xgb_1,
            wflw_xgb_2,
            wflw_thief
        ),

        control = control_nested_fit(
            verbose   = TRUE,
            allow_par = TRUE
        )
    )

nested_modeltime_tbl

# FILES REMOVED: Too large
# nested_modeltime_tbl %>% write_rds("artifacts/nested_modeltime_tbl.rds")
# nested_modeltime_tbl <- read_rds("artifacts/nested_modeltime_tbl.rds")

# * Review Any Errors ----
nested_modeltime_tbl %>% extract_nested_error_report()

nested_modeltime_tbl %>%
    filter(item_id == "HOUSEHOLD_2_101") %>%
    extract_nested_train_split()

# * Review Test Accuracy ----
nested_modeltime_tbl %>%
    extract_nested_test_accuracy() %>%
    table_modeltime_accuracy()

# * Visualize Test Forecast ----
nested_modeltime_tbl %>%
    extract_nested_test_forecast() %>%
    filter(item_id == "FOODS_3_090") %>%
    group_by(item_id) %>%
    plot_modeltime_forecast(.facet_ncol = 3)

# * Capture Results:
#   - Deal with small time series (<=90 days)
ids_small_timeseries <- "HOUSEHOLD_2_101"

nested_modeltime_subset_tbl <- nested_modeltime_tbl %>%
    filter(!item_id %in% ids_small_timeseries)

# 3.0 SELECT BEST ----

nested_best_tbl <- nested_modeltime_subset_tbl %>%
    modeltime_nested_select_best(metric = "rmse")

# * Visualize Best Models ----
nested_best_tbl %>%
    extract_nested_test_forecast() %>%
    filter(as.numeric(item_id) %in% 1:12) %>%
    group_by(item_id) %>%
    plot_modeltime_forecast(.facet_ncol = 3)


# 4.0 REFIT ----
#  - Long Running Script: 25 sec

nested_best_refit_tbl <- nested_best_tbl %>%
    modeltime_nested_refit(
        control = control_refit(
            verbose   = TRUE,
            allow_par = TRUE
        )
    )

# FILES REMOVED: Too large
# nested_best_refit_tbl %>% write_rds("artifacts/nested_best_refit_tbl.rds")
# nested_best_refit_tbl <- read_rds("artifacts/nested_best_refit_tbl.rds")

# * Review Any Errors ----
nested_best_refit_tbl %>% extract_nested_error_report()

# * Visualize Future Forecast ----
nested_best_refit_tbl %>%
    extract_nested_future_forecast() %>%
    filter(as.numeric(item_id) %in% 1:12) %>%
    group_by(item_id) %>%
    plot_modeltime_forecast(.facet_ncol = 3)

# 5.0 HANDLE ERRORS (SMALL TIME SERIES) ----

# * Nested Time Series ----
nested_data_small_ts_tbl <- sales_raw_tbl %>%
    filter(item_id %in% ids_small_timeseries) %>%
    group_by(item_id) %>%
    extend_timeseries(.id_var = item_id, .date_var = date, .length_future = 90) %>%
    nest_timeseries(.id_var = item_id, .length_future = 90) %>%
    split_nested_timeseries(.length_test = 30)

# * Fit, Select Best, & Refit ----
nested_best_refit_small_ts_tbl <- nested_data_small_ts_tbl %>%
    modeltime_nested_fit(

        model_list = list(
            wflw_xgb_1,
            wflw_xgb_2,
            wflw_thief
        ),

        control = control_nested_fit(
            verbose   = TRUE,
            allow_par = FALSE
        )
    ) %>%
    modeltime_nested_select_best() %>%
    modeltime_nested_refit()

nested_best_refit_small_ts_tbl %>%
    extract_nested_future_forecast() %>%
    group_by(item_id) %>%
    plot_modeltime_forecast(.facet_ncol = 3)

# * Recombine ----

nested_best_refit_all_tbl <- nested_best_refit_tbl %>%
    bind_rows(nested_best_refit_small_ts_tbl)

nested_best_refit_all_tbl %>% write_rds("artifacts/best_models_tbl.rds")

# BONUS 2: NEW WORKFLOW ----
#   - New Function: modeltime_nested_forecast()
#   - Used to make changes to your future forecast

parallel_stop()

parallel_start(6)
new_forecast_tbl <- nested_best_refit_all_tbl %>%
    modeltime_nested_forecast(
        h = 365,
        conf_interval = 0.99,
        control = control_nested_forecast(
            verbose   = TRUE,
            allow_par = FALSE
        )
    )

new_forecast_tbl %>%
    filter(as.numeric(item_id) %in% 1:12) %>%
    group_by(item_id) %>%
    plot_modeltime_forecast(.facet_ncol = 3)

# BONUS 3: SHINY APP ----

# CONCLUSIONS ----

# 1. Time Series - Pretty important for businesses.
#    We can save millions by improving their forecasting.
#    Modeltime can help. Need to learn it.

# 2. Production - Not just making models, but providing businesses
#    applications they can use. Shiny is extremely powerful.
#    Unlike Tableau and PowerBI, can run analysis through R on the fly.
#    Need to learn it.






