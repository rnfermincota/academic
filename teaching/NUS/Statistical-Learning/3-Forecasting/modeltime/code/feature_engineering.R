if (F){
  rm(list = ls())
  
  path_root=".."
  path_data=file.path(path_root, "data")
  path_features=file.path(path_root, "features")
  path_images=file.path(path_root, "images")
  
  #-------------
  
  # PART 1 - FEATURE ENGINEERING
  library(tidyverse)
  library(timetk)
  library(fastDummies)
  library(skimr)
  
  # PART 2 - FEATURE ENGINEERING WITH RECIPES
  library(tidymodels)
}

#-------------

aus_retail=readr::read_rds(file.path(path_data, "aus_retail.rds"))

monthly_retail_tbl <- aus_retail %>%
  filter(State == "Australian Capital Territory") %>%
  # mutate(Month = as.Date(Month)) %>%
  mutate(Industry = as_factor(Industry)) %>%
  select(Month, Industry, Turnover)

#-------------

Industries <- monthly_retail_tbl %>% distinct(Industry) %>% pull(Industry)

groups <- map(
  Industries,
  ~monthly_retail_tbl %>%
    filter(Industry == .x) %>%
    arrange(Month) %>%
    mutate(Turnover = log1p(x = Turnover)) %>%
    mutate(Turnover = standardize_vec(Turnover)) %>%
    future_frame(Month, .length_out = "12 months", .bind_data = TRUE) %>%
    mutate(Industry = .x) %>%
    tk_augment_fourier(.date_var = Month, .periods = 12, .K = 1) %>%
    # tk_augment_lags(.value = Turnover, .lags = c(12)) %>%
    tk_augment_lags(.value = Turnover, .lags = c(12, 13)) %>%
    tk_augment_slidify(
      # .value   = c(Turnover_lag12),
      .value   = c(Turnover_lag12, Turnover_lag13),
      .f       = ~ mean(.x, na.rm = TRUE),
      .period  = c(3, 6, 9, 12),
      .partial = TRUE,
      .align   = "center"
    )
)

groups_fe_tbl <- bind_rows(groups) %>%
  rowid_to_column(var = "rowid")

#-------------

tmp <- monthly_retail_tbl %>%
  group_by(Industry) %>% 
  arrange(Month) %>%
  mutate(Turnover = log1p(x = Turnover)) %>% 
  group_map(
    ~c(
      mean = mean(.x$Turnover, na.rm = TRUE),
      sd = sd(.x$Turnover, na.rm = TRUE)
    )
  ) %>% 
  bind_rows()

std_mean <- tmp$mean
std_sd <- tmp$sd

#-------------

data_prepared_tbl <- groups_fe_tbl %>%
  filter(!is.na(Turnover)) %>%
  drop_na()

#-------------

future_tbl <- groups_fe_tbl %>%
  filter(is.na(Turnover))

#-------------

# data_prepared_tbl %>% group_by(Industry) %>% tally() # = 428 and 428*0.2 ~ 85.6
splits <- data_prepared_tbl %>%
  time_series_split(
    Month,
    assess = "86 months", # same as 428*0.2 ~ 86 (rounded up)
    cumulative = TRUE
  )

splits # split object

#-------------

recipe_spec <- recipe(Turnover ~ ., data = training(splits)) %>%
  update_role(rowid, new_role = "indicator") %>%
  step_other(Industry) %>%
  step_timeseries_signature(Month) %>%
  step_rm(matches("(.xts$)|(.iso$)|(hour)|(minute)|(second)|(day)|(week)|(am.pm)")) %>%
  step_dummy(all_nominal(), one_hot = TRUE) %>%
  step_normalize(Month_index.num, Month_year)

#-------------
myskim <- skim_with(numeric = sfl(max, min), append = TRUE)

tmp_rec=recipe(Turnover ~ ., data = training(splits)) %>%
  step_timeseries_signature(Month) %>%
  prep() %>%
  juice() %>%
  myskim()

# Month_index.num: 
Month_index.num_limit_lower <- tmp_rec$numeric.p0[which(tmp_rec$skim_variable=="Month_index.num")]
Month_index.num_limit_lower

Month_index.num_limit_upper <- tmp_rec$numeric.max[which(tmp_rec$skim_variable=="Month_index.num")]
Month_index.num_limit_upper

# Month_year:
Month_year_limit_lower <- tmp_rec$numeric.p0[which(tmp_rec$skim_variable=="Month_year")]
Month_year_limit_lower

Month_year_limit_upper <- tmp_rec$numeric.max[which(tmp_rec$skim_variable=="Month_year")]
Month_year_limit_upper

#-------------

feature_engineering_artifacts_list <- list(
  # Data
  data = list(
    data_prepared_tbl = data_prepared_tbl,
    future_tbl      = future_tbl,
    industries = Industries
  ),
  
  # Recipes
  recipes = list(
    recipe_spec = recipe_spec
  ),
  
  # Splits
  splits = splits,
  
  # Inversion Parameters
  standardize = list(
    std_mean = std_mean,
    std_sd   = std_sd
  ),
  
  normalize = list(
    Month_index.num_limit_lower = Month_index.num_limit_lower, 
    Month_index.num_limit_upper = Month_index.num_limit_upper,
    Month_year_limit_lower = Month_year_limit_lower,
    Month_year_limit_upper = Month_year_limit_upper
  )  
)

feature_engineering_artifacts_list %>% 
  write_rds(file.path(path_features, "feature_engineering_artifacts_list.rds"))

#-------------
rm(data_prepared_tbl, future_tbl, Industries, recipe_spec, splits, std_mean, std_sd)
rm(groups, groups_fe_tbl)
rm(Month_index.num_limit_lower, Month_index.num_limit_upper, Month_year_limit_lower, Month_year_limit_upper)
rm(feature_engineering_artifacts_list)
rm(tmp, myskim, tmp_rec)
rm(monthly_retail_tbl, aus_retail)
