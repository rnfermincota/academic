# BUSINESS SCIENCE UNIVERSITY
# LEARNING LAB 63: MODELTIME NESTED FORECASTING
# SCRIPT: MAKE DATA

library(tidyverse)
library(vroom)

# M5 DATA ----
sales_train <- vroom::vroom("~/Downloads/sales_train_validation.csv")

calendar <- vroom::vroom("~/Downloads/calendar.csv")

calendar_tbl <- calendar %>% select(date, d)

# SUMMARIZE ITEMS ----
df <- sales_train %>%
    select(item_id, starts_with("d_")) %>%
    pivot_longer(cols = starts_with("d_"), names_to = "d")

df2 <- df %>%
    group_by(item_id, d) %>%
    summarise(value = sum(value)) %>%
    ungroup()

df3 <- df2 %>%
    group_by(item_id) %>%
    summarise(value = sum(value)) %>%
    ungroup()

top_100 <- df3 %>%
    arrange(desc(value)) %>%
    slice(1:100) %>%
    pull(item_id)

top_100

bottom <- df3 %>%
    arrange(desc(value)) %>%
    slice_tail(n=2) %>%
    pull(item_id)

df4 <- df2 %>%
    filter(item_id %in% c(top_100, bottom)) %>%
    mutate(item_id = factor(item_id, levels = c(top_100, bottom))) %>%
    arrange(item_id, d)



df5 <- df4 %>%
    pivot_wider(
        id_cols     = item_id,
        names_from  = d,
        values_from = value
    )


# CREATE SMALL DATA ----

sales_100_tbl <- df5 %>%
    pivot_longer(cols = starts_with("d_"), names_to = "d") %>%
    left_join(calendar_tbl) %>%
    select(-d) %>%
    arrange(item_id, date)

household_2_101_short <- sales_100_tbl %>%
    filter(as.numeric(item_id) == 102) %>%
    slice_tail(n = 90)

# SAVE DATA ----

sales_100_tbl %>%
    filter(!item_id == "HOUSEHOLD_2_101") %>%
    bind_rows(household_2_101) %>%
    write_rds("data/walmart_item_sales.rds")
