#----------------------------------------------------------------------------------------------------------
# Cashflows
#----------------------------------------------------------------------------------------------------------
# Here is an example of using unpivotr to tidy spreadsheets of cashflows. The techniques are:
# 1. Filter out `TOTAL` rows
# 2. Create an ordered factor of the months, which follow the fiscal year April to March. This is done using 
# the fact that the months appear in column-order as well as year-order, so we can sort on `col`.
#----------------------------------------------------------------------------------------------------------
library(dplyr)
library(stringr)
library(tidyxl)
library(unpivotr)
library(magrittr)
#----------------------------------------------------------------------------------------------------------

path_root <- "."
path_data <- file.path(path_root, "data")
path_file <- file.path(path_data, "cashflows.xlsx")

cashflows <- xlsx_cells(path_file) %>%
  dplyr::filter(
    !is_blank, 
    row >= 4L
  ) %>%
  select(row, col, data_type, character, numeric) %>%
  behead("N", "month") %>%
  behead("WNW", "main_header") %>%
  behead("W", "sub_header") %>%
  dplyr::filter(
    month != "TOTALS",
    !str_detect(sub_header, "otal")
  ) %>%
  arrange(col) %>%
  mutate(
    month = factor(
      month, 
      levels = unique(month), 
      ordered = TRUE
    ),
    sub_header = str_trim(sub_header)
  ) %>%
  select(
    main_header, 
    sub_header, 
    month, 
    value = numeric
  )
#----------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------

out1=cashflows %>% # Ending Cash Balance
  group_by(main_header, month) %>%
  summarise(value = sum(value)) %>%
  arrange(month, main_header) %>%
  dplyr::filter(str_detect(main_header, "ows")) %>%
  mutate(value = if_else(str_detect(main_header, "Income"), value, -value)) %>%
  group_by(month) %>%
  summarise(value = sum(value)) %>%
  mutate(value = cumsum(value))
#----------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------
out2=c("Beginning Cash Balance", "Ending Cash Balance") %>%
  map_int(
    ~xlsx_cells(path_file) %>%
      filter(character==.x) %>%
      pull(row)
  )
out2[1]=out2[1]-1

out2=out2 %>%
  map(
    ~xlsx_cells(path_file) %>%
      filter(row == .x) %>%
      select(col, character, numeric)
  )
out2=left_join(
  out2[[1]][,c(1,2)], 
  out2[[2]][,c(1,3)], 
  by="col"
) %>%
  drop_na() %>%
  select(-col, month=character, value=numeric) %>%
  mutate(
    month = factor(
      month, 
      levels = unique(month), 
      ordered = TRUE
    )
  )
all.equal(out1, out2)
#----------------------------------------------------------------------------------------------------------
# A tibble: 12 x 2
# month  value
# <ord>  <dbl>
#----------------------------------------------------------------------------------------------------------
# 1 April -39895
# 2 May   -43080
# 3 June  -39830
# 4 July  -14108
# 5 Aug   -25194
# 6 Sept  -42963
# 7 Oct   -39635
# 8 Nov   -29761
# 9 Dec   -49453
# 10 Jan   -30359
# 11 Feb   -33747
# 12 Mar   -27016
#----------------------------------------------------------------------------------------------------------
rm(out1, out2)
rm(path_root, path_data, path_file, cashflows)
