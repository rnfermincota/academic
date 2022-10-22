rm(list=ls())
#----------------
# Deadline: October 15th.
# https://rpubs.com/rafael_nicolas/singapore_screener
# https://www.linkedin.com/posts/rnfc_the-monetary-authority-of-singapore-together-activity-6850991876242313216-J7dZ/

#----------------
path_root <- ".."
path_data <- file.path(path_root, "data") 
path_xlsx <- file.path(path_data, "xlsx")
# https://github.com/rnfermincota/academic/tree/main/research/traditional_assets/database/data/singapore
#----------------

library(tidyxl)
library(tibble)
library(dplyr)
library(purrr)
library(furrr); plan(multicore)
library(unpivotr) # https://nacnudus.github.io/unpivotr/
#----------------

extract.sheets.f <- function(dsheet, excl=NULL){
  # inp=dat[1,]$dataset[[1]]
  # dsheet=filter(inp, sheet=="cost_capital" )
  out <- dsheet %>% 
    select(row, col, data_type, character, numeric, logical) %>% 
    rectify %>% # https://rdrr.io/cran/unpivotr/man/rectify.html
    select(-c("row/col"))
  nm <- out %>% slice(1) %>% unlist(use.names=FALSE) # headings
  out <- set_names(out, nm) %>% slice(2:n())
  # glimpse(out)
  if (!is.null(excl)){
    out <- out %>% mutate(across(-excl, as.numeric))
    }
  return(out)
}

#----------------
list_paths <- list.files(
  path = file.path(path_xlsx),
  pattern = ".xlsx",
  full.names = TRUE
)

#----------------
dat=enframe(list_paths, name = NULL, value = "path") %>%
  mutate(sector=gsub(".xlsx", "", basename(path)))

dat=dat %>% # slice(1:2) %>%
  mutate(
    dataset=future_map(path, xlsx_cells) %>% 
      set_names(sector)
  )

#----------------
excl=c("country", "company_name", "industry_group")
singapore_industries=future_map(
  dat$dataset,
  function(inp){
    extract.sheets.f(dplyr::filter(inp, sheet=="earnings_debt"), all_of(excl)) %>%
      slice(1)
  }
) %>% 
  bind_rows(.) %>% 
  arrange(industry_group) %>%
  arrange(desc(roic_cost_capital))

singapore_earnings_debt=future_map(
  dat$dataset,
  function(inp){
    extract.sheets.f(dplyr::filter(inp, sheet=="earnings_debt"), all_of(excl)) %>%
      slice(2:n())
  }
) %>% 
  bind_rows(.) %>% 
  arrange(industry_group) %>%
  arrange(industry_group, desc(roic_cost_capital))

if (FALSE){
  singapore_earnings_debt %>%
    dplyr::filter(dividend_yield>0.05) %>%
    group_by(industry_group) %>%
    slice(1)
}
#----------------
excl=c("company_name", "exchange_ticker", "industry_group", "country", 
       "actual_debt_rating", "optimal_debt_rating", "flag_bankruptcy", 
       "flag_refinanced")

singapore_cost_capital=future_map(
  dat$dataset,
  function(inp){
    if ( (inp$sheet %>% n_distinct) > 1){
      extract.sheets.f(dplyr::filter(inp, sheet=="cost_capital"), all_of(excl))
    }
  }
) %>% 
  bind_rows(.) %>%
  mutate(spread_optimal=actual_debt_capital-optimal_debt_capital) %>% 
  arrange(industry_group, spread_optimal)

singapore_screener=singapore_earnings_debt %>%
  select(industry_group, company_name, dividend_yield, roe, cost_equity, roe_excess_return=roe_cost_equity, roic, cost_capital, roic_excess_return=roic_cost_capital) %>%
  left_join(
    singapore_cost_capital %>% select(company_name, actual_debt_capital, optimal_debt_capital, spread_optimal)
  ) %>%
  select(company_name, dividend_yield, roic_excess_return, roic, cost_capital, 
         roe_excess_return, roe, cost_equity, spread_optimal, actual_debt_capital, optimal_debt_capital) %>%
  dplyr::filter(dividend_yield>0.01, roic_excess_return>0.025, spread_optimal<0)

#-----------
if (FALSE){
  singapore_db=future_map(
    dat$dataset,
    function(inp){
      excl=c("country", "company_name", "industry_group")
      out=list(earnings_debt=extract.sheets.f(dplyr::filter(inp, sheet=="earnings_debt"), all_of(excl)))
      if ( (inp$sheet %>% n_distinct) > 1){
        excl=c("company_name", "exchange_ticker", "industry_group", "country", 
               "actual_debt_rating", "optimal_debt_rating", "flag_bankruptcy", 
               "flag_refinanced")
        out=update_list(
          out, 
          cost_capital=extract.sheets.f(dplyr::filter(inp, sheet=="cost_capital"), all_of(excl))
        )
        
        nm=inp %>%
          select(sheet) %>%
          distinct(.) %>%
          dplyr::filter(!sheet %in% c("earnings_debt", "cost_capital")) %>%
          pull
        
        out=update_list(
          out, 
          optimal_mix=map(
            nm,
            ~extract.sheets.f(dplyr::filter(inp, sheet==.x), NULL)
          ) %>%
            set_names(nm)
        )
      }
      out
    }
  )
  # listviewer::jsonedit(singapore_db)  
}

save(singapore_industries, singapore_earnings_debt, singapore_cost_capital, singapore_screener, file=file.path(path_data, "singapore_fundamental_data.Rda"))

rm(dat)
rm(singapore_industries, singapore_earnings_debt, singapore_cost_capital, singapore_screener)
rm(path_root, path_data, path_xlsx, list_paths)
rm(excl, extract.sheets.f)
