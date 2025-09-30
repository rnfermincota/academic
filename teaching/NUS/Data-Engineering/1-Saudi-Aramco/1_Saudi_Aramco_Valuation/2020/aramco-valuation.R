rm(list = ls(all = TRUE)); graphics.off() #; gc()
#----------------------------------------

library(knitr)
library(kableExtra)
library(ggpage)
library(magrittr)
library(tidyverse)
library(WDI)
library(pdftools)

prospectus.pdf=file.path("data", "saudi-aramco-prospectus-en.pdf")

#----------------------------------------
pdf.file=prospectus.pdf

# tabulizer::locate_areas(pdf.file, pages = 222, widget = "shiny")


page=220
names=c("Total borrowings", "Cash and cash equivalents", "Total equit")


require(tabulizer)
require(fuzzyjoin) # regex_inner_join
  
# https://www.saudiaramco.com/-/media/images/investors/saudi-aramco-prospectus-en.pdf
area = case_when( # tabulizer::locate_areas(f, pages = 222, widget = "shiny")
  page == 220  ~ c(459.77, 69.76, 601, 427.98), # Table 42 (pg 131)
  page == 221  ~ c(168.03, 69.76, 394.53, 404.59), # Table 43 (pg 132)
  page == 222  ~ c(180.11, 68.38, 413.04, 412.05), # Table 45 (pg 133)
  page == 233  ~ c(181.57, 70.99, 673.96, 448.91) # Table 52 (pg 144)
)

out1=extract_tables(
  pdf.file, pages = page, area = list(area), 
  guess = FALSE, output = "data.frame"
)

out2=out1 %>% purrr::pluck(1) 

out3=out2 %>%
  map_dfc(~trimws(gsub("\\.|[[:punct:]]", "", .x))) %>%
  set_names( c("Heading", paste0("X", if(page==233){1:4}else{0:4})) )

out4=out3 %>%
  regex_inner_join(
    data.frame(regex_name = names, stringsAsFactors = FALSE), 
    by = c(Heading = "regex_name")
  ) 

# Total borrowings, Cash and cash equivalents, Total equity:
out5=out4 %>%
  select(X4) %>% 
  pull %>% 
  as.numeric

rm(out1, out2, out3, out4, out5)
rm(area)
rm(pdf.file, page, names)

#----------------------------------------
# f <- file.path("data", "saudi-aramco-prospectus-en.pdf")
download.f <- function(url) {
  data.folder = file.path(getwd(), 'data')  # setup temp folder
  if (!dir.exists(data.folder)){dir.create(data.folder, F)}
  filename = file.path(data.folder, basename(url))
  if(!file.exists(filename))
    tryCatch({ download.file(url, filename, mode='wb') }, 
             error = function(ex) cat('', file=filename))
  filename
}

extract.values.f <- function(pdf.file, page, names){
  require(tabulizer)
  require(fuzzyjoin) # regex_inner_join
  
  # https://www.saudiaramco.com/-/media/images/investors/saudi-aramco-prospectus-en.pdf
  area = case_when( # tabulizer::locate_areas(f, pages = 222, widget = "shiny")
    page == 220  ~ c(459.77, 69.76, 601, 427.98), # Table 42 (pg 131)
    page == 221  ~ c(168.03, 69.76, 394.53, 404.59), # Table 43 (pg 132)
    page == 222  ~ c(180.11, 68.38, 413.04, 412.05), # Table 45 (pg 133)
    page == 233  ~ c(181.57, 70.99, 673.96, 448.91) # Table 52 (pg 144)
  )
  
  extract_tables(
    pdf.file, pages = page, area = list(area), 
    guess = FALSE, output = "data.frame"
  ) %>% 
    purrr::pluck(1) %>%
    map_dfc(~trimws(gsub("\\.|[[:punct:]]", "", .x))) %>%
    set_names( c("Heading", paste0("X", if(page==233){1:4}else{0:4})) ) %>%
    regex_inner_join(
      data.frame(regex_name = names, stringsAsFactors = FALSE), 
      by = c(Heading = "regex_name")
    ) %>%
    select(X4) %>% 
    pull %>% 
    as.numeric
}


inputs <- prospectus.pdf %>% 
  pdf_text() %>% read_lines() %>% 
  grep("proved reserves life", ., value = TRUE) %>% 
  str_match_all("[0-9]+") %>% 
  purrr::pluck(1) %>% 
  unlist %>% first() %>% as.numeric() %>% 
  set_names(c("LONG_RESERVES_LIFE")) %>% as.list

# Table 42 - Gearing and reconciliation
inputs <- extract.values.f(
  prospectus.pdf, 220, 
  c("Total borrowings", "Cash and cash equivalents", "Total equity")
) %>% 
  set_names(
    c("TOTAL_BORROWINGS", "CASH_AND_CASH_EQUIVALENTS", "TOTAL_EQUITY")
  ) %>%
  as.list %>% append(inputs)

# Table 43 - Return on Average Capital Employed (ROACE) and reconciliation
inputs <- extract.values.f(
  prospectus.pdf, 221, 
  c("Capital employed")
) %>% 
  last() %>% 
  set_names(c("CAPITAL_EMPLOYED")) %>%
  as.list %>% append(inputs)

# Table 45 - Income statement
inputs <- extract.values.f(
  prospectus.pdf, 222, 
  c("Operating income", "Income taxes", 
    "Income before income taxes", "Net income")
) %>% 
  set_names(c("OPERATING_INCOME", "INCOME_BEFORE_INCOME_TAXES", "INCOME_TAXES", "NET_INCOME")) %>%
  as.list %>% append(inputs)


# Table 52 - Balance sheet
inputs <- extract.values.f(
  prospectus.pdf, 233, 
  c("Shareholders equity", "Investment in joint ventures and associates", 
  "Investment in securities", "Noncontrolling interests")
) %>% 
  purrr::discard(is.na) %>% 
  set_names(c("INVESTMENT_JOINT_VENTURES_ASSOCIATES", "INVESTMENT_SECURITIES", 
              "SHAREHOLDERS_EQUITY", "NON_CONTROLLING_INTERESTS")) %>%
  as.list %>% append(inputs)

listviewer::jsonedit(inputs)


# World Development Indicators (WDI)
if (F){
  inputs <- WDI::WDI(
    country=c("SAU"), 
    indicator="NY.GDP.MKTP.KD.ZG", # = GDP growth (annual %)
    start=2018, 
    end=2018
  )$NY.GDP.MKTP.KD.ZG[[1]] %>%
    set_names(c("GDP_GROWTH")) %>% #  (annual %)
    as.list %>% append(inputs)
}

load(file="data/gdp_growth.rda")

inputs <- gdp_growth %>%
  set_names(c("GDP_GROWTH")) %>% #  (annual %)
  as.list %>% append(inputs)


if (F){
  treasury.rates.f <- function(year=2019){
    require(dplyr)
    require(quantmod)
    # year=calendar year to pull results for
    
    getSymbols.FRED("DGS10", env = environment())
    rates_raw <- na.locf(DGS10)
    
    rates_raw <- rates_raw[paste0(year, "/")] %>%
      timetk::tk_tbl(rename_index="date") %>%
      rename(trates=DGS10)
    
    # Returns treasury rates for the given duration
    rates <- rates_raw %>%
      # clean_names(.) %>%
      mutate(
        date = as.Date(date, "%m/%d/%y"),
        month = factor(months(date), levels=month.name)
      ) %>%
      mutate_at(
        vars(-one_of("date", "month")),
        as.numeric
      )
    
    summary <- rates %>%
      select(-date) %>%
      group_by(month) %>%
      summarise_all(list(mean))
    
    return(summary)
  }
  treasury_rates <- treasury.rates.f(2019)
  # save(treasury_rates, file="data/treasury_rates.rda")
  # load(file="data/treasury_rates.rda")
  
  inputs <- rates %>%
    select(x10_yr) %>%
    slice(n()) %>% # Dec 10_yr Avg.
    pull %>% 
    set_names(c("TREASURY_YIELD_10YR")) %>%
    as.list %>% append(inputs)
}

load(file="data/treasury_rates.rda")

inputs <- treasury_rates %>% 
  set_names(c("TREASURY_YIELD_10YR")) %>%
  as.list %>% append(inputs)


risk.premium.f <- function(){
  require(tidyxl) # It does not support the binary file formats '.xlsb' or '.xls'.
  
  # data_file <- file.path("data", "ctrypremJuly19.xlsx")
  url <- 'http://pages.stern.nyu.edu/~adamodar/pc/datasets/ctrypremJuly19.xlsx'
  data_file <- download.f(url)
  tidy_table <- xlsx_cells(data_file, sheets = "ERPs by country") %>% 
    filter(!is_blank, row >= 7 & row <=162) %>%
    select(row, col, data_type, character, numeric)
  # equity risk premium with a country risk premium for Saudi Arabia added to 
  # the mature market premium estimated for the US. 
  i <- tidy_table %>% filter(character=="Saudi Arabia") %>% pull(row)
  j <- tidy_table %>% filter(character=="Total Equity Risk Premium") %>% pull(col)
  # print(i)
  # print(j)
  v = tidy_table %>% filter(row == i & col == j) %>% pull(numeric)
  return(v * 100)
}

erp <- risk.premium.f()
inputs <- erp %>%
  set_names(c("EQUITY_RISK_PREMIUM")) %>%
  as.list %>% append(inputs)


rating.spread.f <- function(){
  require(readxl)
  # data_file <- file.path("data", "ratings.xls")
  # Ratings, Interest Coverage Ratios and Default Spread
  url <- 'http://www.stern.nyu.edu/~adamodar/pc/ratings.xls'
  data_file <- download.f(url)
  v <- read_excel(
    data_file, sheet = "Start here Ratings sheet", 
    range = "A18:D33") %>% # A18:D33 -> rating table for large manufacturing firms
    janitor::clean_names() %>%
    filter(rating_is=="A1/A+") %>%
    # https://www.bloomberg.com/news/articles/2019-04-01/saudi-oil-giant-aramco-starts-bond-roadshow-gets-a-rating
    pull(spread_is)
  return(v * 100)
}

cs <- rating.spread.f()
inputs <- cs %>%
  set_names(c("CREDIT_SPREAD")) %>%
  as.list %>% append(inputs)


unlevered.beta.f <- function(){
  require(readxl)
  # data_file <- file.path("data", "betaGlobal.xls")
  # Unlevered Betas (Global)
  url <- 'http://www.stern.nyu.edu/~adamodar/pc/datasets/betaGlobal.xls'
  data_file <- download.f(url)
  # A10:F106 -> Industry Name, Number of firms, Beta, D/E Ratio, 
  v <- read_excel(data_file, sheet = "Industry Averages", range = "A10:F106") %>%
    janitor::clean_names() %>%
    filter(industry_name=="Oil/Gas (Integrated)") %>%
    pull(unlevered_beta)
  return(v)
}

ub <- unlevered.beta.f()
inputs <- ub %>%
  set_names(c("UNLEVERED_BETA")) %>%
  as.list %>% append(inputs)

marginal.tax.f <- function(){
  require(readxl)
  # data_file <- file.path("data", "countrytaxrates.xls")
  url <- 'http://www.stern.nyu.edu/~adamodar/pc/datasets/countrytaxrates.xls'
  data_file <- download.f(url)
  # Corporate Marginal Tax Rates - By country
  v <- read_excel(data_file, sheet = "Sheet1") %>%
    janitor::clean_names() %>%
    filter(country=="Saudi Arabia") %>%
    pull(x2018)
  return(v * 100)
}

mtr <- marginal.tax.f()
inputs <- mtr %>%
  set_names(c("MARGINAL_TAX_RATE")) %>%
  as.list %>% append(inputs)

#----------------------------------------
# Valuation
equity.valuation.f <- function(inp){
  
  for (j in 1:length(inp)) assign(names(inp)[j], inp[[j]])
  #-------------------------------------------------------------------------------------
  # Calculated inputs
  
  EFFECTIVE_TAX_RATE <- INCOME_TAXES / INCOME_BEFORE_INCOME_TAXES
  INVESTED_CAPITAL <- CAPITAL_EMPLOYED - CASH_AND_CASH_EQUIVALENTS
  DEBT_RATIO <- TOTAL_BORROWINGS / ( TOTAL_BORROWINGS + TOTAL_EQUITY )
  
  COST_DEBT <- ( CREDIT_SPREAD + TREASURY_YIELD_10YR ) / 100
  COST_EQUITY <- ( TREASURY_YIELD_10YR + UNLEVERED_BETA * EQUITY_RISK_PREMIUM ) / 100
  COST_CAPITAL <- COST_DEBT * ( 1 - ( MARGINAL_TAX_RATE / 100 ) ) * DEBT_RATIO + 
    COST_EQUITY * ( 1 - DEBT_RATIO )
  
  NUMBER_YEARS <- LONG_RESERVES_LIFE
  
  #-------------------------------------------------------------------------------------
  # Free Cash Flow to Equity (FCFE)
  
  EXPECTED_RETURN_EQUITY <- NET_INCOME / SHAREHOLDERS_EQUITY
  EXPECTED_GROWTH_EARNINGS <- GDP_GROWTH / 100
  PAYOUT_RATIO <- 1 - EXPECTED_GROWTH_EARNINGS / EXPECTED_RETURN_EQUITY
  
  VALUE_EQUITY <- NET_INCOME * PAYOUT_RATIO * 
    ( 1 - ( ( 1 + EXPECTED_GROWTH_EARNINGS ) ^ NUMBER_YEARS / 
              ( 1 + COST_EQUITY ) ^ NUMBER_YEARS ) ) / 
    ( COST_EQUITY - EXPECTED_GROWTH_EARNINGS )
  
  FCFE_EQUITY_VALUATION <- VALUE_EQUITY + CASH_AND_CASH_EQUIVALENTS + 
    INVESTMENT_JOINT_VENTURES_ASSOCIATES + INVESTMENT_SECURITIES
  
  #-------------------------------------------------------------------------------------
  # Free Cash Flow to Firm (FCFF)
  EXPECTED_GROWTH_RATE <- GDP_GROWTH / 100
  EXPECTED_ROIC <- OPERATING_INCOME * ( 1 - EFFECTIVE_TAX_RATE ) / INVESTED_CAPITAL
  REINVESTMENT_RATE <- EXPECTED_GROWTH_RATE / EXPECTED_ROIC
  
  EXPECTED_OPERATING_INCOME_AFTER_TAX <- OPERATING_INCOME * 
    ( 1 - EFFECTIVE_TAX_RATE ) * ( 1 + EXPECTED_GROWTH_RATE )
  
  EXPECTED_FCFF <- EXPECTED_OPERATING_INCOME_AFTER_TAX * ( 1 - REINVESTMENT_RATE )
  
  VALUE_OPERATING_ASSETS <- EXPECTED_FCFF * 
    ( 1 - ( ( 1 + EXPECTED_GROWTH_RATE ) ^ NUMBER_YEARS / 
              ( 1 + COST_CAPITAL ) ^ NUMBER_YEARS ) ) / 
    ( COST_CAPITAL - EXPECTED_GROWTH_RATE )
  
  FCFF_EQUITY_VALUATION <- VALUE_OPERATING_ASSETS + CASH_AND_CASH_EQUIVALENTS + 
    INVESTMENT_JOINT_VENTURES_ASSOCIATES + INVESTMENT_SECURITIES - 
    TOTAL_BORROWINGS - NON_CONTROLLING_INTERESTS
  
  #-------------------------------------------------------------------------------------
  # Use set_names to name the elements of the vector
  out <- c(INVESTED_CAPITAL, DEBT_RATIO, EFFECTIVE_TAX_RATE) %>% 
    set_names(c("INVESTED_CAPITAL", "DEBT_RATIO", "EFFECTIVE_TAX_RATE"))
  
  out <- c(NUMBER_YEARS, COST_CAPITAL, COST_EQUITY, COST_DEBT) %>% 
    set_names(c("NUMBER_YEARS", "COST_CAPITAL", "COST_EQUITY", "COST_DEBT")) %>%
    as.list %>% append(out)
  
  out <- c(FCFE_EQUITY_VALUATION, VALUE_EQUITY, PAYOUT_RATIO, 
           EXPECTED_GROWTH_EARNINGS, EXPECTED_RETURN_EQUITY) %>% 
    set_names(c("FCFE_EQUITY_VALUATION", "VALUE_EQUITY", "PAYOUT_RATIO", 
                "EXPECTED_GROWTH_EARNINGS", "EXPECTED_RETURN_EQUITY")) %>%
    as.list %>% append(out)
  
  out <- c(FCFF_EQUITY_VALUATION, VALUE_OPERATING_ASSETS, EXPECTED_FCFF, 
           EXPECTED_OPERATING_INCOME_AFTER_TAX, REINVESTMENT_RATE, 
           EXPECTED_ROIC, EXPECTED_GROWTH_RATE) %>% 
    set_names(c("FCFF_EQUITY_VALUATION", "VALUE_OPERATING_ASSETS", "EXPECTED_FCFF", 
                "EXPECTED_OPERATING_INCOME_AFTER_TAX", "REINVESTMENT_RATE", 
                "EXPECTED_ROIC", "EXPECTED_GROWTH_RATE")) %>%
    as.list %>% append(out)
  #-------------------------------------------------------------------------------------
  
  return(out)
}

output <- equity.valuation.f(inputs)

# listviewer::jsonedit(output)

data.frame(
  Weighted = 0.5 * (output$FCFF_EQUITY_VALUATION + output$FCFE_EQUITY_VALUATION) / 1000000,
  FCFF = output$FCFF_EQUITY_VALUATION / 1000000,
  FCFE = output$FCFE_EQUITY_VALUATION / 1000000,
  check.names = FALSE
) %>%
  mutate_all(scales::dollar) %>% 
  kable() %>%
  kable_styling(c("striped", "bordered")) %>%
  add_header_above(c("Saudi Aramco Equity Valuation ($ trillions)" = 3))
#-------------------------------------------------------------------------------------
# Risk Premium

# Equity Risk Premium
out <- map(
  seq(6, 10, 0.25), 
  ~list_modify(
    inputs, 
    EQUITY_RISK_PREMIUM=.x
  ) %>% 
    equity.valuation.f(.) 
)

map2_dfr(
  out, 
  seq(6, 10, 0.25),
  ~list(
    EQUITY_RISK_PREMIUM=.y, 
    COST_CAPITAL=.x$COST_CAPITAL*100,
    WEIGHTED=(.x$FCFF_EQUITY_VALUATION+.x$FCFE_EQUITY_VALUATION) / 2 / 1000000,
    FCFF=.x$FCFF_EQUITY_VALUATION / 1000000, 
    FCFE=.x$FCFE_EQUITY_VALUATION / 1000000
  )
) %>%
  # arrange(-EQUITY_RISK_PREMIUM) %>%
  mutate_at(
    vars(one_of("FCFF", "FCFE", "WEIGHTED")),
    scales::dollar
  ) %>% 
  mutate_at(
    vars(one_of("EQUITY_RISK_PREMIUM", "COST_CAPITAL")),
    function(v) sprintf(v, fmt = "%.2f%%")
  ) %>%
  rmarkdown::paged_table()

#-------------------------------------------------------------------------------------
# Treasury
out <- map(
  seq(1, 4, 0.25), 
  ~list_modify(
    inputs, 
    TREASURY_YIELD_10YR=.x
  ) %>% 
    equity.valuation.f(.)
)

map2_dfr(
  out, 
  seq(1, 4, 0.25),
  ~list(
    TREASURY_YIELD_10YR=.y, 
    WEIGHTED=(.x$FCFF_EQUITY_VALUATION+.x$FCFE_EQUITY_VALUATION) / 2 / 1000000,
    FCFF=.x$FCFF_EQUITY_VALUATION / 1000000, 
    FCFE=.x$FCFE_EQUITY_VALUATION / 1000000
  )
) %>%
  # arrange(-TREASURY_YIELD_10YR) %>%
  mutate_at(
    vars(one_of("FCFF", "FCFE", "WEIGHTED")),
    scales::dollar
  ) %>% 
  mutate_at(
    vars(one_of("TREASURY_YIELD_10YR")),
    function(v) sprintf(v, fmt = "%.2f%%")
  ) %>%
  rmarkdown::paged_table()

#-------------------------------------------------------------------------------------
# Reserve Life

out <- map(
  40:52, # Long reserves life
  ~list_modify(
    inputs, 
    LONG_RESERVES_LIFE=.x
  ) %>% 
    equity.valuation.f(.)
)

map_dfr(
  out, 
  ~list(
    RESERVES_LIFE=.x$NUMBER_YEARS, 
    WEIGHTED=(.x$FCFF_EQUITY_VALUATION+.x$FCFE_EQUITY_VALUATION) / 2 / 1000000,
    FCFF=.x$FCFF_EQUITY_VALUATION / 1000000, 
    FCFE=.x$FCFE_EQUITY_VALUATION / 1000000
  )
) %>% 
  arrange(-RESERVES_LIFE) %>%
  mutate_at(
    vars(one_of("FCFF", "FCFE", "WEIGHTED")),
    scales::dollar
  ) %>% 
  rmarkdown::paged_table()

