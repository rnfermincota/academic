#-----------------------------------------------------------------------------------------
# Read Me First: https://rpubs.com/rafael_nicolas/saudi_arabia_extracting_data
#-----------------------------------------------------------------------------------------
# Cleaning the R Global Environment
# We begin by first making sure that we delete all objects in the environment and from 
# the memory to prevent it from being too full.
gc(); rm(list=ls()); graphics.off()

#----------------
# Load Data Source
# Specify the working file path containing the working RMD. file.
# Construct the directory to the data folder and then to the xlsx folder containing 
# the excel files.
path_root <- "."
path_output <- file.path(path_root, "output") 
path_root <- ".."
path_data <- file.path(path_root, "data") 
# https://github.com/rnfermincota/academic/tree/main/research/traditional_assets/database/data/saudi_arabia

#----------------
# Loading Necessary Libraries
# These libraries will be required for all following code below. Explanation of relevant 
# libraries included as well.
library(tidyxl) # Imports non-tabular data from Excel files and exposes cell contents 
# in a tidy structure for manipulation
library(tibble) # Modern take on data frames and encapsulates best practices for them
library(dplyr) # Grammar of data manipulation
library(purrr) # Enhances functional programming practices by allowing us to apply 
# functions for iterations
library(furrr) # Complements the purrr package to parallelize computations easily
plan(multisession) # Uses multicore evaluation whereby values are computed and resolved in 
# parallel
library(unpivotr) # Converts data from complex or irregular layouts to a columnar structure
# https://nacnudus.github.io/unpivotr/

#----------------
# Code Explanations
# In the following code chunk, we are writing a function called extract.sheets.f with 
# the inputs to be the excel sheets. For the input excel sheet, we select only the row, 
# col, data_type, character, numeric, logical columns and we rectify them.
# The rectify function takes the "melted" output of the cells of the data and projects 
# them into their original positions. In that projected output, we then select everything 
# except the row/col column. We then take that output and slice the first row out and 
# simplify it to produce a vector with all the components in the first row (flattening).
# We set names for the output using nm as the vector of names and label the dataset 
# with indexes as the headings. We also slice the data from the second row onward till 
# the last row, leaving out the header row. The function also checks that if the columns 
# country, company_name and industry_group are not null, we transform across all 
# the other columns excluding country, company_name and industry_group by forcing 
# them to be numeric data type. Finally, we return the fully extracted data in out.
extract.sheets.f <- function(dsheet, excl=NULL){
  # inp=dat[1,]$dataset[[1]]
  # dsheet=filter(inp, sheet=="cost_capital" )
  out <- dsheet %>% 
    select(row, col, data_type, character, numeric, logical) %>% 
    rectify %>% # https://github.com/nacnudus/unpivotr/blob/main/R/rectify.R
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
## Determine file paths in directory
# We utilize the list.files function to list all the files in the identified directory
# with extension .xlxs in their full names and save them into the list_paths variable.
# The directory path is also prepended to the file names directly. Essentially, we are 
# listing all Microsoft Excel data in the directory.

list_paths <- list.files(
  path = file.path(path_data),
  pattern = ".xlsx",
  full.names = TRUE
)

#----------------
## Creation of main dataframe with all datasets

# We utilize the enframe function to convert the vector list_paths, the path of where 
# the data sets are being stored, into a dataframe with a single column name as path.
# We also create a new column, sector, with only the sector names by replacing the 
# ".xlsx" extension with blank from the vector of base names of all excel files in 
# the path. Next, we further create a new column called dataset. This is created using 
# the future_map function whereby we are mapping the xlsx_cells function to each dataset
# in the path, while allowing us to map in parallel. For each dataset, the xlsx_cells 
# function is importing data from each spreadsheet into a tidy structure. This means 
# that each cell will be represented by a row in a data frame. Thus, in the dataset column
# we have each data frame from each spreadsheet in the corresponding cell. We then assign 
# names to the column corresponding to the respective sector name from the sector variable.

df_paths=enframe(list_paths, name = NULL, value = "path") %>%
  mutate(sector=gsub(".xlsx", "", basename(path)))

dat=df_paths %>% # slice(1:2) %>%
  mutate(
    dataset=future_map(path, xlsx_cells) %>% 
      set_names(sector)
  )

#---------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------
## Creation of validation lists

lst_xlsx=future_map(
  df_paths$path,
  function(x) {
    nms=excel_sheets(x) # %>% str_subset('earnings_debt|cost_capital', negate = TRUE)
    lst=map(nms, function(s) {
      read_excel(x, sheet=s) %>%
        mutate(across(where(function(x) is.logical(x) & all(is.na(x))), ~NA_real_))
    })
    lst=set_names(lst, nms)
    return(lst)
  }
) %>% 
  set_names(sub('\\.xlsx$', '', basename(df_paths$path)))

lst_earnings_debt=map(lst_xlsx, ~.x$earnings_debt)
lst_cost_capital=map(lst_xlsx, ~.x$cost_capital) %>% discard(is_empty) # summary(map_lgl(lst_cost_capital, is_empty))
lst_optimal_mix=map(lst_xlsx, ~keep(.x, !str_detect(names(.x), 'earnings_debt|cost_capital'))) %>% discard(is_empty)

if(F){
  bind_rows(lst_earnings_debt, .id="country_sector")
  bind_rows(lst_cost_capital, .id="country_sector")
  lst_optimal_mix %>% 
    map(., ~bind_rows(.x, .id="company")) %>%
    bind_rows(.id="country_sector")
}

#----------------
## Overall Earnings Debt of Saudi Arabia Industries

# We create the variable excl containing the column names we would like to exclude. This
# variable indicates the columns in the dataset that are of character data type.

# For each dataset in the data frame, we are using the future_map function to map each 
# dataset to the extract.sheets.f function we created earlier. In the function we are 
# mapping, we filter the input dataset by only retrieving the sheet name that is equal 
# to earnings_debt, while also satisfying the input excl. The input excl are the columns 
# of character datatype and are the columns we are excluding when forcing the other 
# columns to become numeric data type. We then slice only the first row of the 
# intermediary data output.

# Next, we then bind all the rows together for each dataframe from each sector and first 
# arrange them by industry_group in ascending order (alphabetically). We then further 
# arrange them by roic_cost_capital in descending order. The resulting saudi_arabia_industries 
# output consists of a table of financial earnings and debt by overall industry in 
# Saudi Arabia.

excl=c("country", "company_name", "industry_group")
saudi_arabia_industries=future_map(
  dat$dataset,
  function(inp){
    extract.sheets.f(dplyr::filter(inp, sheet=="earnings_debt"), all_of(excl)) %>%
      slice(1)
  }
)

# cell_grd [1 × 43] (S3: cell_grid/tbl_df/tbl/data.frame)
# tibble [1 × 43] (S3: tbl_df/tbl/data.frame)
all.equal(saudi_arabia_industries, map(lst_earnings_debt, ~.x %>% slice(1)), check.attributes=F)

saudi_arabia_industries=saudi_arabia_industries %>% 
  bind_rows(.) %>% 
  arrange(industry_group) %>%
  arrange(desc(roic_cost_capital))

#----------------
## Earnings Debt of Saudi Arabian Companies

# In the next code chunk, we are performing a similar step to the previous code chunk.

# For each dataset in the data frame, we are using the future_map function to map each 
# dataset to the extract.sheets.f function we created earlier. In the function we are 
# mapping, we filter the input dataset by only retrieving the sheet name that is equal 
# to earnings_debt. We indicate the excl input variable as the columns that contain 
# character data type, from which the function will exclude them when forcing the
# other columns to become numeric data type.

# In this step however, we will be slicing data from the second row onward till the 
# respective last row for each dataset. The table output for this step contains the 
# financial earnings and debt by each individual company name and its respective 
# industry_group in Saudi Arabia.

# We then bind all the rows for each dataframe and first arrange them by industry_group
# in ascending order. Then further arrange them by industry_group in ascending order 
# (alphabetically) and roic_cost_capital in descending order within each window of 
# industry_group.

saudi_arabia_earnings_debt=future_map(
  dat$dataset,
  function(inp){
    extract.sheets.f(dplyr::filter(inp, sheet=="earnings_debt"), all_of(excl)) %>%
      slice(2:n())
  }
) 

all.equal(saudi_arabia_earnings_debt, map(lst_earnings_debt, ~.x %>% slice(2:n())), check.attributes=F)

saudi_arabia_earnings_debt=saudi_arabia_earnings_debt %>% 
  bind_rows(.) %>% 
  arrange(industry_group) %>%
  arrange(industry_group, desc(roic_cost_capital))

#----------------
## Earnings Debt of Saudi Arabian Companies with more than 5% dividend yield

# In this step, we are taking all earnings and debt information from 
# saudi_arabia_earnings_debt, filtering and only keeping those companies with
# dividend yield of more than 0.05 or more than 5%. We then group them by their 
# industry and only select the company belonging to the first row among each respective 
# industry group.

if (FALSE){
  saudi_arabia_earnings_debt %>%
    dplyr::filter(dividend_yield>0.05) %>%
    group_by(industry_group) %>%
    slice(1)
}

#----------------
## Cost Capital of Saudi Arabian companies

# In the following code chunk, we re-create the variable excl containing the column 
# names we would like to exclude. This variable indicates the columns in the dataset
# that are of character and boolean data type. Essentially, we are excluding any 
# columns that are of non-numeric data type.

# For each dataset in the dataframe, if there are more than 1 unique sheets in each 
# dataset, the future_map function will be used to map each dataset to the 
# extract.sheets.f function we created at the beginning. In the function we are mapping,
# we filter the input dataset by only retrieving the sheet name that is equal to 
# cost_capital, while also satisfying the input excl. The input excl are the columns of
# character and boolean data type and we are excluding them when forcing the other 
# columns to become numeric data type. 

# Next, we then bind all the rows together for each dataframe from each sector. We 
# create a new column called spread_optimal which is the difference between 
# current_debt_capital and the optimal_debt_capital. Then, we first arrange the data 
# by industry_group in ascending order (Alphabetically) and then by spread_optimal 
# in ascending order among each window of industry group. The resulting 
# saudi_arabian_cost_capital output consists of a final table of cost capital by companies 
# in Saudi Arabia.

excl=c("company_name", "exchange_ticker", "industry_group", "country", 
       "reported_debt_rating", "current_debt_rating", "optimal_debt_rating", 
       "flag_bankruptcy", "flag_refinanced")

saudi_arabia_cost_capital=future_map(
  dat$dataset,
  function(inp){
    if ( (inp$sheet %>% n_distinct) > 1){
      extract.sheets.f(dplyr::filter(inp, sheet=="cost_capital"), all_of(excl)) %>%
        # mutate(across(where(is.factor), as.character)) %>%
        # modify_if(is.factor, as.character) 
        mutate(
          #current_debt_ebitda=as.numeric(current_debt_ebitda),
          #optimal_debt_ebitda=as.numeric(optimal_debt_ebitda),
          flag_bankruptcy=as.logical(flag_bankruptcy),
          flag_refinanced=as.logical(flag_refinanced)
        )
    }
  }
)

all.equal(saudi_arabia_cost_capital %>% discard(is.null), map(lst_cost_capital, ~.x), check.attributes=F)

saudi_arabia_cost_capital=saudi_arabia_cost_capital %>% 
  bind_rows(.) %>%
  mutate(spread_optimal=current_debt_capital-optimal_debt_capital) %>% 
  arrange(industry_group, spread_optimal)

#----------------
## Screen Saudi Arabian companies

# Moving on to the next chunk of code, in order to build the saudi_arabia_screener 
# dataframe, we first take the saudi_arabia_earnings_debt dataframe and only select 
# the respective columns: industry_group, company_name, dividend_yield, roe,
# cost_equity, roe_cost_equity, roic, cost_capital, roic_cost_capital. In this 
# selection, we also rename roe_cost_equity as roe_excess_return and rename
# roic_cost_capital as roic_excess_return.

# We then perform a left join with the saudi_arabia_cost_capital dataframe on the 
# company_name as the variable to join by. From the saudi_arabia_cost_capital dataframe
# during the join, we are only selecting the columns company_name, current_debt_capital,
# optimal_debt_capital and spread_optimal.

# The resulting joined dataframe consists of 528 rows and 12 columns.

# From the joined dataframe, we the further select only the respective columns: 
# company_name, dividend_yield, roic_excess_return, roic, cost_capital, roe_excess_return, 
# roe, cost_equity, spread_optimal, current_debt_capital, optimal_debt_capital. 
# This table now has 11 columns.

# Finally, we filter the resulting table with companies that have a dividend yield of 
# more than 0.01, companies with a roic excess return of more than 0.025 and companies 
# with a negative spread optimal (less than 0). The filter on the rows are based on 
# these 3 conditions. The final resulting table (saudi_arabia_screener) has 52 rows and 
# 11 columns.
  
saudi_arabia_screener=saudi_arabia_earnings_debt %>%
  select(industry_group, company_name, dividend_yield, roe, cost_equity, 
         roe_excess_return=roe_cost_equity, roic, cost_capital, 
         roic_excess_return=roic_cost_capital) %>%
  left_join(
    saudi_arabia_cost_capital %>% 
      select(company_name, current_debt_capital, optimal_debt_capital, spread_optimal),
    by = join_by(company_name)
  ) %>%
  select(company_name, dividend_yield, roic_excess_return, roic, cost_capital, 
         roe_excess_return, roe, cost_equity, spread_optimal, current_debt_capital, 
         optimal_debt_capital) %>%
  dplyr::filter(dividend_yield>0.01, roic_excess_return>0.025, spread_optimal<0)

#-----------
## Combining all Sheets to Large List

# In the next code chunk, for each dataset in the original dat data frame, we are using 
# the future_map function to map each dataset to the complex function we will be creating.
# In the function we are mapping, we first indicate the excl input variable as the 
# columns that contain character data type, from which the extract.sheets.f function we 
# created earlier will exclude these variables when forcing the other columns
# to become numeric data type.

# We then create a list of dataframes whereby the extract.sheets.f function will extract
# the sheet from each dataset and we filter the input dataset by only retrieving the 
# sheet name that is equal to earnings_debt. 

# For each dataset in the dataframe, if there are more than 1 unique sheets in each 
# dataset, we re-define the  excl variable to contain columns of character and 
# boolean data type and we are excluding them when forcing the other columns to 
# become numeric data type.

# We update the list created earlier in out and re-define it by including another 
# input in the list called cost_capital. In order to get the cost_capital, we apply 
# the extract.sheets.f function and only filter the input dataset by only retrieving
# the sheet name that is equal to cost_capital, while also satisfying the input excl 
# variable.

# Next, for each input dataframe, we select the sheets available and obtain the 
# distinct sheet names. We then filter the sheets to only retrieve
# the sheets that are NOT earnings_debt and cost_capital and extract it out.

# We update the list again created earlier in out and re-define it by including 
# another input in the list called optimal_mix. In order to generate the optimal_mix,
# we are creating a formula in a map function with our x variables that are sheets
# other than earnings_debt and cost_capital, specified in the nm variable. In the 
# formula, we re-apply the extract.sheets.f function but this time round, we filter
# the input dataset to retrieve sheet names excluding earnings_debt and cost_capital.
# We also set the names of each vector in the large list corresponding to the sheet
# names extracted in the nm variable.

# Finally, we return the complete out variable which is a Large List of of 73 elements,
# each corresponding to an industry group in Saudi Arabia. In each element in the large 
# list, we contain a list of earnings_debt or 3 lists namely, earnings_debt, 
# cost_capital and optimal_mix.

# The final jsonedit function provides an output of the Large List that is flexible 
# and interactive in a tree-like view of lists. The final output, saudi_arabia_db, is 
# below.

saudi_arabia_db=future_map(
  dat$dataset,
  function(inp){
    excl=c("country", "company_name", "industry_group")
    out=list(
      earnings_debt=extract.sheets.f(dplyr::filter(inp, sheet=="earnings_debt"), all_of(excl))
    )
    if ( (inp$sheet %>% n_distinct) > 1){
      excl=c("company_name", "exchange_ticker", "industry_group", "country", 
             "reported_debt_rating", "current_debt_rating", "optimal_debt_rating", 
             "flag_bankruptcy", "flag_refinanced")
      out=list_modify(
        out, 
        cost_capital=extract.sheets.f(dplyr::filter(inp, sheet=="cost_capital"), all_of(excl)) %>%
          mutate(
            flag_bankruptcy=as.logical(flag_bankruptcy),
            flag_refinanced=as.logical(flag_refinanced)
          )
      )
      
      nm=inp %>%
        select(sheet) %>%
        distinct(.) %>%
        dplyr::filter(!sheet %in% c("earnings_debt", "cost_capital")) %>%
        pull
      
      out=list_modify(
        out, 
        optimal_mix=map(
          nm,
          ~extract.sheets.f(dplyr::filter(inp, sheet==.x), NULL) %>%
            mutate(across(c(where(is.character), -debt_rating), as.numeric))
        ) %>%
          set_names(nm)
      )
    }
    out
  }
)

all.equal(
  saudi_arabia_db,
  future_map( # saudi_arabia_db1
    lst_xlsx,
    function(x){
      list(
        earnings_debt=x$earnings_debt,
        cost_capital=x$cost_capital,
        optimal_mix=keep(x, !str_detect(names(x), 'earnings_debt|cost_capital'))
      ) %>% discard(is_empty)
    }
  ),
  check.attributes = F
)
  
# listviewer::jsonedit(saudi_arabia_db)  

#-----------
## Saving all Variables
# Finally, we save all dataframes created in the process into an .Rda file.
# And clear all objects from the current environment to preserve memory in Rstudio.

save(saudi_arabia_industries, saudi_arabia_earnings_debt, saudi_arabia_cost_capital, 
     saudi_arabia_screener, file=file.path(path_output, "saudi_arabia_fundamental_data.Rda"))


#---------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------

rm(path_data, path_output, path_root)
rm(list_paths, df_paths, dat)
rm(lst_xlsx, lst_earnings_debt, lst_cost_capital, lst_optimal_mix)
rm(excl, saudi_arabia_industries, saudi_arabia_earnings_debt, saudi_arabia_cost_capital)
rm(extract.sheets.f)
