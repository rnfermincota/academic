rm(list = ls(all = TRUE)); gc(); graphics.off()
#----------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------
library(dplyr)
#-----------------------------------------------------------
#-------------------------------------------------------------------
path_root=".."
path_libs=file.path(path_root, "libs")
#-------------------------------------------------------------------

f1db <- new.env() #rlang::new_environment() # 
list.files(path = path_libs, pattern = "*.R") %>% 
  file.path(path_libs, .) %>% 
  # purrr::map(pryr::partial(source, local=f1db), encoding = 'UTF8') %>% 
  purrr::map(source, local=f1db, encoding = 'UTF8') %>% 
  invisible()

nm <- ls(env=f1db)
# fl <- purrr::map(nm, pryr::partial(get, envir=f1db)) %>% 
fl <- purrr::map(nm, get, envir=f1db) %>% setNames(nm)
fl %>% attach()

#----------------------------------------------------------------------------------------------------------

path_root="." # Download configure and connect Ergast files to SQLite database
con <- f1db_connect() # remeber to call dbDisconnect() when finished working with a connection!
f1db_collect_tables(con)

#----------------------------------------------------------------------------------------------------------
# Lowest grid position to a win a race?
#-----------------------------------------------------------

# Joins & column selection
race_results <- drivers %>% 
    left_join(results, by = c("driverId" = "driverId")) %>% 
    left_join(races, by = c("raceId" = "raceId")) %>% 
    mutate(driver = paste0(forename, " ", surname)) %>% 
    select(date, name, driver, grid, positionOrder) 

#-----------------------------------------------------------
# List of current drivers in 2022
current_drivers <- race_results %>% 
    filter(date %>% str_starts("2022")) %>% 
    select(driver) %>% 
    distinct()

# Results 
race_results %>% 
    filter(driver %in% current_drivers[[1]]) %>% 
    filter(positionOrder == 1) %>% 
    arrange(desc(grid)) %>% 
    head(1)

# A tibble: 1 × 5
# date       name                 driver           grid positionOrder
# <chr>      <chr>                <chr>           <int>         <int>
# 1 2008-09-28 Singapore Grand Prix Fernando Alonso    15             1

#-----------------------------------------------------------
# List of current drivers 1950 to today
current_drivers <- race_results %>% 
  select(driver) %>% 
  distinct()

# Results 
race_results %>% 
  filter(driver %in% current_drivers[[1]]) %>% 
  filter(positionOrder == 1) %>% 
  arrange(desc(grid)) %>% 
  head(1)
# A tibble: 1 × 5
# date       name                          driver       grid positionOrder
# <chr>      <chr>                         <chr>       <int>         <int>
# 1 1983-03-27 United States Grand Prix West John Watson    22             1

#-----------------------------------------------------------
dbDisconnect(con)
fl %>% detach()
rm(fl, nm)
file.remove(file.path(path_root, "f1_db.SQLite"))
unlink(file.path(path_root, "f1db_csv"), recursive = TRUE) 
rm(con, f1db)
rm(path_libs, path_root)

rm(circuits, constructor_results, constructor_standings, constructors)
rm(current_drivers, driver_standings, drivers, lap_times, pit_stops)
rm(qualifying, race_results, races, results, seasons, status)
