#' @title Outputs tbl that gives descriptive information about the tables and their contents
#' @description Outputs tbl that gives descriptive information about the tables and their contents
#' @param con
#' A sqlite database connection object
#' @return A tbl of database info
#' @importFrom magrittr %>%
#' @export

f1db_schema_info <- function(con){

    table_names <-  c("constructors","constructor_standings",
                      "constructor_results", "circuits", "drivers",
                      "driver_standings", "lap_times", "pit_stops",
                      "seasons", "status", "races", "results", "qualifying")

    for(i in 1:length(table_names)){
        assign(table_names[i], tbl(con, table_names[i]) %>% dplyr::collect())
    }


    schema_info <- data.frame(endpoint = table_names,
                               endpoint_title = c("teams","teams leaderboard",
                                                  "team results", "circuits",
                                                  "drivers", "driver leaderboard",
                                                  "lap times", "pit stops", "years",
                                                  "status", "races", "results",
                                                  "starting order"),
                               endpoint_description = c(
                                   "information about the teams",
                                   "points and position by team per race",
                                   "results by team per race",
                                   "tracks being raced on",
                                   "driver information",
                                   "points and position by driver",
                                   "lap time and position by driver and race",
                                   "pitstop time, lap, stop number by driver and race",
                                   "year and URL link to wiki page for each season",
                                   "status ID codes and their meanings",
                                   "race information by year, circuit, name and date",
                                   "result information by race, driver and results",
                                   "qualifying results by q1-3, by driver, race, constructor"),
                               properties = c(
                                   prop_function(constructors),
                                   prop_function(constructor_standings),
                                   prop_function(constructor_results),
                                   prop_function(circuits),
                                   prop_function(drivers),
                                   prop_function(driver_standings),
                                   prop_function(lap_times),
                                   prop_function(pit_stops),
                                   prop_function(seasons),
                                   prop_function(status),
                                   prop_function(races),
                                   prop_function(results),
                                   prop_function(qualifying)

                               ))  %>% as_tibble()
    return(schema_info)
}

#' @title Helper function used inside schema_info
#' @description Helper function used inside schema_info
#' @param t value indicating dimensions
#' @return dimension string
prop_function <- function(t){
    n_cols <- length(t)
    n_rows <- nrow(t)
    property = paste0("# A tibble: ", n_rows, " x ", n_cols)
    property
}
