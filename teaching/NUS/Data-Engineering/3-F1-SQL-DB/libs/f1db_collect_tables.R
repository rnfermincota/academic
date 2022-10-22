#' @title Save all f1db tables as global variable tbl's with the same name
#' @description Save all f1db tables as global variable tbl's with the same name
#' @param con A SQLite database connection object
#'
#' @return Saves all 13 Formula One tables as global objects
#' @export
#' @importFrom magrittr %>%
#' @examples \dontrun{f1db_collect_tables(con)}

f1db_collect_tables <- function(con){
    constructors          <<- dplyr::tbl(con, "constructors") %>% dplyr::collect()
    constructor_standings <<- dplyr::tbl(con, "constructor_standings") %>% dplyr::collect()
    constructor_results   <<- dplyr::tbl(con, "constructor_results") %>% dplyr::collect()
    circuits              <<- dplyr::tbl(con, "circuits") %>% dplyr::collect()
    drivers               <<- dplyr::tbl(con, "drivers") %>% dplyr::collect()
    driver_standings      <<- dplyr::tbl(con, "driver_standings") %>% dplyr::collect()
    lap_times             <<- dplyr::tbl(con, "lap_times") %>% dplyr::collect()
    pit_stops             <<- dplyr::tbl(con, "pit_stops") %>% dplyr::collect()
    seasons               <<- dplyr::tbl(con, "seasons") %>% dplyr::collect()
    status                <<- dplyr::tbl(con, "status") %>% dplyr::collect()
    races                 <<- dplyr::tbl(con, "races") %>% dplyr::collect()
    results               <<- dplyr::tbl(con, "results") %>% dplyr::collect()
    qualifying            <<- dplyr::tbl(con, "qualifying") %>% dplyr::collect()

}
