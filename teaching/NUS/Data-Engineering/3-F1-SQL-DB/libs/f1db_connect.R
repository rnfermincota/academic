#' @title Downloads configures connect to database
#' @description Downloads configures connect to database
#' @param collect_tables
#' Boolean Value for additionally executing the f1db_collect_tables(con) function
#' @return A SQLite database connection object
#' @export
#' @examples \dontrun{con <- f1db_connect()}
#'
f1db_connect <- function(collect_tables = FALSE){

    # Load packages, install if not available
    install_load_packages("DBI", "dm", "RSQLite",
                          "tidyverse", "lubridate",
                          "rlang", "datamodelr")

    # Database connection
    if(file.exists("f1_db.sqlite")){
        f1db_cleardb()
    }
    con_dm <- createF1db()
    con <- con_dm[[1]]
    dm_obj <<- con_dm[[2]]

    # Save all tables as global variable tbl's with the same name
    if(collect_tables == TRUE){
        f1db_collect_tables(con)
    }
    check_SQL_duplicates(con, dm_obj)


    return(con)

}
