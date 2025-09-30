#' @title Returns dm object for f1db database, for internal use only by f1db_connect
#' @description Returns dm object for f1db database, for internal use only by f1db_connect
#' @return dm_obj
#' @export
#'
f1db_get_dm <- function(){
    return(dm_obj)
}

#' @title Disconnects from database
#' @description Disconnects from database
#' @param con A SQLite database connection object
#' @param shutdown Boolean value for additionally shutting down the database connection
#' @export
#'
f1db_disconnect <- function(con, shutdown = TRUE){
    DBI::dbDisconnect(con, shutdown = shutdown)
}


# Reconnects to database, currently does not work
# f1db_reconnect <- function(){
#     con <- DBI::dbConnect(duckdb("f1_db.duckdb", read_only = TRUE))
#     return(con)
#
#}


#' @title Deletes database files of the exist, for internal use only by f1db_connect
#' @description Deletes database files of the exist, for internal use only by f1db_connect
#' @export
#'
f1db_cleardb <- function(){
    if(file.exists("f1_db.sqlite")){
        file.remove("f1_db.sqlite")
    }
    if(file.exists("f1_db.sqlite.wal")){
        file.remove("f1_db.sqlite.wal")
    }

}
