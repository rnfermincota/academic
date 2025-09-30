#' @title Checks database for duplicate entries
#' @description Checks database for duplicate entries
#' @param dm_obj a dm object
#' @param con a database connection object
#' @export
check_SQL_duplicates <- function(con, dm_obj){

    # Get table names, primary keys and dimensions
    tables <- dm::dm_get_tables(dm_obj)
    names_and_keys <- dm::dm_get_all_pks(dm_obj)
    table_names <- names_and_keys[1]
    primary_keys <- names_and_keys[2]
    num_tables <- dim(table_names)[1]
    duplicate_counter = 0

    # Check each table for duplicate rows, print error for each table containing duplicate row
    for(i in 1:num_tables){
        # Assemble query
        query <- paste0("SELECT ",
                        primary_keys[i,]$pk_col[[1]],
                        ", COUNT(*) occurences FROM ",
                        table_names[i, ]$table[[1]],
                        " GROUP BY ",
                        primary_keys[i,]$pk_col[[1]],
                        " HAVING COUNT(*) > 1;")

        # Test query for database
        result <- DBI::dbGetQuery(con, statement = query)

        # Conditional for when errors are found
        if(!is.na(result[1,1])){
            cat("The table", table_names[i, ]$table[[1]],
                "contains a duplicate row and is not valid\n")
            duplicate_counter = duplicate_counter + 1
        }
    }

    # Conditional for when no errors are found
    if(duplicate_counter == 0){
        print("No duplicates rows found, all tables appear valid")
    }
}
