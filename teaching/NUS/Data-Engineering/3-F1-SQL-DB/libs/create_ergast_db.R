#' Download 'Ergast' CSV Database Tables
#' This function downloads the 'Ergast' database as a set of CSV files and unzips them into a local directory
#' @param destfile a character string with the name of the directory in which the files are saved
#' @details The CSV files downloaded by this function have column headers and are UTF-8 encoded. Each file contains a database table.
downloadErgastCSV <- function(destfile = paste0(getwd(), "/f1db_csv")){
  zipdest <- paste0(getwd(), "/f1db_csv.zip")
  download.file("http://ergast.com/downloads/f1db_csv.zip", destfile = zipdest)
  csv_dir <- destfile
  unzip(zipdest, exdir = csv_dir)
  unlink(zipdest)
}


#' Create an F1 database
#' Creates a local 'RSQLite' database using the latest 'Ergast' data and establish a connection.
#' @param csv_dir either NULL or the name of a directory containing csv files from Ergast.
#' If NULL, the files will be downloaded and placed in a directory within the working directory named "/f1db_csv"
#' @param rm_csv logical indicating whether the csv directory should be deleted after initializing the database
#' @param type Indicates the type of database backend used
#' @details \code{createF1db()} creates a local 'RSQLite' database using csv files downloaded from Ergast.
#' The database will be located in a file 'f1_db.RSQLite' within the working directory.
#'
# Databases created with this function can be interacted with using functions from the 'DBI' Packa ge. You can also use the convenience function \code{\link{F1dbConnect}} to reconnect to a database created by \code{createF1db}
#' @return an object of class RSQLite
#' @examples \donttest{
#' library(DBI)
# con <- createF1db()
# dbListTables(con)
# dbDisconnect(con)
#' }
createF1db <- function(csv_dir = NULL, rm_csv = FALSE, type = "sqlite"){
  # devtools::install_github("krlmlr/dm")
  # Download ergast Data
  if(file.exists(paste0(getwd(), "/f1_db.", type))){
    stop("Database file already exists", call. = FALSE)
  }
  if(is.null(csv_dir)){
    downloadErgastCSV()
    csv_dir <- paste0(getwd(), "/f1db_csv")
  }

  # Open sqlite connection
  con <- DBI::dbConnect(RSQLite::SQLite(), "f1_db.SQLite")
  table_names <-  c("constructors","constructor_standings",
                    "constructor_results", "circuits", "drivers",
                    "driver_standings", "lap_times", "pit_stops",
                    "seasons", "status", "races", "results", "qualifying")
  csv_names <- paste0("f1db_csv", "/", table_names, ".csv")
  num_tables <- length(table_names)

  tryCatch(
    {

      # Import from CSV files
      for(i in 1:length(table_names)){
        assign(table_names[i], utils::read.csv(file = csv_names[i], fileEncoding='UTF-8'))
      }

      # Compile list of tables for mapping
      f1_tables <-  list(constructors,constructor_standings,
                         constructor_results, circuits, drivers,
                         driver_standings, lap_times, pit_stops,
                         seasons, status, races, results, qualifying)

      # constructors %>% rename(`\name` = name)
      # Set N/A values to NULL
      # f1_tables_null <- map(f1_tables, null_format)
      #
      # # Unpack updated tables to original table names
      # for(i in 1:num_tables){
      #   assign(table_names[i], f1_tables_null[[i]])
      # }

      constructors <- as.data.frame(lapply(constructors,function(x) if(is.character(x)|is.factor(x)) gsub("\\\\N","NULL",x) else x))
      constructor_standings <- as.data.frame(lapply(constructor_standings,function(x) if(is.character(x)|is.factor(x)) gsub("\\\\N","NULL",x) else x))
      constructor_results <- as.data.frame(lapply(constructor_results,function(x) if(is.character(x)|is.factor(x)) gsub("\\\\N","NULL",x) else x))
      circuits <- as.data.frame(lapply(circuits,function(x) if(is.character(x)|is.factor(x)) gsub("\\\\N","NULL",x) else x))
      drivers <- as.data.frame(lapply(drivers,function(x) if(is.character(x)|is.factor(x)) gsub("\\\\N",NA,x) else x))
      driver_standings <- as.data.frame(lapply(driver_standings,function(x) if(is.character(x)|is.factor(x)) gsub("\\\\N","NULL",x) else x))
      lap_times <- as.data.frame(lapply(lap_times,function(x) if(is.character(x)|is.factor(x)) gsub("\\\\N","NULL",x) else x))
      pit_stops <- as.data.frame(lapply(pit_stops,function(x) if(is.character(x)|is.factor(x)) gsub("\\\\N","NULL",x) else x))
      seasons <- as.data.frame(lapply(seasons,function(x) if(is.character(x)|is.factor(x)) gsub("\\\\N","NULL",x) else x))
      status <- as.data.frame(lapply(status,function(x) if(is.character(x)|is.factor(x)) gsub("\\\\N","NULL",x) else x))
      races <- as.data.frame(lapply(races,function(x) if(is.character(x)|is.factor(x)) gsub("\\\\N",NA,x) else x))
      results <- as.data.frame(lapply(results,function(x) if(is.character(x)|is.factor(x)) gsub("\\\\N",NA,x) else x))
      qualifying <- as.data.frame(lapply(qualifying,function(x) if(is.character(x)|is.factor(x)) gsub("\\\\N","NULL",x) else x))

      # Initial DM
      dm <- dm::dm(constructors, constructor_standings,
               constructor_results, circuits, drivers,
               driver_standings, lap_times, pit_stops,
               seasons, status, races, results, qualifying)


      # Primary constraints (table, primary key)
      dm_primary_keys <-
        dm %>%
        dm::dm_add_pk(circuits, circuitId) %>%
        dm::dm_add_pk(constructors, constructorId) %>%
        dm::dm_add_pk(drivers, driverId) %>%
        dm::dm_add_pk(results, resultId) %>%
        dm::dm_add_pk(races, raceId) %>%
        dm::dm_add_pk(constructor_standings, constructorStandingsId) %>%
        dm::dm_add_pk(constructor_results, constructorResultsId) %>%
        dm::dm_add_pk(qualifying, qualifyId) %>%
        dm::dm_add_pk(seasons, year) %>%
        dm::dm_add_pk(status, statusId) %>%
        dm::dm_add_pk(driver_standings, driverStandingsId)

      # Confirm constraints are valid
      pk_check <- dm_primary_keys %>% dm::dm_examine_constraints()

      # Add foreign key constraints (table, columns, ref_table, ref_columns)
      if(all(pk_check$is_key)){
        dm_foreign_keys <-
          dm_primary_keys %>%
          dm::dm_add_fk(pit_stops, raceId, races, raceId) %>%
          dm::dm_add_fk(pit_stops, driverId, drivers, driverId) %>%
          dm::dm_add_fk(lap_times, raceId, races, raceId) %>%
          dm::dm_add_fk(lap_times, driverId, drivers, driverId)
      }

      # Confirm constraints are valid
      fk_check <- dm_foreign_keys %>% dm::dm_examine_constraints()

      if(all(fk_check$is_key)){
        # Additional constraints
        dm_all_keys <-
          dm_foreign_keys %>%
          dm::dm_add_fk(constructor_standings, raceId, races, raceId) %>%
          dm::dm_add_fk(constructor_standings, constructorId, constructors) %>%
          dm::dm_add_fk(results, constructorId, constructors, constructorId) %>%
          dm::dm_add_fk(results, statusId, status, statusId) %>%
          dm::dm_add_fk(results, driverId, drivers, driverId) %>%
          dm::dm_add_fk(results, raceId, races, raceId) %>%
          dm::dm_add_fk(races, year, seasons, year) %>%
          dm::dm_add_fk(qualifying, raceId, races, raceId) %>%
          dm::dm_add_fk(qualifying, constructorId, constructors, constructorId) %>%
          dm::dm_add_fk(qualifying, driverId, drivers, driverId)  %>%
          dm::dm_add_fk(constructor_results, constructorId, constructors, constructorId) %>%
          dm::dm_add_fk(constructor_results, raceId, races, raceId)  %>%
          dm::dm_add_fk(driver_standings, raceId, races, raceId) %>%
          dm::dm_add_fk(driver_standings, driverId, drivers, driverId)
      }

      # Confirm constraints are valid
      all_check <- dm_all_keys %>% dm::dm_examine_constraints()

      # Copy dm to sqlite with all constraints
      if(all(all_check$is_key)){
        db_dm <- dm::copy_dm_to(con, dm_all_keys,
                            temporary = FALSE,
                            set_key_constraints = TRUE)
      }

    },

    error = function(e){
      DBI::dbDisconnect(con)
      unlink("f1_db.sqlite")
      stop(e)
    }

  )


  if(rm_csv){
    unlink(csv_dir)
  }
  print(db_dm)
  return(list(con, db_dm))


}

# Connect to an existing F1 database
# Establishes a connection to a 'DuckDB' database previously created by \code{\link{createF1db}}.
# @param file path to the database file
# @return an object of class \code{\link[duckdb:duckdb_connection-class]{duckdb_connection}}
# @examples \donttest{
# # a file "f1_db.duckdb" already exists in the working directory
# con <- F1dbConnect()
# dbListFields(con)
# circuits <- dbReadTable(con, "circuits")
# }

# null_format <- function(table){
#   df <- as.data.frame(lapply(table,function(x) if(is.character(x)|is.factor(x)) gsub("\\\\N","NULL",x) else x))
#   return(df)
#
# }
# F1dbConnect <- function(file = "f1_db.duckdb"){
#   if(!file.exists(file)){
#     stop(paste0("Database file ", file, " not found"), call. = FALSE)
#   }
#   if(grepl(".duckdb", file, ignore.case = TRUE)){
#     DBI::dbConnect(duckdb::duckdb(), file)
#   } else {
#
#     stop("file does not have extension '.duckdb'", call. = FALSE)
#   }
#
# }

