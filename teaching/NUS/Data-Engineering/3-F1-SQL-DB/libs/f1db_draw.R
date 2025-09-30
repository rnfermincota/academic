build_schema_diagram <- function(constructors,constructor_standings,
                                 constructor_results, circuits, drivers,
                                 driver_standings, lap_times, pit_stops,
                                 seasons, status, races, results, qualifying){
    # Set colour groups
    display <- list(
        accent1 = c("constructor_results", "constructors", "constructor_standings"),
        accent2 = c("drivers", "driver_standings"),
        accent3 = c("lap_times", "pit_stops", "qualifying"),
        accent4 = c("results", "circuits","races", "seasons"),
        accent5 = c("status")
        )
    datamodelr::dm_from_data_frames(constructors,constructor_standings,
                        constructor_results, circuits, drivers,
                        driver_standings, lap_times, pit_stops,
                        seasons, status, races, results, qualifying) %>%
        datamodelr::dm_set_key("circuits", "circuitId") %>%
        datamodelr::dm_set_key("constructors", "constructorId") %>%
        datamodelr::dm_set_key("drivers", "driverId") %>%
        datamodelr::dm_set_key("results", "resultId") %>%
        datamodelr::dm_set_key("races", "raceId") %>%
        datamodelr::dm_set_key("constructor_standings", "constructorStandingsId") %>%
        datamodelr::dm_set_key("constructor_results", "constructorResultsId") %>%
        datamodelr::dm_set_key("qualifying", "qualifyId") %>%
        datamodelr::dm_set_key("seasons", "year") %>%
        datamodelr::dm_set_key("status", "statusId") %>%
        datamodelr::dm_set_key("driver_standings", "driverStandingsId") %>%
        datamodelr::dm_set_key("pit_stops", c("raceId", "driverId")) %>%
        datamodelr::dm_set_key("lap_times", c("raceId", "driverId")) %>%
        datamodelr::dm_add_references(
            constructor_standings$raceId == races$raceId,
            constructor_standings$constructorId == constructors$constructorId,
            results$constructorId == constructors$constructorId,
            results$statusId == status$statusId,
            results$driverId == drivers$driverId,
            results$raceId == races$raceId,
            races$year == seasons$year,
            qualifying$raceId == races$raceId,
            qualifying$constructorId == constructors$constructorId,
            qualifying$driverId == drivers$driverId,
            constructor_results$constructorId == constructors$constructorId,
            constructor_results$raceId == races$raceId,
            driver_standings$raceId == races$raceId,
            driver_standings$driverId == drivers$driverId,
            pit_stops$raceId == lap_times$raceId,
            pit_stops$raceId == races$raceId,
            pit_stops$driverId == lap_times$driverId,
            pit_stops$driverId == drivers$driverId,
            lap_times$raceId == races$raceId,
            lap_times$driverId == drivers$driverId

        ) %>%
        datamodelr::dm_set_display(display = display) %>%
        datamodelr::dm_create_graph(rankdir = "BT",
                        graph_name = "F1 Ergast Schema",
        ) %>%
        datamodelr::dm_render_graph()
}

#' @title Constructs database schema diagram.
#' @description Constructs database schema diagram.
#' @param con
#' A SQLite database connection object
#' @return A sqlite database connection object.
#' @importFrom magrittr %>%
#' @export

f1db_draw <- function(con){
    table_names <-  c("constructors","constructor_standings",
                      "constructor_results", "circuits", "drivers",
                      "driver_standings", "lap_times", "pit_stops",
                      "seasons", "status", "races", "results", "qualifying")

    # Database tables
    for(i in 1:length(table_names)){
        assign(table_names[i], tbl(con, table_names[i]) %>% collect())
    }

    # Visualizer
    build_schema_diagram(constructors,constructor_standings,
                 constructor_results, circuits, drivers,
                 driver_standings, lap_times, pit_stops,
                 seasons, status, races, results, qualifying)

}
