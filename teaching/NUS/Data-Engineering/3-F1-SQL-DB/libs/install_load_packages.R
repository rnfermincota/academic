#' @title Installs all dependencies, for internal use only
#' @description Installs all dependencies, for internal use only
#' @param package1 the first package listed that needs to be installed
#' @param ... the remaining n packages
#' @export
#'
# https://stackoverflow.com/questions/15155814/check-if-r-package-is-installed-then-load-library
install_load_packages <- function (package1, ...)  {
    # convert arguments to vector
    packages <- c(package1, ...)

    # start loop to determine if each package is installed
    for(package in packages){
        # if package is installed locally, load
        if(package %in% rownames(installed.packages()))
            do.call('library', list(package))

        # if package is not installed locally, download, then load
        else {
            install.packages(package)
            do.call("library", list(package))
        }
    }
}
