rm(list = ls(all = TRUE)); gc(); graphics.off()
#----------------------------------------------------------------------------------------------------------
library(dm)
library(DBI)
#----------------------------------------------------------------------------------------------------------
# library(magrittr)
path_root=".."
path_libs=file.path(path_root, "libs")

f1db <- new.env() #rlang::new_environment() # 
list.files(path = path_libs, pattern = "*.R") %>% 
  file.path(path_libs, .) %>% 
  # purrr::map(pryr::partial(source, local=f1db), encoding = 'UTF8') %>% 
  purrr::map(source, local=f1db, encoding = 'UTF8') %>% 
  invisible()

nm <- ls(env=f1db)
# fl <- purrr::map(nm, pryr::partial(get, envir=f1db)) %>% 
fl <- purrr::map(nm, get, envir=f1db) %>% 
  setNames(nm)
fl %>% attach()
#----------------------------------------------------------------------------------------------------------

path_root="." # Download configure and connect Ergast files to SQLite database
con <- f1db_connect() # remeber to call dbDisconnect() when finished working with a connection!
f1db_draw(con) # db schema

#----------------------------------------------------------------------------------------------------------
dbDisconnect(con)
fl %>% detach()
rm(fl, nm)
file.remove(file.path(path_root, "f1_db.SQLite"))
unlink(file.path(path_root, "f1db_csv"), recursive = TRUE) 
rm(con, f1db, dm_obj)
rm(path_libs, path_root)
