# Written by Carlos Arias, Joaquin Calderon, and Rafael Nicolas Fermin Cota
#-----------------------------------------------------------------------------------------
rm(list=ls())
graphics.off()
#-----------------------------------------------------------------------------------------
library(dplyr)
library(purrr)
#-----------------------------------------------------------------------------------------
options(scipen=999)
options(digits=6)
#-----------------------------------------------------------------------------------------
path_roots=".."
#-----------------------------------------------------------------------------------------
path_libs=file.path(path_roots, "libs")
source(file.path(path_libs, "code.R"))
#-----------------------------------------------------------------------------------------
path_inputs=file.path(path_roots, "inputs")
path_output=file.path(path_roots, "output")
#-----------------------------------------------------------------------------------------
excel_risk=readr::read_csv(file=file.path(path_output, "risk.csv"))
excel_risk[,52]=NA_real_; excel_risk[,53]=NA_real_
excel_risk[which(excel_risk[,90]=="#DIV/0!"), 90]=NA_character_
excel_risk[,90]=as.numeric(excel_risk[[90]])
excel_risk[which(excel_risk[,91]=="#DIV/0!"), 91]=NA_character_
excel_risk[,91]=as.numeric(excel_risk[[91]])

excel_no_risk=readr::read_csv(file=file.path(path_output, "no_risk.csv"))
excel_no_risk[,32]=NA_real_; excel_no_risk[,33]=NA_real_
excel_no_risk[,42]=NA_real_; excel_no_risk[,43]=NA_real_
excel_no_risk[,54]=NA_real_; excel_no_risk[,55]=NA_real_
excel_no_risk[which(excel_no_risk[,90]=="#DIV/0!"), 90]=NA_character_
excel_no_risk[,90]=as.numeric(excel_no_risk[[90]])
excel_no_risk[which(excel_no_risk[,91]=="#DIV/0!"), 91]=NA_character_
excel_no_risk[,91]=as.numeric(excel_no_risk[[91]])

#-----------------------------------------------------------------------------------------

inputs=readr::read_csv(file.path(path_inputs, "inputs.csv"), col_names=F) %>%
  select(var=X1, val=X2) %>%
  mutate(
    val=ifelse(val == TRUE, 1, ifelse(val == FALSE, 0, val)),
    val=as.numeric(val)
  )

nm=inputs$var
inputs=inputs$val
names(inputs)=nm

#-----------------------------------------------------------------------------------------
cc=1:102
# which(names(excel_risk)!=names(mat))
# names(mat)[which(names(excel_risk)!=names(mat))]
# names(excel_risk)[which(names(excel_risk)!=names(mat))]

mat_no_risk=WATER_UTILITY_SAMPLING_FUNC(inputs, F)
mat_no_risk %>% map(~.x)

mat_risk=WATER_UTILITY_SAMPLING_FUNC(inputs, T)
mat_risk %>% map(~.x)

tol=10^-4

valid_no_risk=abs(mat_no_risk[,cc]-excel_no_risk[,cc])
any(as.vector(colSums(valid_no_risk, na.rm = T)) > tol)
sum(as.vector(colSums(valid_no_risk, na.rm = T)))

valid_risk=abs(mat_risk[,cc]-excel_risk[,cc])
any(as.vector(colSums(valid_risk, na.rm = T)) >  tol)
sum(as.vector(colSums(valid_risk, na.rm = T)))

names(mat_risk)

#-----------------------------------------------------------------------------------------
rm(valid_no_risk, valid_risk)
rm(cc, nm, tol)
rm(excel_no_risk, excel_risk)
rm(mat_no_risk, mat_risk)
rm(inputs)
rm(path_roots, path_libs, path_inputs, path_output)
rm(WATER_UTILITY_SAMPLING_FUNC)
