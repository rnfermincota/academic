rm(list=ls())
#--------------------------------------------------------------------
# Written by Carlos and Nico
# Nov 20th, 2022
#--------------------------------------------------------------------

library(dplyr)
library(purrr)
library(readxl)
library(lubridate)


# library(lubridate)
#--------------------------------------------------------------------

path_root="."
path_libs=file.path(path_root, "libs")

source(file.path(path_libs, "day_count.R"))
source(file.path(path_libs, "coupons.R"))
source(file.path(path_libs, "interpolation.R"))
source(file.path(path_libs, "cheap_rich.R"))

#---------------------------------------------
#---------------------------------------------

INPUTS_LST=list(
  VALOR_VECTOR=1:8,
  PRICE_VECTOR=c(1.005727,1.0281965,1.084659,0.9674985,1.0513405,1.0666335,1.0072635,0.9185795),
  MATURITY_VECTOR=as.Date(c("2000-02-15","2001-02-15","2002-03-15","2003-04-15","2004-04-15","2006-11-15","2009-07-15","2011-11-15"), origin="1970-01-01"),
  COUPON_VECTOR=c(0.065,0.08,0.1,0.055,0.08,0.08,0.07,0.060),
  SHORT_RATE=0,
  COEFFICIENT_VECTOR=NULL,
  # TENORS_VECTOR=NULL,
  TENORS_VECTOR=c(0.000001,0.5,1,1.5,2,2.5,3,3.5,4,4.5,5,6,7,8,9,10),
  SETTLEMENT=as.Date("1999-02-14", origin="1970-01-01"),
  NDEG=3,
  FREQUENCY=2,
  PAR_VALUE=1,
  CONVENTION="30/360"
)
# RST_LST=lift_dl(jpm.cheap.rich.model.f)(INPUTS_LST)
# for (j in 1:length(INPUTS_LST)) assign(names(INPUTS_LST)[j], INPUTS_LST[[j]])

#---------------------------------------------
# https://rpubs.com/rafael_nicolas/fixed_income_relative_value
RST_LST=lift_dl(jpm.cheap.rich.model.f)(update_list(INPUTS_LST, SHORT_RATE=0))
RST_LST %>%
  pluck("cheap_rich") %>%
  pull("fair_price") %>%
  all.equal(c(1.00342136529557,1.02808433059722,1.09297442809564,0.956517606453,1.05357743634413,1.06863895063501,1.00627438070289,0.918772184688613), tolerance=0.0000001)

RST_LST %>%
  pluck("interpolation") %>% pull("discount_factor") %>%
  all.equal(c(1.00614432170623,0.973104189882685,0.941025121136586,0.909886936128736,0.879669388464918,0.850352231750915,0.82191521959251,0.794338105595487,0.767600643365627,0.741682586508715,0.716563688630532,0.668642384233488,0.62367475902076,0.58149884183861,0.541952661533302,0.504874246951102), tolerance=0.0000001)

RST_LST %>%
  pluck("coefficients") %>% 
  all.equal(c(1.00614438876125,-0.0670550254857419,0.00196275305336657,-2.69951922893869E-05), tolerance=0.000001)

#---------------------------------------------
RST_LST=lift_dl(jpm.cheap.rich.model.f)(update_list(INPUTS_LST, SHORT_RATE=1))
RST_LST %>%
  pluck("coefficients") %>% 
  all.equal(c(1,-0.0631554759835518,0.00134995935411772,4.03586344652727E-07), tolerance=0.000001) # SHORT_RATE=1

RST_LST %>%
  pluck("interpolation") %>% pull("discount_factor") %>%
  all.equal(c(0.999999936844525,0.968759802295047,0.938194886956911,0.90830555667535,0.879092114140124,0.850554862040992,0.82269410306771,0.795510139910038,0.769003275257734,0.743173811800557,0.718022052228266,0.669752855497372,0.624198106583122,0.581360227003582,0.541241638276821,0.503844761920907), tolerance=0.0000001)

RST_LST %>%
  pluck("coefficients") %>% 
  all.equal(c(1,-0.0631554759835518,0.00134995935411772,4.03586344652727E-07), tolerance=0.000001)

#---------------------------------------------
RST_LST=lift_dl(jpm.cheap.rich.model.f)(update_list(INPUTS_LST, SHORT_RATE=0.05))
RST_LST %>%
  pluck("coefficients") %>% 
  all.equal(c(1,-0.048790164169432,-0.00222865577372254,0.000197075914729369), tolerance=0.000001) # SHORT_RATE=0.05

RST_LST %>%
  pluck("interpolation") %>% pull("discount_factor") %>%
  all.equal(c(0.999999951209834,0.975072388461194,0.949178255971575,0.922465409467188,0.895081655884081,0.8671748021583,0.838892655225894,0.810383022022908,0.781793709485391,0.753272524549388,0.724967274150947,0.66959580471094,0.616861756653745,0.567947585467738,0.524035746641296,0.486308695662794), tolerance=0.0000001)

RST_LST %>%
  pluck("coefficients") %>% 
  all.equal(c(1,-0.048790164169432,-0.00222865577372254,0.000197075914729369), tolerance=0.000001)

#---------------------------------------------
#---------------------------------------------
rm(path_root, path_libs)
rm(INPUTS_LST, RST_LST)
