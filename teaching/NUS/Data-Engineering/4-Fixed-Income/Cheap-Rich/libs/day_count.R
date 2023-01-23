#--------------------------------------------------------------------
# Written by Carlos and Nico
# Nov 20th, 2022
#--------------------------------------------------------------------

#------------------------------------------------------------------------------------------------
# Day count and year fraction for bond pricing
#------------------------------------------------------------------------------------------------
# Implements 30/360, ACT/360, ACT/360 and 30/360E day count conventions.
#------------------------------------------------------------------------------------------------
# Arguments
#------------------------------------------------------------------------------------------------
# - d1: The starting date of period for day counts

# - d2: The ending date of period for day counts

# - r1: The starting date of reference period for ACT/ACT day counts

# - r2: The ending date of reference period for ACT/ACT day counts

# - freq: The frequency of coupon payments: 1 for annual, 2 for semi-annual, 12 for monthly.

# - convention: The daycount convention

# - variant: Three variants of the 30/360 convention are implemented, but only one variant of 
# ACT/ACT is currently implemented

#------------------------------------------------------------------------------------------------
# References
#------------------------------------------------------------------------------------------------
# The 30/360 day count was converted from C++ code in the QuantLib library
#------------------------------------------------------------------------------------------------

yearFraction.f <- function(d1, d2, r1, r2, freq=2, convention=c("30/360", "ACT/ACT", "ACT/360", "30/360E")){
  convention <- match.arg(convention)
  if (convention == "ACT/ACT") (1/freq) * daycount.actual.f(d1,d2) / daycount.actual.f(r1,r2)
  else if (convention == "ACT/360") daycount.actual.f(d1,d2) / 360
  else if (convention == "30/360") daycount.30.360.f(d1,d2) / 360
  else if (convention == "30/360E") daycount.30.360.f(d1,d2, "E") / 360
}

daycount.actual.f <- function(d1, d2, variant = c("bond")){
  as.integer(as.Date(d2) - as.Date(d1))
}

daycount.30.360.f <- function(d1, d2, variant = c("US", "EU", "IT")){
  ## The algorithm is taken from the QuantLib source code
  D1 <- as.POSIXlt(d1)
  D2 <- as.POSIXlt(d2)
  dd1 <- D1$mday
  dd2 <- D2$mday
  mm1 <- D1$mon
  mm2 <- D2$mon
  yy1 <- D1$year
  yy2 <- D2$year
  variant <- match.arg(variant)
  if (variant == "US" && dd2 == 31 && dd1 < 30) {
    dd2 <- 1
    mm2 <- mm2 + 1
  }
  if (variant == "IT" && mm1 == 2 && dd1 > 27) dd1 = 30
  if (variant == "IT" && mm2 == 2 && dd2 > 27) dd2 = 30
  360*(yy2-yy1) + 30*(mm2-mm1-1) + max(0,30-dd1) + min(30,dd2)
}

# Shift date by a number of months
# Convenience function for finding the same date in different months. Used for example
# to find coupon dates of bonds given the maturity date.
edate.f <- function(date, n = 1){
  # This function doesn't require any packages to be installed. You give it a Date object 
  # (or a character that it can convert into a Date), and it adds n months to that date 
  # without changing the day of the month (unless the month you land on doesn't have enough
  # days in it, in which case it defaults to the last day of the returned month). Just in 
  # case it doesn't make sense reading it, there are some examples below.
  # source: https://stackoverflow.com/questions/14169620/add-a-month-to-a-date
  if (n == 0){return(date)}
  if (n %% 1 != 0){stop("Input Error: argument 'n' must be an integer.")}
  
  # Check to make sure we have a standard Date format
  if (class(date) == "character"){date = as.Date(date)}
  
  # Turn the year, month, and day into numbers so we can play with them
  y = as.numeric(substr(as.character(date),1,4))
  m = as.numeric(substr(as.character(date),6,7))
  d = as.numeric(substr(as.character(date),9,10))
  
  # Run through the computation
  i = 0
  # Adding months
  if (n > 0){
    while (i < n){
      m = m + 1
      if (m == 13){
        m = 1
        y = y + 1
      }
      i = i + 1
    }
  }
  # Subtracting months
  else if (n < 0){
    while (i > n){
      m = m - 1
      if (m == 0){
        m = 12
        y = y - 1
      }
      i = i - 1
    }
  }
  
  # If past 28th day in base month, make adjustments for February
  if (d > 28 & m == 2){
    # If it's a leap year, return the 29th day
    if ((y %% 4 == 0 & y %% 100 != 0) | y %% 400 == 0){d = 29}
    # Otherwise, return the 28th day
    else{d = 28}
  }
  # If 31st day in base month but only 30 days in end month, return 30th day
  else if (d == 31){if (m %in% c(1, 3, 5, 7, 8, 10, 12) == FALSE){d = 30}}
  
  # Turn year, month, and day into strings and put them together to make a Date
  y = as.character(y)
  
  # If month is single digit, add a leading 0, otherwise leave it alone
  if (m < 10){m = paste('0', as.character(m), sep = '')}
  else{m = as.character(m)}
  
  # If day is single digit, add a leading 0, otherwise leave it alone
  if (d < 10){d = paste('0', as.character(d), sep = '')}
  else{d = as.character(d)}
  
  # Put them together and convert return the result as a Date
  return(as.Date(paste(y,'-',m,'-',d, sep = '')))
}
