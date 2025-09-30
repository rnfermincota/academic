#--------------------------------------------------------------------
# Written by Carlos and Nico
# Nov 20th, 2022
#--------------------------------------------------------------------

#------------------------------------------------------------------------------------------------
# Convenience functions for finding coupon dates and number of coupons of a bond.
#------------------------------------------------------------------------------------------------
# Arguments
#------------------------------------------------------------------------------------------------
# - settle: The settlement date for which the bond is traded. Can be a character string or any 
# object that can be converted into date using as.Date.

# - mature: The maturity date of the bond. Can be a character string or any object that can be
# converted into date using as.Date

# - freq: The frequency of coupon payments: 1 for annual, 2 for semi-annual, 12 for monthly.
#------------------------------------------------------------------------------------------------

coupons.dates.f <- function(settle, mature, freq=2){
  # settle="2012-04-15"; mature="2022-01-01"; freq=2
  settle <- as.Date(settle)
  mature <- as.Date(mature)
  m <- -12/freq
  n <- coupons.n.f(settle, mature, freq)
  as.Date(sapply((n:1 - 1), function(i) edate.f(mature, m *i)), origin="1970-01-01")
  # edate.f(mature, m * (n:1 - 1))
}

coupons.n.f <- function(settle, mature, freq=2){
  mature <- as.Date(mature)
  settle <- as.Date(settle)
  n <- as.integer(freq * (mature - settle) / 365.25)
  m <- -12/freq
  while(edate.f(mature, n * m) <= settle) n <- n - 1
  while(edate.f(mature, (n + 1) * m) > settle) n <- n + 1
  n+1
}

coupons.next.f <- function(settle, mature, freq=2){
  settle <- as.Date(settle)
  mature <- as.Date(mature)
  m <- -12/freq
  n <- coupons.n.f(settle, mature, freq)
  edate.f(mature, m * (n-1))
}

coupons.prev.f <- function(settle, mature, freq=2){
  settle <- as.Date(settle)
  mature <- as.Date(mature)
  m <- -12/freq
  n <- coupons.n.f(settle, mature, freq)
  edate.f(mature, m * n)
}

#------------------------------------------------------------------------------------------------
# bond.tcf.f returns a list of three components
# - t: A vector of cash flow dates in number of years
# - cf: A vector of cash flows
# - accrued: The accrued interest
#------------------------------------------------------------------------------------------------
bond.tcf.f <- function(settle, mature, coupon, freq=2, convention=c("30/360", "ACT/ACT", "ACT/360", "30/360E"), redemption_value=100){
  settle <- as.Date(settle)
  mature <- as.Date(mature)
  nextC <- coupons.next.f(settle, mature, freq)
  prevC <- coupons.prev.f(settle, mature, freq)
  accrued <- 100 * coupon * yearFraction.f(prevC, settle, prevC, nextC, freq, convention)
  start <- yearFraction.f(settle, nextC, prevC, nextC, freq, convention)
  t <- seq(from=start, by=1/freq, length.out=coupons.n.f(settle, mature, freq))
  cf <- rep(coupon*100/freq, length(t))
  cf[length(cf)] <- redemption_value + coupon*100/freq
  list(t=t,cf=cf, accrued=accrued)
}

#------------------------------------------------------------------------------------------------
# Maturity time vector in years
bond.tenors.f <- function(
  SETTLEMENT,
  MATURITIES,
  FREQUENCY=2,
  CONVENTION=c("30/360", "ACT/ACT", "ACT/360", "30/360E")
){
  CONVENTION <- match.arg(CONVENTION)
  SETTLEMENT <- as.Date(SETTLEMENT)
  MATURITIES <- as.Date(MATURITIES)
  if (!CONVENTION == "ACT/ACT") {
    TENORS = sapply(X=MATURITIES, function(x) yearFraction.f(d1=SETTLEMENT, d2=x, convention=CONVENTION))
  } else {
    TENORS = sapply(X=MATURITIES, function(x){
      NEXT_COUPON <- coupons.next.f(SETTLEMENT, x, FREQUENCY)
      PREV_COUPON <- coupons.prev.f(SETTLEMENT, x, FREQUENCY)
      yearFraction.f(d1=SETTLEMENT, d2=NEXT_COUPON, r1=PREV_COUPON, r2=NEXT_COUPON, freq=FREQUENCY, convention=CONVENTION)
    })
  }
  return(TENORS)
}

