#--------------------------------------------------------------------
# Written by Carlos and Nico
# Nov 20th, 2022
#--------------------------------------------------------------------

# The JPM Cheap Rich Model refrains from modelling the unique interest rates at each
# point in time because these are, in colloquial terms, “not very well behaved”. Experience
# shows that interest term structures come in all kinds of shapes. There are not just 
# upward or downward sloping curves but they often have “humps” and “kinks” that make 
# the formulation of an suitable model very demanding. On the other hand, modelling the 
# discount factors is an easier approach. This function has a clear boundary at time 
# zero and is monotonously declining over time. In the JPM model it is modelled as a 
# polynomial with coefficients determined such that the sum of least square errors of 
# market price minus model price is minimized.
#--------------------------------------------------------------------

jpm.cheap.rich.model.f <- function(
  VALOR_VECTOR,PRICE_VECTOR, MATURITY_VECTOR, COUPON_VECTOR, SHORT_RATE, 
  COEFFICIENT_VECTOR=NULL, TENORS_VECTOR=NULL, 
  SETTLEMENT=Sys.Date(), NDEG=3, FREQUENCY=2, PAR_VALUE=1, CONVENTION="30/360"
){
  # This function implements the slightly more complex JP Morgan and McCulloch
  # discount function routines, an approach of term-structure of interest modeling.
  # The JP Morgan model was inspired by McCulloch's (1971,75) spline models that are
  # now widely used in industry. Its limitations come to light when bond maturities in
  # the basket are not evenly distributed on the time scale which means the fit is biased
  # towards the concentration of bonds.
  INTEREST_VECTOR=COUPON_VECTOR * vapply(MATURITY_VECTOR, function(m){
    t=as.Date(m, origin="1970-01-01")
    nextC=coupons.next.f(SETTLEMENT, t, FREQUENCY)
    prevC=coupons.prev.f(SETTLEMENT, t, FREQUENCY)
    yearFraction.f(prevC, SETTLEMENT, prevC, nextC, FREQUENCY, CONVENTION)
  }, numeric(1))
  # Cash Price Vector / Dirty Price vector
  CASH_VECTOR=PRICE_VECTOR / PAR_VALUE + INTEREST_VECTOR  
  
  NCOLS=1:(NDEG+1)
  NROWS=length(MATURITY_VECTOR)
  
  TEMP_MATRIX=lapply( # foreach(i=1:NROWS, .combine=rbind) %dopar%
    1:NROWS,
    function(i){
      vapply(NCOLS, function(j){
        jpm.polynomial.sum.f(
          j - 1, 
          SETTLEMENT, 
          MATURITY_VECTOR[i], 
          COUPON_VECTOR[i], 
          FREQUENCY, 
          CONVENTION)
      }, numeric(1))
    }
  )
  TEMP_MATRIX=do.call(rbind, TEMP_MATRIX)
  colnames(TEMP_MATRIX)=NCOLS; rownames(TEMP_MATRIX)=NULL
  
  if(is.null(COEFFICIENT_VECTOR)){
    if (SHORT_RATE == 0){ # 0 - no restriction for d(0)
      y=CASH_VECTOR; x=TEMP_MATRIX
      COEFFICIENT_VECTOR=as.vector((solve(t(x) %*% x)) %*% (t(x) %*% y))  # as.vector(lm(y ~ 0 + x)$coef)
    } else if  (SHORT_RATE == 1){ # 1 - d(0)=1: restriction: Alpha = 1 and Beta 1 = -ln(1+short_term rate) 
      y=CASH_VECTOR - TEMP_MATRIX[,1]; x=TEMP_MATRIX[,-1]
      COEFFICIENT_VECTOR=c(1, as.vector((solve(t(x) %*% x)) %*% (t(x) %*% y))) # as.vector(lm(y ~ 0 + x)$coef)
    } else { # any other value:  short-term rate at time t=0. e.g. 0.05 means short-term rate = 5%
      t_val=-log(1 + SHORT_RATE)
      y=CASH_VECTOR - TEMP_MATRIX[,1] - (t_val * TEMP_MATRIX[,2]); x=TEMP_MATRIX[,-(1:2)]
      COEFFICIENT_VECTOR=c(1, t_val, as.vector((solve(t(x) %*% x)) %*% (t(x) %*% y))) # as.vector(lm(y ~ 0 + x)$coef)
    }
    # (solve(t(x) %*% x)) %*% (t(x) %*% y)
    # (chol2inv(chol(t(x) %*% x))) %*% (t(x) %*% y)
    # (MASS::ginv(t(x) %*% x)) %*% (t(x) %*% y)
    # xt = t(x); xtx = xt %*% x; xtxi=solve(xtx); xty = xt %*% Y; xtxi %*% xty
    # https://github.com/rnfermincota/academic/blob/main/teaching/Ivey/8.%20Nico-Add-Ins/backup/STAT_REGRESSION_COEF_LIBR.bas
  }
  
  FAIR_PRICE_VECTOR=as.vector(((TEMP_MATRIX %*% COEFFICIENT_VECTOR) - INTEREST_VECTOR * PAR_VALUE))
  CHEAP_RICH_VECTOR=PRICE_VECTOR - FAIR_PRICE_VECTOR
  CHEAP_RICH_TBL=data.frame(
    valor=VALOR_VECTOR,
    clean_price=PRICE_VECTOR,
    fair_price=FAIR_PRICE_VECTOR,
    # If a bond is trading above the present value determined by discounting
    # coupon and principal at the zero coupon rates, we say the bond is "rich",
    # otherwise it is "cheap".
    cheap_rich=CHEAP_RICH_VECTOR
  )
  
  if(!is.null(TENORS_VECTOR)){
    # TENORS_VECTOR=seq(MIN_TENOR, MAX_TENOR, WIDTH_TENOR)
    # TENORS_VECTOR=c(0.000001,0.5,1,1.5,2,2.5,3,3.5,4,4.5,5,6,7,8,9,10)
    DISCOUNTS_VECTOR=as.vector(t(polynomial.interpolation.f(COEFFICIENT_VECTOR, TENORS_VECTOR)))
    # DISCOUNTS_VECTOR=as.vector(polynomial.interpolation.f(COEFFICIENT_VECTOR, TENORS_VECTOR))
    RATES_VECTOR=DISCOUNTS_VECTOR^(-1/TENORS_VECTOR)-1
    # DISCOUNTS_VECTOR=1/(1+RATES_VECTOR)^TENORS_VECTOR
    if (SHORT_RATE !=0 & SHORT_RATE!=1){RATES_VECTOR=c(SHORT_RATE, RATES_VECTOR[-1])}
    RATES_INTERPOLATION_TBL=data.frame(
      tenor=TENORS_VECTOR, 
      discount_factor=DISCOUNTS_VECTOR,
      rates=RATES_VECTOR
    )
  } else {
    RATES_INTERPOLATION_TBL=NA
  }
  
  return(list(
    coefficients=COEFFICIENT_VECTOR,
    cheap_rich=CHEAP_RICH_TBL,
    interpolation=RATES_INTERPOLATION_TBL
  ))
}
#--------------------------------------------------------------------

# Sum coefficient matrix with (C*t^NDEG) expressions
jpm.polynomial.sum.f <- function(
  NDEG, SETTLEMENT, MATURITY,
  COUPON, FREQUENCY=2, 
  CONVENTION=c("30/360", "ACT/ACT", "ACT/360", "30/360E")
){
  TENOR_VAL=bond.tenors.f(SETTLEMENT, MATURITY, FREQUENCY, CONVENTION)
  D_VAL=1 / FREQUENCY
  TEMP_VAL=(TENOR_VAL * FREQUENCY - as.integer(TENOR_VAL * FREQUENCY)) * D_VAL # / FREQUENCY
  TENOR_VAL=TENOR_VAL - TEMP_VAL
  if (TEMP_VAL == 0){
    TEMP_VAL=D_VAL
    TENOR_VAL=TENOR_VAL - D_VAL
  }
  TEMP_SUM=sum((COUPON * D_VAL) * (seq(0, TENOR_VAL, D_VAL) + TEMP_VAL) ^ NDEG)
  # TEMP_SUM=0; for (i in seq(0, TENOR_VAL, D_VAL)){TEMP_SUM=TEMP_SUM + (COUPON * D_VAL) * (i + TEMP_VAL) ^ NDEG}
  TEMP_SUM=TEMP_SUM + (TENOR_VAL + TEMP_VAL) ^ NDEG  # Principal at MATURITY
  return(TEMP_SUM)
}
