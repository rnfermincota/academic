#--------------------------------------------------------------------
# Written by Carlos and Nico
# Nov 20th, 2022
#--------------------------------------------------------------------

# Least-squares regression with polynomials
# https://github.com/rnfermincota/academic/blob/main/teaching/Ivey/8.%20Nico-Add-Ins/backup/POLYNOMIAL_lu_LIBR.bas
polynomial.regression.f <- function(x, y, NDEG=3){
  NROWS=length(y)
  X=cbind(rep(1, NROWS), vapply(X=1:NDEG, function(d) x ^ d, numeric(NROWS)))
  XT=t(X); # XTX=XT %*% X; XTXI=solve(XTX); XTXIXT = XTXI %*% XT; XTXIXT %*% y
  # chol2inv(chol(XTX)) or MASS::ginv(XTX) # X'X -1
  as.vector((solve(XT %*% X) %*% XT) %*% y)
}

# interpolate discount factors from a polynomial coefficient vector
polynomial.interpolation.f <- function(COEF_VECTOR, XDATA_VECTOR){
  NDEG=length(COEF_VECTOR)   # No. of Degrees in the Polynomial 
  NROWS=length(XDATA_VECTOR)
  COEF_VECTOR=matrix(data=COEF_VECTOR, nrow=NDEG, ncol=1)
  XDATA_VECTOR=matrix(data=XDATA_VECTOR, nrow=NROWS, ncol=1)
  # sapply(X=(1:NDEG) - 1, function(d) XDATA_VECTOR ^ d) %*% COEF_VECTOR
  vapply(X=(1:NDEG) - 1, function(d) XDATA_VECTOR ^ d, numeric(NROWS)) %*% COEF_VECTOR
}
