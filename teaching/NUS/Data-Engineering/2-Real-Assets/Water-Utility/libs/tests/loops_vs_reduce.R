library(purrr)
# http://adv-r.had.co.nz/Functionals.html
#-----------------------------------------------

RND_VEC=c(1.75556803913375, -0.0799365289157126, 1.63025368036105, 0.268892180257083, 
          -0.0817688484778561, 0.700559159611574, -0.668244340991706, 0.216952495929201, 
          -0.962442548216423, -0.717517445102393, 0.0986384764455445, -0.667207935530494, 
          -0.619421357051072, -0.991318012291069, 0.634356804887895, 1.06147409139581, 
          0.495289926518602, -0.76369614118954, 0.504454895729237, 0.120628984655496, 
          -1.13045729301784)

EXCHANGE_RATE_CURRENT_VAL=100
EXCHANGE_RATE_DEPRECIATION_VAL=0.05
EXCHANGE_RATE_VOLATILITY_VAL=0.05

#-----------------------------------------------
# This function can estimate the value varies with fluctuations in the exchange rate.
exchange.rate.f <- function(prv, nxt){
  prv * exp(EXCHANGE_RATE_DEPRECIATION_VAL - EXCHANGE_RATE_VOLATILITY_VAL * 
              EXCHANGE_RATE_VOLATILITY_VAL / 2 + EXCHANGE_RATE_VOLATILITY_VAL * nxt)
}


#-----------------------------------------------
n=length(RND_VEC)
out0=vector(mode = "numeric", length = n)
out0[1]=EXCHANGE_RATE_CURRENT_VAL
for (i in 2:n){ # Loop
  prv=out0[i-1]
  nxt=RND_VEC[i]
  out0[i]= exchange.rate.f(prv, nxt)
}
out0


#-----------------------------------------------
# It's hard to convert a for loop into a functional when the relationship between
# elements is not independent, or is defined recursively. For example, exponential 
# smoothing works by taking a weighted average of the current and previous data
# points. We can't eliminate the for loop because none of the functionals we've 
# seen allow the output at position `i` to depend on both the input and output at 
# position `i - 1`. One way to eliminate the for loop in this case is to [solve 
# the recurrence relation](http://en.wikipedia.org/wiki/Recurrence_relation#Solving) 
# by removing the recursion and replacing it with explicit references. This requires 
# a new set of mathematical tools, and is challenging, but it can pay off by 
# producing a simpler function.
#-----------------------------------------------

out1=Reduce( # Reduce Base R
  # Expected annual change in the exchange rate. A positive change is a depreciation of
  # the currency (more pesos per dollar), and a negative change is an appreciation of
  # the peso.
  
  # Create a function 
  function(prv,nxt) {
    # Do the following calculations with the values from the original input and the function created above
    exchange.rate.f(prv, nxt)
  }, 
  # Apply the function created above to the numbers from the vector rnd2 and without the first index
  RND_VEC[-1], # next
  # Initial value
  EXCHANGE_RATE_CURRENT_VAL, 
  # Accumulate the results
  accumulate = TRUE
  # Skip the first index
)
all.equal(out0, out1)

#-----------------------------------------------

out2=purrr::accumulate( # accumulate Purrr package
  # Apply the function created above to the numbers from the vector rnd2 and without the first index
  .x=RND_VEC[-1], # next
  
  # Expected annual change in the exchange rate. A positive change is a depreciation of
  # the currency (more pesos per dollar), and a negative change is an appreciation of
  # the peso.

  # Initial value
  .init=EXCHANGE_RATE_CURRENT_VAL, 
  
  # Create a function 
  .f=function(prv,nxt) {
    # Do the following calculations with the values from the original input and the function created above
    exchange.rate.f(prv, nxt)
  }
)
all.equal(out0, out2)

#-----------------------------------------------

out3=purrr::reduce( # reduce Purrr package
  # Apply the function created above to the numbers from the vector rnd2 and without the first index
  .x=RND_VEC[-1], # next
  
  # Expected annual change in the exchange rate. A positive change is a depreciation of
  # the currency (more pesos per dollar), and a negative change is an appreciation of
  # the peso.
  
  # Initial value
  .init=EXCHANGE_RATE_CURRENT_VAL, 
  
  # Create a function 
  .f=function(prv,nxt) {
    # Do the following calculations with the values from the original input and the function created above
    exchange.rate.f(prv, nxt)
  }
)
out3
all.equal(out0[length(out0)], out3)

