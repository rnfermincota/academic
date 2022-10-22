################################################################################
#Functions to process data.

################################################################################
##Libs
library(xts)
library(lubridate)

################################################################################
##Required Funcitons
################################################################################

### costs - Add point distribution, accumulate with interest
add.item.fixed <- function(df, 
                           item,
                           name         =NULL, 
                           group        =NULL, 
                           from.1       =start.period, #test if this goes to the period end i.e. t+1 = t+ 1 day
                           to.1         =est.length,
                           from.2       =NULL,
                           to.2         =NULL,
                           distribution = c("uniform", "tiered", "Scurve", "point"),
                           flag         = c("cost", "revenue", "item"),
                           interest     = NULL, period = c("Annual", "Monthly"),      #Fix
                           params       =list(sd=6, split=0.5)){ #sd
                                                  #portion of costs
        
    if(is.null(name)) name = "Miscellaneous"
    if(is.null(group))group= "Miscellaneous"
    
    if(length(distribution) > 1 | is.null(distribution)) 
      stop("distribution can't be empty or take multiple inputs")
    
    if(from.1 < 0) stop("Period length must be positive")
    
    if(distribution == "uniform"){
        if(to.1 > est.length) stop("Period beyond end date")
        if(from.1 > to.1)     stop("Length of Time is negative")
      
        start = start.date %m+% months(from.1)
        end   = start.date %m+% months(to.1)
        
        len.time= as.numeric(difftime(end, start, units="days"))
        ret.val = xts(rep(1/len.time, len.time), start + 1:len.time)
        
    }else if(distribution == "tiered"){
      if(is.null(from.2)) from.2 = to.1 + 1
      if(is.null(to.2))   to.2   = est.length
      
      if(to.1 > from.2)     stop("Periods overlap")
      if(to.2 > est.length) stop("Period beyond end date")
      
      if(is.null(params)) params = list(split=0.5)
      
      start.1= start.date %m+% months(from.1)
      end.1  = start.date %m+% months(to.1)
      
      start.2= (start.date %m+% months(to.1))
      end.2  = start.date %m+% months(to.2)
      
      len.time.1= as.numeric(difftime(end.1, start.1, units="days"))
      ret.val.1 = xts(rep(1/len.time.1 * params$split, len.time.1), start.1 + 1:len.time.1)
      
      len.time.2= as.numeric(difftime(end.2, start.2, units="days"))
      ret.val.2 = xts(rep(1/len.time.2 * (1 - params$split), len.time.2), start.2 + 1:len.time.2)
      
      ret.val   = rbind(ret.val.1, ret.val.2)  
      
    }else if(distribution == "point"){
      start = start.date %m+% months(from.1)
      end   = start.date %m+% months(to.1)
      
      len.time= as.numeric(difftime(end, start, units="days"))
      ret.val = xts(1, start + len.time)
      
    }else if(distribution == "Scurve"){
      if(to.1 > est.length) stop("Period beyond end date")
      
      start = start.date %m+% months(from.1)
      end   = start.date %m+% months(to.1)

      len.time= as.numeric(difftime(end, start, units="days"))
      vec     = pnorm(1:len.time, mean=len.time/2, sd=params$sd*30) - 
                  pnorm(0:(len.time - 1), mean=len.time/2, sd=params$sd*30) 
      vec     = vec / sum(vec)
      ret.val = xts(vec, start + 1:len.time)
      
    }else stop("No Sufficient distribution required")
    
    if(flag=="cost"){
      flag = -1
    }else if(flag == "revenue"){
      flag = 1
    }else if(flag == "item"){
      flag = 1
    }else stop("Don't recognize flag")
      
    if(!is.null(interest)){
      
      start = start.date
      end   = start.date %m+% months(est.length)
      len.time= as.numeric(difftime(end, start, units="days"))
      
      if(period == "Annual"){
        interval = 365.5
        item     = item * len.time / interval
      }else stop("Period Error")
      
      interest = (1 + interest / interval)^( (1:len.time) / interval) - 1
      interest = xts(interest, start + 1:len.time)
      ret.val = ret.val * interest
    }
      
    ret.val = item * ret.val * flag
    df[[group]][[name]] = ret.val
    return(df)
}

add.item.driver = function(df, 
                           driver,
                           cost.name,
                           cost.group,
                           driver.name, 
                           driver.group){
  
  temp = df[[cost.group]][[cost.name]]
  
  if(is.null(temp)){ 
    temp = df[[as.character(driver.group)]][[as.character(driver.name)]] * 0
  }else df[[cost.group]][[cost.name]] = NULL
  
  df[[cost.group]][[cost.name]] = apply.daily(rbind(temp, df[[as.character(driver.group)]][[as.character(driver.name)]] * driver), sum)
  return(df)
}

nice.merge <- function(df, group, flag=NULL){
  
  if(!is.null(flag)) df[[group]][["Total"]] = NULL
  
  if(length(df[[group]]) == 1){
    item.name = names(df[[group]])
    df[[group]][["Total"]] == df[[group]][[item.name]]
  }else{
    df[[group]][["Total"]] = apply.daily(do.call('rbind', df[[group]]), sum)
  }
  
  return(df)
}


total.merge <- function(df, group){
  
  names = names(df)
  
  for(j in names){
    if(is.null(df[[j]][["Total"]])) next
    
    if(j == names[1]){
      cf =   df[[j]][["Total"]]
      next
    }
    
    cf = apply.daily(merge(df[[j]][["Total"]], cf), sum)
  }
  return(cf)
}

#Financing

#Loan Calc
loan.calc.in <- function(vec.cf,
                            start, 
                            end, 
                            value, 
                            leverage, 
                            interest, 
                            fee,
                            type=c("revolving", "fixed")){
          
  max.loan = leverage * value
  
  if(type == "fixed"){
    
    cum.bal[time(cum.bal) == start] = max.loan
  
  }else if(type == "revolving"){
    
    bal.vec  = vec.cf[time(vec.cf) >= start & time(vec.cf) <= end]
    bal.vec  = xts(pmax(0, bal.vec), time(bal.vec))
    
    new.cum.bal = 
    
    bal.vec  = pmax(pmin(cumsum(bal.vec), max.loan), 0)
    
    
    fee      = xts(fee * max.loan, start)
    interest = bal.vec * interest
    remainder= vec - bal.vec
  }else stop("Type is not recognized")
}

#Equity calc
equity.calc.in <- function(vec.cf, 
                           cum.bal,
                           amount){
  
  cum.bal = xts(cum.bal, time(cum.bal) + 1)
  bal.vec = xts(pmax(vec.cf, 0), time(vec.cf))
  bal.vec = 
  
  bal.vec
}
