rm(list=ls())
source("libs/utils.R")
################################################################################
##Meta Inputs
################################################################################

##Timing
start.date      = as.Date("2016-05-01")

start.period           = 0
start.period.len       = 0
land.aq.len            = 12
construction.len       = 15

est.length             = start.period            +
  start.period.len       +
  land.aq.len            +
  construction.len       

##Interest
cost.interest = 0.05

#Building Info
Total.GFA   = 7399
Total.Units = 17

####
#Need to Calculated total rent in a function

Total.Rent  = 18920 #function to  process data frame of rent
Annual.Operating.Expenses =  62933 #function to do annual operating expenses

df <- list()

################################################################################
##Costs
################################################################################

flag = "cost"

################################################################################
##Land

start = start.period.len
end   = start

GROUP                            = "Land"

Land.Cost         = 922500
Realty.Taxes      = 4500
Land.Transfer.Tax = 28800

#note: first input should be tiered, scurve, or uniform. Not point

df = add.item.fixed(df, 
                    Realty.Taxes,
                    name         ="Realty Taxes", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =est.length,
                    distribution ="uniform",
                    flag         =flag,
                    interest     = cost.interest,
                    period       ="Annual") 

df = add.item.fixed(df, 
                    Land.Cost,
                    name         ="Land Cost", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         = end,
                    distribution ="point",
                    flag         =flag) 

df = add.item.fixed(df, 
                    Land.Transfer.Tax,
                    name         ="Land Transfer Tax", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         = end,
                    distribution ="point",
                    flag         =flag) 

df = nice.merge(df, GROUP)

################################################################################
##Hard Costs

start = start.period.len + land.aq.len
end   = start + construction.len

GROUP                             = "Hard Costs"

Construction.Costs            = 135 * Total.GFA
Environmental.Remediation     = 95000
Site.Servicing                = 40000
Furniture.Fixtures.Equipment = 10000

df = add.item.fixed(df, 
                    Construction.Costs,
                    name         ="Construction Costs", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="Scurve",
                    flag         =flag,
                    params       =list(sd=6)) 


df = add.item.fixed(df, 
                    Environmental.Remediation,
                    name         ="Environmental Remediation", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Site.Servicing,
                    name         ="Site Servicing", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Furniture.Fixtures.Equipment,
                    name         ="Furniture Fixtures Equipment", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = nice.merge(df, GROUP)

################################################################################
##Architects & Engineers

start = start.period.len 
end   = start + land.aq.len + construction.len

GROUP                            ="Architects & Engineers"

Architect = 15000
Structureal.Engineers = 10000
Mechanical.Engineers  = 7500
Electrical.Engineers  = 5000
Civil.Engineers       = 10000 
Traffic.Study         = 5000

df = add.item.fixed(df, 
                    Architect,
                    name         ="Architect", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Structureal.Engineers,
                    name         ="Structureal Engineers", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Mechanical.Engineers,
                    name         ="Mechanical Engineers", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Electrical.Engineers,
                    name         ="Electrical Engineers", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Civil.Engineers,
                    name         ="Civil Engineers", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Traffic.Study,
                    name         ="Traffic Study", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = nice.merge(df, GROUP)

################################################################################
##Other Consultant Costs

start = start.period.len 
end   = start + land.aq.len + construction.len

GROUP                            ="Other Consultant Costs"

Geotechnical.Consultants = 5000
Environmental.Reports    = 17276 + 3297 + 2000
Planning.Zoning          = 20000
Shoring.Consultant       = 5000

df = add.item.fixed(df, 
                    Geotechnical.Consultants,
                    name         ="Geotechnical Consultants", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Environmental.Remediation,
                    name         ="Environmental Reports", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Planning.Zoning,
                    name         ="Planning & Zoning", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Shoring.Consultant,
                    name         ="Shoring Consultant", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = nice.merge(df, GROUP)

################################################################################
##Fees & Permits

start = start.period.len 
end   = start + land.aq.len + construction.len

GROUP                            = "Fees & Permits"

Project.Management.Fees = 3100000 * 0.01
Building.Permit         = 49.83 * Total.Units + Total.GFA * 0.0929 * 16.47 #0.0929 is square feet to square metres
Development.Charges     = 0.05 * 3000000 + 14749 * (Total.Units-1) + #16 res 1 comm unit
  175.78 * 215 / 10.7639
Site.Plan.Approval      = 25000

df = add.item.fixed(df, 
                    Project.Management.Fees,
                    name         ="Project Management Fees", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Building.Permit,
                    name         ="Building Permit", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Development.Charges,
                    name         ="Development Charges", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Site.Plan.Approval,
                    name         ="Site Plan Approval", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = nice.merge(df, GROUP)

################################################################################
##Marketing Legal & Administration

start = start.period.len 
end   = start + land.aq.len + construction.len

GROUP                            = "Marketing Legal & Administration"

Insurance          = 5000
Lease.Up.Commission= 0.75 * Total.Rent
Appraisal          = 5000
Legal              = 20000
Operating.Expenses = Annual.Operating.Expenses / 12 #do 2 months

df = add.item.fixed(df, 
                    Insurance,
                    name         ="Insurance", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Lease.Up.Commission,
                    name         ="Lease Up Commission", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Appraisal,
                    name         ="Appraisal", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Legal,
                    name         ="Legal", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Operating.Expenses,
                    name         ="Operating Expenses", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = nice.merge(df, GROUP)

################################################################################
##Fixed Finance Fees - Combine with Variable Financew once done calculating

start = start.period 
end   = start + land.aq.len + construction.len

GROUP                            = "Fixed Finance Fees"

Lender.Legal.Fees  = 15000
Project.Monitor    = 1800 * 6 + 4500

df = add.item.fixed(df, 
                    Lender.Legal.Fees,
                    name         ="Lender Legal Fees", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = add.item.fixed(df, 
                    Project.Monitor,
                    name         ="Project Monitor", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = nice.merge(df, GROUP)

################################################################################
##Government Taxes

start = start.period 
end   = start + land.aq.len + construction.len

GROUP                            = "Government Taxes"

HST.Monthly.Payables = 0.13
HST.Input.Tax.Credits= -1
HST.Self.Assessment  = 150000  

HST.Applicable.DF = data.frame(Groups = c("Hard Costs",
                                          "Architects & Engineers",
                                          "Other Consultant Costs",
                                          "Fees & Permits",
                                          "Marketing Legal & Administration",
                                          "Marketing Legal & Administration",
                                          "Fixed Finance Fees",
                                          "Fixed Finance Fees"),
                               Name  = c("Total",
                                         "Total",
                                         "Total",
                                         "Project Management Fees",
                                         "Appraisal",
                                         "Legal",
                                         "Lender Legal Fees",
                                         "Project Monitor"))

for(i in 1:nrow(HST.Applicable.DF)){
  
  df = add.item.driver(df, 
                       HST.Monthly.Payables,
                       cost.name   ="HST Monthly Payables",
                       cost.group  = GROUP,
                       driver.name = HST.Applicable.DF$Name[i], 
                       driver.group= HST.Applicable.DF$Groups[i])
  
}

df = add.item.driver(df, 
                     HST.Input.Tax.Credits,
                     cost.name   ="HST Input Tax Credits",
                     cost.group  = GROUP,
                     driver.name = "HST Monthly Payables", 
                     driver.group= GROUP)

df = add.item.fixed(df, 
                    HST.Self.Assessment,
                    name         ="HST Self Assessment", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="uniform",
                    flag         =flag)

df = nice.merge(df, GROUP)

################################################################################
##Revenue
start = start.period + land.aq.len +  construction.len 
end   = start

GROUP                            = "Operating"

Residential.Units = 16
Commercial.Units  = Total.Units - Residential.Units
Bike.Stalls       = 20
Storage.Lockers   = 26

Residential.Avg.Sq.Ft = 347
Commerical.Avg.Sq.Ft  = 215

#Monthly Rents
Residential.Rent.PSF = 3.27
Commercial.Rent.PSF = 3.72
Bike.Stall.Rent      = 15
Locker.Rent          = 45

#calculate Revene
Residential.Revenue = Residential.Units * Residential.Avg.Sq.Ft * Residential.Rent.PSF
Commercial.Revenue  = Commercial.Units *  Commerical.Avg.Sq.Ft  * Commercial.Rent.PSF
Bike.Stall.Revenue  = Bike.Stalls * Bike.Stall.Rent
Locker.Revenue      = Storage.Lockers * Locker.Rent

#Calculate net value
Vacancy.Allowance  = 0.015

Net.Revenue        = (1 - Vacancy.Allowance) * (Residential.Revenue + Commercial.Revenue + Bike.Stall.Revenue + Locker.Revenue)

##costs
Management.Fee     = 0.05 #On Net Revenue
Utilities          = 500  #On Total Units
Repair.Maintenance = 500  #On Total Units
Insurance          = 180  #On Total Units
Wages              = 300  #On Total Units
Admin.Sundry       = 0.015#On Net Revenue
Property.Tax       = 13 * 0.007056037 #On Net Revenue

Monthly.Operating.Expenses = Management.Fee      * Net.Revenue     +
  Utilities          * Total.Units /12 +
  Repair.Maintenance * Total.Units /12 +
  Insurance          * Total.Units /12 +
  Wages              * Total.Units /12 +
  Admin.Sundry       * Net.Revenue     +
  Property.Tax       * Net.Revenue

Net.Operating.Income      = Net.Revenue - Monthly.Operating.Expenses

##cap rate and Sales Costs
Cap.Rate   = 0.0475
Sales.Cost = 0.025

#Project Valuation
Proceeds   = (1 - Sales.Cost) * Net.Operating.Income * 12 / Cap.Rate

df = add.item.fixed(df, #This changes if we are assuming operations
                    Net.Operating.Income * 0,
                    name         ="Operating Income", 
                    group        =GROUP, 
                    from.1       =start - 1,
                    to.1         =end,
                    distribution ="uniform",
                    flag         ="cost")

df = add.item.fixed(df, 
                    Proceeds,
                    name         ="Sales", 
                    group        =GROUP, 
                    from.1       =start,
                    to.1         =end,
                    distribution ="point",
                    flag         ="revenue")

df = nice.merge(df, GROUP)
################################################################################
##Finance Variable Initialize

GROUP = "Variable Finance Fees"

#Project Cash
names = names(df) 

for(j in names){
  if(is.null(df[[j]][["Total"]])) next
  if(j == names[1]){
    total.cash = df[[j]][["Total"]]
    next
  }
  total.cash = apply.daily(rbind(total.cash, df[[j]][["Total"]]), sum)
}


#Cost cash
names = names[!(names == "Operating")]

for(j in names){
  if(is.null(df[[j]][["Total"]])) next
  if(j == names[1]){
    total.fixed.costs = df[[j]][["Total"]]
    next
  }
  total.fixed.costs = apply.daily(rbind(total.fixed.costs, df[[j]][["Total"]]), sum)
}

total.variable.costs = total.fixed.costs * 0

#############Start loop here
total.variable.costs = total.fixed.costs * 0

Land.Loan.1.Interest          = 0.08
Land.Loan.1.Commitment        = 0.02
Land.Loan.1.Leverage          = 0.65

Land.Loan.2.Interest          = 0.12
Land.Loan.2.Commitment        = 0.02
Land.Loan.2.Leverage          = 0.1

Draw.Fees                     = 350
Contruction.Loan.Interest     = 0.065
Construction.Loan.Commitment  = 0.01
Construction.Loan.Leverage    = 0.8


################################################################################
##Financial Engineering

#dates
land.acquisition   = start.date %m+% months(start.period + start.period.len)
construction.start = start.date %m+% months(start.period + start.period.len + land.aq.len)
terminal.period    = start.date %m+% months(start.period + start.period.len + land.aq.len + construction.len)

iter.store = 0

for(kkk in 1:100){
  iter = 0
  
  cf = apply.daily(rbind(total.variable.costs, total.cash), sum)
  
  cf.p = xts(pmax(cf, 0), time(cf))
  cf.n = xts(pmin(cf, 0), time(cf))
  
  #max loan amounts
  max.costs             = -sum(rbind(total.fixed.costs, total.variable.costs))
  max.loan.1            = Land.Cost * Land.Loan.1.Leverage
  max.loan.2            = Land.Cost * Land.Loan.2.Leverage
  max.Construction.Loan = max.costs * Construction.Loan.Leverage
  max.equity            = max.costs - max.Construction.Loan
  
  cf.e.p   = cf.n * 0
  cf.e.n   = cf.n * 0
  cf.e.bal = cf.n * 0
  
  cf.l1.p  = cf.n * 0
  cf.l1.n  = cf.n * 0
  cf.l1.bal= cf.n * 0
  
  cf.l2.p  = cf.n * 0
  cf.l2.n  = cf.n * 0
  cf.l2.bal= cf.n * 0
  
  cf.lc.p  = cf.p * 0
  cf.lc.n  = cf.p * 0
  cf.lc.bal= cf.p * 0
  
  cf.error = cf.p * 0
  
  n = length(cf)
  
  
  for(i in time(cf)){
    
    #print(iter)
    
    iter = iter + 1
    #if(iter > n) break
    
    if(iter == 1){
      
      if(i == land.acquisition){
        cf.l1.n[iter] = max(-max.loan.1, cf.n[iter])
        cf.l2.n[iter] = max(-max.loan.2, cf.n[iter] - cf.l1.n[iter])
      }
      
      cf.e.n[iter] = max(cf.n[iter] - cf.l1.n[iter] - cf.l2.n[iter], -max.equity)
      
      if(i >= construction.start & i <= terminal.period ){
        cf.lc.n[iter] = max(-max.Construction.Loan, cf.n[iter] - cf.e.n[iter] - cf.l1.n[iter] - cf.l2.n[iter])
      }
      
      #We have reached the bottome of the stack.
      #we are in the middle
      
      if(i >= construction.start & i < terminal.period ){
        cf.lc.p[iter] = 0
      }
      
      if(i == construction.start){
        cf.l1.p[iter] = 0
        cf.l2.p[iter] = 0
      }
      
      cf.e.p[iter]    = cf.p[iter]
      
      cf.e.bal[iter]     = cf.e.p[iter] + cf.e.n[iter]
      cf.l1.bal[iter]    = cf.l1.p[iter]+ cf.l1.n[iter]
      cf.l2.bal[iter]    = cf.l2.p[iter]+ cf.l2.n[iter]
      cf.lc.bal[iter]    = cf.lc.p[iter]+ cf.lc.n[iter]
      next
    }
    
    #print("Mark 1")
    #periods ahead of period 1
    if(i == land.acquisition){
      
      cf.l1.n[iter] = min(max(-max.loan.1, xts(cf.l1.bal[iter-1], time(cf.l1.bal[iter])) + cf.n[iter]) - xts(cf.l1.bal[iter-1], time(cf.l1.bal[iter])), 0)
      cf.l2.n[iter] = min(max(-max.loan.2, xts(cf.l2.bal[iter-1], time(cf.l2.bal[iter])) + cf.n[iter] - cf.l1.n[iter]) - xts(cf.l2.bal[iter-1], time(cf.l2.bal[iter])), 0)
      
    }
    
    #print("Mark 2")
    
    if(i == construction.start){
      cf.l1.p[iter] = xts(-cf.l1.bal[iter-1], time(cf.l1.bal[iter])) - cf.l1.n[iter]
      cf.l2.p[iter] = xts(-cf.l2.bal[iter-1], time(cf.l2.bal[iter])) - cf.l2.n[iter]
    }
    
    #print("Mark 3")
    
    #adjsut for land loans paid
    cf.n[iter]   = cf.n[iter] - (cf.l1.p[iter] + cf.l1.n[iter]) - (cf.l2.p[iter] + cf.l2.n[iter])
    
    cf.e.n[iter] = xts(min(max(xts(cf.e.bal[iter-1], time(cf.e.bal[iter])) + cf.n[iter], 
                               -max.equity - xts(cf.e.bal[iter-1], time(cf.e.bal[iter]))) - xts(cf.e.bal[iter-1], time(cf.e.bal[iter])), 0), time(cf.e.bal[iter]))
    
    if(i >= construction.start & i < terminal.period ){
      cf.lc.n[iter] = min(max(-max.Construction.Loan, xts(cf.lc.bal[iter-1], time(cf.lc.bal[iter])) + cf.n[iter] - cf.e.n[iter]) - xts(cf.lc.bal[iter-1], time(cf.lc.bal[iter])), 0)
    }
    
    #print("Mark 4")
    
    ####
    ####
    #Costs should be absorbed. Now account for inflows
    
    if(i >= construction.start & i < terminal.period ){
      cf.lc.p[iter] = max(min(cf.p[iter], xts(-as.numeric(cf.lc.bal[iter - 1]), time(cf.lc.p[iter]))), 0)
      
    }else if(i == terminal.period){
      cf.lc.p[iter] = xts(-as.numeric(cf.lc.bal[iter-1]), time(cf.lc.p[iter])) - cf.lc.n[iter]
    }
    
    #print("Mark 5")
    
    cf.e.p[iter] = cf.p[iter] - cf.lc.p[iter]  
    
    cf.e.bal[iter]     = xts(as.numeric(cf.e.bal[iter-1]),  time(cf.e.bal[iter]))  + cf.e.p[iter] + cf.e.n[iter]
    cf.l1.bal[iter]    = xts(as.numeric(cf.l1.bal[iter-1]), time(cf.l1.bal[iter])) + cf.l1.p[iter]+ cf.l1.n[iter]
    cf.l2.bal[iter]    = xts(as.numeric(cf.l2.bal[iter-1]), time(cf.l2.bal[iter])) + cf.l2.p[iter]+ cf.l2.n[iter]
    cf.lc.bal[iter]    = xts(as.numeric(cf.lc.bal[iter-1]), time(cf.lc.bal[iter])) + cf.lc.p[iter]+ cf.lc.n[iter]
    
  }
  
  land.acquisition   = start.date %m+% months(start.period + start.period.len)
  construction.start = start.date %m+% months(start.period + start.period.len + land.aq.len)
  terminal.period    = start.date %m+% months(start.period + start.period.len + land.aq.len + construction.len)
  
  df[[GROUP]][["Land Loan 1 Interest"]]   = Land.Loan.1.Interest * cf.l1.bal / 365
  df[[GROUP]][["Land Loan 1 Commitment"]] = xts(max(cf.l1.bal) * Land.Loan.1.Commitment, land.acquisition) 
  
  df[[GROUP]][["Land Loan 2 Interest"]]   = Land.Loan.2.Interest * cf.l2.bal / 365
  df[[GROUP]][["Land Loan 2 Commitment"]] = xts(max(cf.l2.bal) * Land.Loan.2.Commitment, land.acquisition) 
  
  df[[GROUP]][["Land Loan 1 Interest"]]   = Contruction.Loan.Interest * cf.lc.bal / 365
  df[[GROUP]][["Land Loan 1 Commitment"]] = xts(max(cf.lc.bal) * Construction.Loan.Commitment, construction.start)
  
  ##################Revise This
  df[[GROUP]][["Draw Fees"]] = df[[GROUP]][["Land Loan 1 Interest"]] * 0
  ##############################
  
  df = nice.merge(df, GROUP, flag=-1)
  total.variable.costs = df[[GROUP]][["Total"]]
  
  new.iter.store = sum(total.variable.costs)
  
  if(abs(iter.store - new.iter.store) < 1){
    print("Solution Found")
    break
  }
  cat("\n Completed ", kkk, " of 100 Tries \n")
  iter.store = new.iter.store
}

##find Total project.cash
names = names(df) 

for(j in names){
  if(is.null(df[[j]][["Total"]])) next
  if(j == names[1]){
    total.project.cash = df[[j]][["Total"]]
    next
  }
  total.project.cash = apply.daily(rbind(total.project.cash, df[[j]][["Total"]]), sum)
}

#Find total equity investor cash
total.equity.cash = cf.e.p + cf.e.n
total.l1          = cf.l1.p + cf.l1.n
total.l2          = cf.l2.p + cf.l2.n

