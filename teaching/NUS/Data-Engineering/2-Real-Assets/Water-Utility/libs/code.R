# library(dplyr)
# library(purrr)
#-----------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------
# Authors: Carlos Arias, Joaquin Calderon, and Rafael Nicolas Fermin Cota
#-----------------------------------------------------------------------------------------
WATER_UTILITY_SAMPLING_FUNC <- function(INPUTS, RISK_TOGGLE){
  #-----------------------------------------------------------------------------------------
  for (j in 1:length(INPUTS)) assign(names(INPUTS)[j], INPUTS[[j]])
  # [1] CONTRACT_LENGTH_VAL (eg., 20)
  # [2] INFLATION_CURRENT_VAL: Current value of inflation. It is the rate of change of the Consumer Price Index. (eg., 0.05)
  # [3] INFLATION_SPEED_ADJUSTMENT_VAL: The rate at which current inflation will tend towards its long term value. (eg., 5)
  # [4] INFLATION_LONG_TERM_VAL: Long term value of inflation. This is the long term rate that the current inflation will tend towards. (eg., 0.08)
  # [5] INFLATION_VOLATILITY_VAL: Uncertainty around the inflation over oneyear period. A 1 volatility means that over the period of one year inflation has the potential to go up or down by 1.  (eg., 0.02)
  # [6] INFLATION_INTIAL_CPI_VAL (eg., 100)
  # [7] INFLATION_INTIAL_DEFLATION_FACTOR_VAL (eg., 1)
  # [8] EXCHANGE_RATE_CURRENT_VAL: Current value of the exchange rate expressed in terms of pesos per dollar. (eg., 100)
  # [9] EXCHANGE_RATE_DEPRECIATION_VAL: Expected annual change in the exchange rate. A positive change is a depreciation of the currency more pesos per dollar and a negative change is an appreciation of the peso. (eg., 0.05)
  # [10] EXCHANGE_RATE_VOLATILITY_VAL: Uncertainty around the exchange rate over oneyear period. A 1 volatility means that over the period of one year the exchange rate has the potential to go up or down by 1.  (eg., 0.05)
  # [11] DISCOUNT_INITIAL_FACTOR_VAL (eg., 1)
  # [12] DISCOUNT_RATE_VAL: Discount rate applied to determine net present values. This is also the Weighted Average Cost of Capital. (eg., 0.07)
  # [13] CUSTOMERS_CURRENT_VAL: Current value of the number of households in the economy. This is the maximum number of connections that the private operator can hope to gain. (eg., 50000)
  # [14] CUSTOMERS_ANNUAL_GROWTH_VAL (eg., 0)
  # [15] SERVICE_COVERAGE_CURRENT_VAL: Current service coverage is the number of households that are currently connected divided by the total number of households. (eg., 0.5)
  # [16] SERVICE_COVERAGE_TARGET_VAL: Coverage extension target coverage. Existing service coverage is 50 (eg., 1)
  # [17] SERVICE_COVERAGE_YEAR_VAL: Year in which service target is to be achieved (eg., 10)
  # [18] DEMAND_CURRENT_VAL: Current demand per connection per day expressed in litres. (eg., 150)
  # [19] DEMAND_OTHER_SOURCES_VAL: Current usage of water from other sources expressed as litres per household per day.This demand is assumed to stay constant. (eg., 75)
  # [20] DEMAND_ANNUAL_GROWTH_VAL: Household demand is assumed to grow at a constant rate with some variability around this. (eg., 0.02)
  # [21] DEMAND_VOLATILITY_VAL (eg., 0.03)
  # [22] DENOMINATIONS_WATER_VOLUMES_VAL (eg., 1000000)
  # [23] NON_REVENUE_WATER_CURRENT_VAL: Nonrevenue water include Unaccounted for water UFW and water used for public services such as firefighting. (eg., 0.01)
  # [24] NON_REVENUE_WATER_TARGET_VAL: Nonrevenue water is assumed to decrease from current level to a target level. The minimum value this can take is 0. (eg., 0.03)
  # [25] NON_REVENUE_WATER_YEAR_VAL: The year in which the target level of nonrevenue water is to be reached. (eg., 5)
  # [26] COLLECTION_RATE_CURRENT_VAL: Collection rate is the proportion of total billed demand that the operator can collect. The remainder is assumed to be bad debt. (eg., 1)
  # [27] COLLECTION_RATE_TARGET_VAL: Collection rate is assumed to gradually increase from current level to a target level. (eg., 1)
  # [28] COLLECTION_RATE_YEAR_VAL: Year in which the collection target is to be achieved. (eg., 10)
  # [29] DENOMINATIONS_MONETARY_VAL (eg., 1000000)
  # [30] INVESTMENT_COST_PER_NEW_CONNECTION_LOCAL_CURRENCY_VAL: Investment cost per connection expressed in pesos. This is a oneoff charge. (eg., 2000)
  # [31] FIXED_OPERATING_COSTS_EXISTING_PERCENT_INVESTMENT_COST_PER_CONNECTION_VAL: Proportion of investment cost assumed to be fixed operating cost. (eg., 0.4)
  # [32] FIXED_OPERATING_COSTS_EXISTING_COST_VAL: Current level of fixed operating cost Proportion of investment assumed to be fixed Investment cost Current level of connections. (eg., 20000000)
  # [33] FIXED_OPERATING_COSTS_ANNUAL_REAL_GROWTH_RATE_VAL: Annual real growth rate of fixed operating costs. (eg., 0.03)
  # [34] FIXED_OPERATING_COSTS_PROPORTION_OF_COSTS_DENOMINATED_IN_FOREIGN_CURRENCY_VAL: Proportion of fixed costs that is assumed to be denominated in dollars and are therefore affected by changes in the exchange rate. (eg., 0.5)
  # [35] VARIABLE_OPERATING_COSTS_EXISTING_COST_VAL: Variable costs per cubic meter. (eg., 20)
  # [36] VARIABLE_OPERATING_COSTS_ANNUAL_REAL_GROWTH_RATE_VAL: Annual real growth rate of operating operating costs. (eg., 0.03)
  # [37] VARIABLE_OPERATING_COSTS_PROPORTION_OF_COSTS_DENOMINATED_IN_FOREIGN_CURRENCY_VAL: Proportion of variable costs that is assumed to be denominated in dollars and are therefore affected by changes in the exchange rate. (eg., 0.5)
  # [38] PERCENTAGE_FUNDED_BY_DEBT_VAL: The costs of coverage extension can be funded through debt for example taking out a loan or through equity that is equity injections by the private operator the contracting authority or both (eg., 0.5)
  # [39] LOAN_FOREIGN_CURRENCY_VAL: If the investment is funded by debt then there is an option of either issuing this loan in domestic currency pesos or foreign currency dollars. (eg., TRUE)
  # [40] LOAN_PERIOD_VAL (eg., 20)
  # [41] LOAN_PERIOD_GRACE_VAL (eg., 5)
  # [42] FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_OPERATOR_VAL: If the private operator pays for all investment then this is set to 100. If the contracting authority is fully responsbile then this is 0.  (eg., 0.5)
  # [43] FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_CONTRACTING_AUTHORITY_VAL (eg., 0.5)
  # [44] INTEREST_RATE_VAL: Real interest rate assumed to remain constant over the life of the contract. (eg., 0.05)
  # [45] ASSET_BASE_INITIAL_LOCAL_CURRENCY_VAL: Initial asset base expressed in pesos. (eg., 100000000)
  # [46] ASSET_BASE_DEPRECIATION_RATE_VAL: Annual depreciation rate. Depreciation charge in each year is Depreciate rate Value of opening assets. (eg., 0.05)
  # [47] TARIFF_REVIEW_PERIOD_VAL: Frequency of tariff resets. (eg., 5)
  # [48] ANNUAL_FIXED_FEES_PAYMENT_LOCAL_CURRENCY_VAL: Lease payments concession fees and license fees can be represented by negative annual payments i.e. payments from the operator to the contracting authority. (eg., 0)
  # [49] EXISTING_TARIFF_CONNECTED_VAL: Tariff charged to households that are connected to the water system prior to the introduction of the private operator expressed in terms of pesos per cubic meter. (eg., 30)
  # [50] EXISTING_TARIFF_OTHER_VAL: Tariff charged to households that obtain water from other methods e.g. water tankers river sources expressed in terms of pesos per cubic meter. (eg., 45)
  # [51] EXISTING_TARIFF_COPYING_COST_VAL: Additional cost incurred by unconnected households in order to obtain water such as time traveled to river source cost of installing water storage. (eg., 1000)
  # [52] WTP_CONNECTION_VAL: Willingness to pay per cubic meter for a household connections versus other sources expressed as a percentage of what is currently charged. (eg., 2.5)
  # [53] WTP_OTHER_VAL: Willingness to pay per cubic meter for a household connections versus other sources expressed as a percentage of what is currently charged. (eg., 1.5)
  # [54] WTP_COPYING_COST_VAL: Willingness to incur coping costs to remain unconnected to the water system. (eg., 1.3)
  #-----------------------------------------------------------------------------------------
  
  HEADINGS_STR = "PERIODS,I) ECONOMIC FACTORS: INFLATION,I) ECONOMIC FACTORS: CONSUMER PRICE INDEX,I) ECONOMIC FACTORS: DEFLATION FACTOR,I) ECONOMIC FACTORS: REAL EXCHANGE RATE,I) ECONOMIC FACTORS: NOMINAL EXCHANGE RATE,I) ECONOMIC FACTORS: DISCOUNT FACTOR,"
  HEADINGS_STR = paste0(HEADINGS_STR, "II) CONSUMPTION: POTENTIAL CONNECTIONS,II) CONSUMPTION: COVERAGE,II) CONSUMPTION: CONNECTIONS,II) CONSUMPTION: NEW CONNECTIONS,II) CONSUMPTION: DEMAND (L/CONNECTION/DAY),II) CONSUMPTION: TOTAL DEMAND (M3),II) CONSUMPTION: NON-REVENUE WATER,II) CONSUMPTION: COLLECTION RATE FORECAST,")
  HEADINGS_STR = paste0(HEADINGS_STR, "III) COSTS & SUBSIDIES: OPERATING FIXED COSTS - DOMESTIC DENOMINATION,III) COSTS & SUBSIDIES: OPERATING FIXED COSTS - FOREIGN DENOMINATION,III) COSTS & SUBSIDIES: OPERATING VARIABLE COSTS - DOMESTIC DENOMINATION,III) COSTS & SUBSIDIES: OPERATING VARIABLE COSTS - FOREIGN DENOMINATION,")
  HEADINGS_STR = paste0(HEADINGS_STR, "IV) FINANCING COSTS: COVERAGE EXTENSION COSTS,IV) FINANCING COSTS: AMOUNT FINANCED BY DEBT,IV) FINANCING COSTS: TOTAL LOAN REPAYMENT (FOREIGN CURRENCY),IV) FINANCING COSTS: OUTSTANDING PRINCIPAL,IV) FINANCING COSTS: PRINCIPAL REPAYMENT,IV) FINANCING COSTS: INTEREST PAYMENT,IV) FINANCING COSTS: OPERATOR - OUTSTANDING PRINCIPAL,IV) FINANCING COSTS: OPERATOR - PRINCIPAL REPAYMENT,IV) FINANCING COSTS: OPERATOR - INTEREST PAYMENT,IV) FINANCING COSTS: CONTRACTING AUTHORITY - OUTSTANDING PRINCIPAL,IV) FINANCING COSTS: CONTRACTING AUTHORITY - PRINCIPAL REPAYMENT,IV) FINANCING COSTS: CONTRACTING AUTHORITY - INTEREST PAYMENT,")
  
  HEADINGS_STR = paste0(HEADINGS_STR, "V) TARIFF RESET CALCULATIONS: EXPECTED INFLATION RATE - FORECAST USED FOR TARIFF SETTING,V) TARIFF RESET CALCULATIONS: EXPECTED INFLATION RATE - DEFLATION FACTOR,V) TARIFF RESET CALCULATIONS: ASSET BASE (NOMINAL) - INITIAL EXISTING ASSET BASE,V) TARIFF RESET CALCULATIONS: EXISTING ASSET BASE (NOMINAL) - DEPRECIATION,V) TARIFF RESET CALCULATIONS: ASSET BASE (NOMINAL) - MAINTENANCE COSTS,V) TARIFF RESET CALCULATIONS: ASSET BASE (NOMINAL) - FINAL EXISTING ASSET BASE,V) TARIFF RESET CALCULATIONS: ASSET BASE (NOMINAL) - INITIAL NEW ASSET BASE,V) TARIFF RESET CALCULATIONS: ASSET BASE (NOMINAL) - INVESTMENT,V) TARIFF RESET CALCULATIONS: NEW ASSET BASE (NOMINAL) - DEPRECIATION,V) TARIFF RESET CALCULATIONS: ASSET BASE (NOMINAL) - FINAL NEW ASSET BASE,")
  HEADINGS_STR = paste0(HEADINGS_STR, "V) TARIFF RESET CALCULATIONS: DEMAND - NO RESET,V) TARIFF RESET CALCULATIONS: DEMAND - FULL RESET,V) TARIFF RESET CALCULATIONS: DEMAND - DISCOUNTED DEMAND,V) TARIFF RESET CALCULATIONS: DEMAND - CUMULATIVE DISCOUNTED DEMAND,V) TARIFF RESET CALCULATIONS: DEMAND - ", TARIFF_REVIEW_PERIOD_VAL, " YEAR DISCOUNTED DEMAND,V) TARIFF RESET CALCULATIONS: RETURN ON NEW ASSETS (REAL) - CONTRACTING AUTHORITY,V) TARIFF RESET CALCULATIONS: RETURN ON NEW ASSETS (REAL) - OPERATOR,V) TARIFF RESET CALCULATIONS: DEPRECIATION (REAL) - CONTRACTING AUTHORITY,V) TARIFF RESET CALCULATIONS: DEPRECIATION (REAL) - OPERATOR,V) TARIFF RESET CALCULATIONS: MAINTENANCE COSTS (REAL),")
  HEADINGS_STR = paste0(HEADINGS_STR, "V) TARIFF RESET CALCULATIONS: DOMESTIC-DENOMINATED OPEX (REAL),V) TARIFF RESET CALCULATIONS: FOREIGN-DENOMINATED OPEX (REAL),V) TARIFF RESET CALCULATIONS: FOREIGN EXCHANGE RATE - NO RESET,V) TARIFF RESET CALCULATIONS: FOREIGN EXCHANGE RATE - FULL RESET,V) TARIFF RESET CALCULATIONS: SUBSIDIES (REAL),V) TARIFF RESET CALCULATIONS: REVENUE REQUIREMENT (REAL) - CONTRACTING AUTHORITY - REVENUE REQUIREMENT,V) TARIFF RESET CALCULATIONS: REVENUE REQUIREMENT (REAL) - CONTRACTING AUTHORITY - DISCOUNTED REVENUE REQUIREMENT,V) TARIFF RESET CALCULATIONS: REVENUE REQUIREMENT (REAL) - CONTRACTING AUTHORITY - CUMULATIVE DISCOUNTED REVENUE REQUIREMENT,V) TARIFF RESET CALCULATIONS: REVENUE REQUIREMENT (REAL) - OPERATOR - REVENUE REQUIREMENT,V) TARIFF RESET CALCULATIONS: REVENUE REQUIREMENT (REAL) - OPERATOR - DISCOUNTED REVENUE REQUIREMENT,")
  HEADINGS_STR = paste0(HEADINGS_STR, "V) TARIFF RESET CALCULATIONS: REVENUE REQUIREMENT (REAL) - OPERATOR - CUMULATIVE DISCOUNTED REVENUE REQUIREMENT,V) TARIFF RESET CALCULATIONS: DISCOUNTED REVENUE REQUIREMENT - CONTRACTING AUTHORITY,V) TARIFF RESET CALCULATIONS: DISCOUNTED REVENUE REQUIREMENT - OPERATOR,V) TARIFF RESET CALCULATIONS: REVENUE REQUIRED PER M3 - NO RESET - CONTRACTING AUTHORITY,V) TARIFF RESET CALCULATIONS: REVENUE REQUIRED PER M3 - NO RESET - OPERATOR,V) TARIFF RESET CALCULATIONS: REVENUE REQUIRED PER M3 - ", TARIFF_REVIEW_PERIOD_VAL, " YEAR RESET PERIOD - CONTRACTING AUTHORITY,V) TARIFF RESET CALCULATIONS: REVENUE REQUIRED PER M3 - ", TARIFF_REVIEW_PERIOD_VAL, " YEAR RESET PERIOD - OPERATOR,")
  
  HEADINGS_STR = paste0(HEADINGS_STR, "VI) TARIFF INDEXATION: INDEXATION,VI) TARIFF INDEXATION: NO RESET - CONTRACTING AUTHORITY,VI) TARIFF INDEXATION: NO RESET - OPERATOR,VI) TARIFF INDEXATION: ", TARIFF_REVIEW_PERIOD_VAL, " YEAR RESET PERIOD - CONTRACTING AUTHORITY,VI) TARIFF INDEXATION: ", TARIFF_REVIEW_PERIOD_VAL, " YEAR RESET PERIOD - OPERATOR,")
  HEADINGS_STR = paste0(HEADINGS_STR, "VII) PROFIT: TOTAL REVENUE COLLECTED - NO RESET,VII) PROFIT: TOTAL REVENUE COLLECTED - ", TARIFF_REVIEW_PERIOD_VAL, " YEAR RESET PERIOD,VII) PROFIT: COSTS - OPERATING COSTS,VII) PROFIT: COSTS - MAINTENANCE COSTS,VII) PROFIT: COSTS - RETURN ON CAPITAL,VII) PROFIT: COSTS - DEPRECIATION,VII) PROFIT: COSTS - SUBSIDIES,VII) PROFIT: COSTS - TOTAL,VII) PROFIT: PRESENT VALUE PROFIT (REAL) - NO RESET,VII) PROFIT: PRESENT VALUE PROFIT (REAL) - ", TARIFF_REVIEW_PERIOD_VAL, " YEAR RESET PERIOD,")
  HEADINGS_STR = paste0(HEADINGS_STR, "VIII) CASH FLOWS: NO RESET - CONTRACTING AUTHORITY,VIII) CASH FLOWS: NO RESET - OPERATOR,VIII) CASH FLOWS: ", TARIFF_REVIEW_PERIOD_VAL, " YEAR RESET PERIOD - CONTRACTING AUTHORITY,VIII) CASH FLOWS: ", TARIFF_REVIEW_PERIOD_VAL, " YEAR RESET PERIOD - OPERATOR,")
  HEADINGS_STR = paste0(HEADINGS_STR, "IX) DEBT SERVICE RATIO: CASH FLOW BEFORE LOAN REPAYMENT - NO RESET,IX) DEBT SERVICE RATIO: CASH FLOW BEFORE LOAN REPAYMENT - ", TARIFF_REVIEW_PERIOD_VAL, " YEAR RESET PERIOD,IX) DEBT SERVICE RATIO: DEBT SERVICE RATIO - NO RESET,IX) DEBT SERVICE RATIO: DEBT SERVICE RATIO - ", TARIFF_REVIEW_PERIOD_VAL, " YEAR RESET PERIOD,")
  
  # Compares welfare change for currently connected and currently unconnected households.
  HEADINGS_STR = paste0(HEADINGS_STR, "X) STAKEHOLDER ANALYSIS: WATER USE & CONSUMPTION - PRIVATE HOUSEHOLD CONNECTION,X) STAKEHOLDER ANALYSIS: WATER USE & CONSUMPTION - OTHER SOURCES,X) STAKEHOLDER ANALYSIS: MONTHLY COST - PRIVATE HOUSEHOLD CONNECTION - ", TARIFF_REVIEW_PERIOD_VAL, " YEAR RESET PERIOD,X) STAKEHOLDER ANALYSIS: MONTHLY COST - PRIVATE HOUSEHOLD CONNECTION - EXISTING TARIFFS,X) STAKEHOLDER ANALYSIS: OTHER SOURCES,X) STAKEHOLDER ANALYSIS: MONTHLY COPING COST - OTHER SOURCES,X) STAKEHOLDER ANALYSIS: WILLINGNESS TO PAY - PRIVATE HOUSEHOLD CONNECTION,X) STAKEHOLDER ANALYSIS: WILLINGNESS TO PAY - OTHER SOURCES - MONTHLY COST,X) STAKEHOLDER ANALYSIS: WILLINGNESS TO PAY - OTHER SOURCES - MONTHLY COPING COST,X) STAKEHOLDER ANALYSIS: CHANGE IN SOCIAL WELFARE (REAL) - PRIVATE HOUSEHOLD CONNECTION,X) STAKEHOLDER ANALYSIS: CHANGE IN SOCIAL WELFARE (REAL) - OTHER SOURCES,")
  #-----------------------------------------------------------------------------------------
  
  
  nm=strsplit(HEADINGS_STR, ",")[[1]]
  # Assign the number of CONTRACT_LENGTH_VAL + 1 to nrows
  nrows=CONTRACT_LENGTH_VAL + 1
  # Assign the length of vector nm to ncols
  ncols=length(nm)
  # Create a matrix with with predetermined values of 0. The number of rows is equal to nrows and the number of columns is equal to ncols
  mat=data.frame(matrix(0,nrows, ncols))
  # Assign names from mat to the variable nm
  names(mat)=nm
  
  # Assign value of 0 to columns from 20 to 68
  mat[,20:68]=0
  
  #-----------------------------------------------------------------------------------------
  
  # Assign value of 0 to the cell of column 1 and row 1
  mat[1,1]=0
  # Assign value of input INFLATION_CURRENT_VAL to the cell of row 1 and column 2
  mat[1,2]=INFLATION_CURRENT_VAL #Inflation_Cur
  # Assign value of input INFLATION_INITIAL_CPI_VAL to the cell of row 1 and column 3
  mat[1,3]=INFLATION_INTIAL_CPI_VAL #CPI_Cur
  # Assign value of input INFLATION_CURRENT_VAL to the cell of row 1 and column 4
  mat[1,4]=INFLATION_INTIAL_DEFLATION_FACTOR_VAL
  # Assign value of input EXCHANGE_RATE_CURRENT_VAL to the cell of row 1 and column 5
  mat[1,5]=EXCHANGE_RATE_CURRENT_VAL
  # Assign value of input of the multiplication of row 1 and column 5 by (1 + row 1 and column 2)
  mat[1,6]=mat[1,5] * (1 + mat[1,2])
  # Assign value of input DISCOUNT_INITIAL_FACTOR_VAL to the cell of row 1 and column 7
  mat[1,7]=DISCOUNT_INITIAL_FACTOR_VAL
  # Assign value of input CUSTOMERS_CURRENT_VAL to the cell of row 1 and column 8
  mat[1,8]=CUSTOMERS_CURRENT_VAL
  # Assign value of input CUSTOMERS_CURRENT_VAL to the cell of row 1 and column 9
  mat[1,9]=SERVICE_COVERAGE_CURRENT_VAL
  # Assign value as an integer the multiplication of row 1 and column 8 and row 1 and column 9
  mat[1,10]=as.integer(mat[1,8] * mat[1,9])
  # Assign value of 0 to the cell of row 1 and column 11
  mat[1,11]=0 #New connections
  # Assign value of input DEMAND_CURRENT_VAL to the cell of row 1 and column 12
  mat[1,12]=DEMAND_CURRENT_VAL
  # Assign value of cell of row 1 and column 12 to the cell of row 1 and column 93
  mat[1,92]=mat[1,12]
  # Assign value of input DEMAND_OTHER_SOURCES_VAL to the cell of row 1 and column 93
  mat[1,93]=DEMAND_OTHER_SOURCES_VAL
  
  # Assign value of the multiplication of row 1 and column 12 times row 1 and column 10 times 365 divided by 365 and also divided by 1000 and 
  # finally divided by the input DENOMINATIONS_WATER_VOLUMES_VAL
  mat[1,13]=mat[1,12] * mat[1,10] * 365 / 1000 / DENOMINATIONS_WATER_VOLUMES_VAL
  # Assign value of the input NON_REVENUE_WATER_CURRENT_VAL to the cell of row 1 and column 14
  mat[1,14]=NON_REVENUE_WATER_CURRENT_VAL
  # Assign value of the input COLLECTION_RATE_CURRENT_VAL to the cell of row 1 and column 15
  mat[1,15]=COLLECTION_RATE_CURRENT_VAL
  
  #Opex_Fixed
  # Multiplication of 4 inputs and assign it to the variable TEMP_VAL
  TEMP_VAL = FIXED_OPERATING_COSTS_EXISTING_PERCENT_INVESTMENT_COST_PER_CONNECTION_VAL *
    INVESTMENT_COST_PER_NEW_CONNECTION_LOCAL_CURRENCY_VAL *
    SERVICE_COVERAGE_CURRENT_VAL *
    CUSTOMERS_CURRENT_VAL
  
  # Assign value of the results of the permutation below to the cell of row 1 and column 16
  mat[1,16]=(TEMP_VAL / DENOMINATIONS_MONETARY_VAL) * 
    (1 - FIXED_OPERATING_COSTS_PROPORTION_OF_COSTS_DENOMINATED_IN_FOREIGN_CURRENCY_VAL)
  # Assign value of the results of the permutation below to the cell of row 1 and column 17
  mat[1,17]=(TEMP_VAL / DENOMINATIONS_MONETARY_VAL) * 
    FIXED_OPERATING_COSTS_PROPORTION_OF_COSTS_DENOMINATED_IN_FOREIGN_CURRENCY_VAL / EXCHANGE_RATE_CURRENT_VAL
  # Assign value of the results of the permutations below to the cell of row 1 and column 18
  mat[1,18]=VARIABLE_OPERATING_COSTS_EXISTING_COST_VAL * 
    (1 - VARIABLE_OPERATING_COSTS_PROPORTION_OF_COSTS_DENOMINATED_IN_FOREIGN_CURRENCY_VAL)
  # Assign value of the results of the permutations below to the cell of row 1 and column 19
  mat[1,19]=VARIABLE_OPERATING_COSTS_EXISTING_COST_VAL * 
    VARIABLE_OPERATING_COSTS_PROPORTION_OF_COSTS_DENOMINATED_IN_FOREIGN_CURRENCY_VAL / EXCHANGE_RATE_CURRENT_VAL
  
  
  #-----------------------------------------------------------------------------------------
  #-----------------------------------------------------------------------------------------
  # Assign to i the numbers from 2 to nrows
  # Assign to j, i minus 1
  # Assign the vector of i without the first value
  i = 2:nrows; j = i - 1; k = i[-1]
  
  # Assign to rows from 2 to nrows for column 1 the value of j
  mat[i, 1] = j
  
  # TARIFF_REVIEW_PERIOD_VAL is equal to 1 or the reminder of j divided by TARIFF_REVIEW_PERIOD_VAL is equal to 1, Output TRUE or FALSE
  TFLAG1_VAL = ((TARIFF_REVIEW_PERIOD_VAL == 1) | (j %% TARIFF_REVIEW_PERIOD_VAL == 1))
  # Assign the reminder of j divided by TARIFF_REVIEW_PERIOD_VAL is equal to 1, Output TRUE or FALSE
  TFLAG2_VAL = (j %% TARIFF_REVIEW_PERIOD_VAL == 1)
  
  #-----------------------------------------------------------------------------------------
  # If RISK_TOGGLE is not equal to true
  if (!RISK_TOGGLE){ # 'With No-Risk
    #-----------------------------------------------------------------------------------------
    # Assign to rows from 2 to nrows for column 2 the function of Reduce
    mat[i, 2]=Reduce(
      # Inflation: Current value of inflation. It is the rate of change of the Consumer
      # Price Index.
      
      # Long term value: Long term value of inflation. This is the long term rate that the
      # current inflation will tend towards.
      
      # Volatility: Uncertainty around the inflation over one-year period. A 1% volatility
      # means that over the period of one year, inflation has the potential to go up or down by 1%.
      
      # Speed of adjustment: The rate at which current inflation will tend towards its long
      # term value
      # Annual growth rate
      
      # Create a function with two inputs
      # The function will take the first input prv and multiply by the exponential raised to the 
      # exponents of (INFLATION_SPEED_ADJUSTMENT_VAL multiplied by (INFLATION_LONG_TERM_VAL-prv)) 
      function(prv,nxt) prv*exp(INFLATION_SPEED_ADJUSTMENT_VAL*(INFLATION_LONG_TERM_VAL-prv)), 
      # Creates a vector that has a value of 1 repated (nrows-1) times and applies the function above 
      rep(1, nrows-1), 
      # The function will initialize at row 1 and column 2
      mat[1, 2], 
      # The results will be accumulated
      accumulate = TRUE
      # Apply it after the first index of the column 
    )[-1]
    
    # Assign to row from 2 to nrows and column 5
    # The cummulative product of row 1 and column 5 times the exponential value raised to the exponent of EXCHANGE_RATE_DEPRECIATION_VAL 
    # this exponential will be repeated nrows-1 times, and skip the first index. 
    mat[i, 5]=cumprod(c(mat[1, 5], rep(exp(EXCHANGE_RATE_DEPRECIATION_VAL), nrows-1)))[-1]
    #mat[i, 5]=accumulate(
    #  .x=rep(exp(EXCHANGE_RATE_DEPRECIATION_VAL), nrows-1), 
    #  .f=function(prv,nxt) prv * nxt, 
    #  .init=mat[1, 2]
    #)[-1]
    
    # Assign from 2 to nrows and column 12
    # The cummulative product of row 1 and column 12 times the exponential value raised to the exponent of DEMAND_ANUAL_GROWTH_VAL
    # this exponential will be repeated nrows-1 times, and skip the first index. 
    mat[i, 12]=cumprod(c(mat[1, 12], rep(exp(DEMAND_ANNUAL_GROWTH_VAL), nrows-1)))[-1]
    # Assign the values from 2 to nrwos and column 12, to the rows from 2 to nrows and column 92
    mat[i, 92]= mat[i, 12]
    # Assign from 2 to nrows and column 93
    # The cummulative product of row 1 and column 93 times the exponential value raised to the exponent of DEMAND_ANUAL_GROWTH_VAL
    # this exponential will be repeated nrows-1 times, and skip the first index. 
    mat[i, 93]=cumprod(c(mat[1, 93], rep(exp(DEMAND_ANNUAL_GROWTH_VAL), nrows-1)))[-1]
    
    #-----------------------------------------------------------------------------------------
    # Else 
  } else { # with risk
    #-----------------------------------------------------------------------------------------
    # Assign 21 random numbers to the first random variable
    rnd1=c(-1.63530374963851, -0.68541620176237, 0.842126679004853, 2.24300314945776, -1.30018563892108, -0.399562916625586, -0.492439250657473, 0.352356975065116, 0.0346640557438982, 2.34607539004499, -0.139182297384679, -1.36517849320719, 2.55653827914048, 0.915935927357113, 0.285107591845216, -0.40912433095963, 0.419106733017931, -1.00891713566328, 0.565359609265667, 0.166988419109934, 1.92580344598388)
    # Assign 21 random numbers to the second random variable
    rnd2=c(1.75556803913375, -0.0799365289157126, 1.63025368036105, 0.268892180257083, -0.0817688484778561, 0.700559159611574, -0.668244340991706, 0.216952495929201, -0.962442548216423, -0.717517445102393, 0.0986384764455445, -0.667207935530494, -0.619421357051072, -0.991318012291069, 0.634356804887895, 1.06147409139581, 0.495289926518602, -0.76369614118954, 0.504454895729237, 0.120628984655496, -1.13045729301784)
    # Assign 21 random numbers to the third random variable
    rnd3=c(-1.04221645740392, -0.629814708524595, -0.447910219752754, -0.0370163283177808, 0.281629466353362, 0.775422988816911, -0.763431339408918, -0.651970442137781, 0.163985782433124, -0.723624801741545, 1.26790409376486, 0.0320993689595026, 0.418757422595019, -0.539957290211319, 0.41082096870385, -2.15717521594011, -0.880684029801793, 0.425257055847261, 0.112015247259036, -0.581296892902507, 0.724839944832975)
    # Assign 21 random numbers to the fourth random variable
    rnd4=c(-1.04221645740392, -0.629814708524595, -0.447910219752754, -0.0370163283177808, 0.281629466353362, 0.775422988816911, -0.763431339408918, -0.651970442137781, 0.163985782433124, -0.723624801741545, 1.26790409376486, 0.0320993689595026, 0.418757422595019, -0.539957290211319, 0.41082096870385, -2.15717521594011, -0.880684029801793, 0.425257055847261, 0.112015247259036, -0.581296892902507, 0.724839944832975)
    
    # set.seed(1); rnorm(nrows)
    # set.seed(1); qnorm(runif(nrows)) # NORMSINV(RAND()))
    
    # Assign to rows from 2 to nrows and 2 column With Risk
    mat[i, 2]=Reduce(
      # Inflation: Current value of inflation. It is the rate of change of the Consumer
      # Price Index.
      
      # Long term value: Long term value of inflation. This is the long term rate that the
      # current inflation will tend towards.
      
      # Volatility: Uncertainty around the inflation over one-year period. A 1% volatility
      # means that over the period of one year, inflation has the potential to go up or down by 1%.
      
      # Speed of adjustment: The rate at which current inflation will tend towards its long
      # term value
      
      # Annual growth rate
      
      
      # Create a function with two inputs
      function(prv,nxt) {
        # Assign to a variable the following calculations done below with the inputs from the original function and 
        # the function created above
        EXPECTED_CHANGE_VAL = INFLATION_SPEED_ADJUSTMENT_VAL * (INFLATION_LONG_TERM_VAL - prv)
        prv * exp(EXPECTED_CHANGE_VAL - INFLATION_VOLATILITY_VAL * 
                    INFLATION_VOLATILITY_VAL / 2 + INFLATION_VOLATILITY_VAL * nxt)
      }, 
      # Apply the function above to the vector rnd1 without the first index
      rnd1[-1], 
      # Start at row 1 and column 2
      mat[1, 2], 
      # accumulate the result 
      accumulate = TRUE
      # Apply it after the first index of the column 
    )[-1]
    
    # Assign to row 2 to nrow and column 5
    mat[i, 5]=Reduce(
      # Expected annual change in the exchange rate. A positive change is a depreciation of
      # the currency (more pesos per dollar), and a negative change is an appreciation of
      # the peso.
      
      # Create a function 
      function(prv,nxt) {
        # Do the following calculations with the values from the original input and the function created above
        prv * exp(EXCHANGE_RATE_DEPRECIATION_VAL - EXCHANGE_RATE_VOLATILITY_VAL * 
                    EXCHANGE_RATE_VOLATILITY_VAL / 2 + EXCHANGE_RATE_VOLATILITY_VAL * nxt)
      }, 
      # Apply the function created above to the numbers from the vector rnd2 and without the first index
      rnd2[-1], # next
      # Initial value
      mat[1, 5], 
      # Accumulate the results
      accumulate = TRUE
      # Skip the first index
    )[-1]
    
    # Assign to the rows from 2 to nrows and columns 12
    mat[i, 12]=Reduce(
      # Household demand is assumed to grow at a constant rate, with some variability
      # around this.
      
      # Create a function with two inputs
      function(prv,nxt) {
        # Do the following calculations with the original inputs and the inputs from the function created above 
        prv * exp(DEMAND_ANNUAL_GROWTH_VAL - DEMAND_VOLATILITY_VAL * DEMAND_VOLATILITY_VAL/2 + 
                    DEMAND_VOLATILITY_VAL*nxt)
      }, 
      # Apply the function above to the vector rnd3 without the first index
      rnd3[-1], # next
      # Initial value
      mat[1, 12], 
      # Accumulate the results
      accumulate = TRUE
      # Skip the first index
    )[-1]
    
    #mat[i, 32]=mat[i, 2]
    #mat[i, 33]=1
    
    # Assign the values from 2 to nrwos and column 12 to column 92 and 2 to nrows
    mat[i, 92]=mat[i, 12]
    # Assign the values to 2 to nrows, and column 93
    mat[i, 93]=Reduce(
      # Other household demand is assumed to grow at a constant rate, with some variability
      # around this.
      
      # Create a function
      function(prv,nxt) {
        # Does the following calculation below with the inputs from the original function and the inputs of the function created above
        prv * exp(DEMAND_ANNUAL_GROWTH_VAL - DEMAND_VOLATILITY_VAL * DEMAND_VOLATILITY_VAL/2 + 
                    DEMAND_VOLATILITY_VAL*nxt)
      }, 
      # Apply the function created above to the vector rnd4 without the first index
      rnd4[-1], # next
      # Initial Value
      mat[1, 93],
      # Accumulate the results
      accumulate = TRUE
    )[-1]
    
    #-----------------------------------------------------------------------------------------
  }
  #-----------------------------------------------------------------------------------------
  
  # Assign the following values to the the matrix to the cells from 2 to nrwos and column 3
  mat[i, 3] = Reduce(
    # Creates a function that takes two inputs and perform the following calculation
    function(prv,nxt) {prv * exp(nxt)}, 
    # Apply the function created above to the vector below
    mat[i, 2], # next
    # Initial Value
    mat[1, 3], # initial value
    # Accumulate the results 
    accumulate = TRUE
    # Skip the first index
  )[-1]
  
  # Assign from to the rows 2 to nrows and column 4
  mat[i, 4] = Reduce(
    # Creates a function that takes two inputs and applys the following function
    function(prv,nxt) {prv * exp(-nxt)}, 
    # Apply the function created above the vector below 
    mat[i, 2], # next
    # Initial Value
    mat[1, 3], 
    # Accumulate the results 
    accumulate = TRUE
    # Skip the first index and divide by 100 the results
  )[-1]/100
  
  #-----------------------------------------------------------------------------------------
  # Assign to the rows from 2 to nrows, and column 6 the following calculations below
  mat[i, 6]=mat[i, 5] * (1 + mat[i, 2])
  # Assign to the rows from 2 to nrows, and column 7 the following calculations below
  # The cummulative product of row 1 and column 7 times the exponential value raised to the exponent of negative DISCOUNT_RATE_VAL
  # this exponential will be repeated nrows-1 times, and skip the first index. 
  mat[i, 7]=cumprod(c(mat[1, 7], rep(exp(-DISCOUNT_RATE_VAL), nrows-1)))[-1]
  # Assign to the rows from 2 to nrows, and column 8 the following calculations below
  # The cummulative product of row 1 and column 8 times the exponential value raised to the exponent of CUSTOMERS_ANNUAL_GROWTH_VAL
  # this exponential will be repeated nrows-1 times, and skip the first index. 
  mat[i, 8]=cumprod(c(mat[1, 8], rep(exp(CUSTOMERS_ANNUAL_GROWTH_VAL), nrows-1)))[-1]
  #-----------------------------------------------------------------------------------------
  
  # Assign to the rows from 2 to nrows and column 9 if the conditions are met
  mat[i, 9] = ifelse(
    # If the values from 2 to nrows and column 1 are less than the values of SERVICE_COVERGE_YEAR_VAL
    mat[i, 1] < SERVICE_COVERAGE_YEAR_VAL, 
    # If the following stated above is true apply the following calculations
    # grow from existing levels to a Target level following an S-shaped curve.
    SERVICE_COVERAGE_CURRENT_VAL + (SERVICE_COVERAGE_TARGET_VAL - SERVICE_COVERAGE_CURRENT_VAL) / 
      (1 + exp(-(10 / SERVICE_COVERAGE_YEAR_VAL) * (mat[i, 1] - 1 - (SERVICE_COVERAGE_YEAR_VAL - 1) / 2))),
    # If the following stated above is false apply the following calculations
    SERVICE_COVERAGE_TARGET_VAL
  )
  
  #-----------------------------------------------------------------------------------------
  # Assign to the rows from 2 to nrwos and column 10, as integer the multiplication of rows from 2 to nrwos from column 8 and 9
  mat[i, 10]=as.integer(mat[i, 8] * mat[i, 9])
  # New connections over a year is the difference between Connections in the current
  # period and Connections in the previous period.
  
  # If the difference between each row from column ten is greater than 0
  # If the operation above is true assign the values of the difference
  # If the operation above is false assign 0
  mat[i, 11]=ifelse(diff(mat[, 10])>0, diff(mat[, 10]), 0)
  
  # Total demand, expressed in millions of cubic meters is: Demand * 365 days / 1000 / 1000000
  # Assign to the matrix from rows 2 to nrwos and the column 13
  # Assign the multiplication of 2 to nrows and column 12 and 13 and the result divided by 365 then by 1000 and then by DENOMINATIONS_WATER_VOLUMES_VAL
  mat[i, 13] = mat[i, 12] * mat[i, 10] * 365 / 1000 / DENOMINATIONS_WATER_VOLUMES_VAL
  #-----------------------------------------------------------------------------------------
  
  # names(mat)[14]
  # Assign to the rows from 2 to nrows and column 14
  mat[i, 14] = ifelse(
    # If the values from 2 to nrwos and column 1 are less than NON_REVENUE_WATER_YEAR_VAL
    mat[i, 1] < NON_REVENUE_WATER_YEAR_VAL, 
    # grow from existing levels to a Target level following an S-shaped curve.
    # Perform the following calculations encounter below if the statement above is true
    NON_REVENUE_WATER_CURRENT_VAL + (NON_REVENUE_WATER_TARGET_VAL - NON_REVENUE_WATER_CURRENT_VAL) / 
      (1 + exp(-(10 / NON_REVENUE_WATER_YEAR_VAL) * (mat[i, 1] - 1 - (NON_REVENUE_WATER_YEAR_VAL - 1) / 2))),
    # Assign the following value if the statement above is false
    NON_REVENUE_WATER_TARGET_VAL
  )
  
  # Assign to the rows from 2 to nrwos and column 15
  mat[i, 15] = ifelse(
    # If the values from 2 to nrwos and column 1 are less than COLLECTION_RATE_YEAR_VAL
    mat[i, 1] < COLLECTION_RATE_YEAR_VAL, 
    # grow from existing levels to a Target level following an S-shaped curve.
    # Perform the following calculations encounter below if the statement above is true
    COLLECTION_RATE_CURRENT_VAL + (COLLECTION_RATE_TARGET_VAL - COLLECTION_RATE_CURRENT_VAL) / 
      (1 + exp(-(10 / COLLECTION_RATE_YEAR_VAL) * (mat[i, 1] - 1 - (COLLECTION_RATE_YEAR_VAL - 1) / 2))),
    # Assign the following value below if the statement is false
    COLLECTION_RATE_TARGET_VAL
  )
  #-----------------------------------------------------------------------------------------
  
  # Assign to the rows from 2 to nrows, and column 16 the following calculations below
  # The cummulative product of row 1 and column 16 times the exponential value raised to the exponent of negative FIXED_OPERATING_COSTS_ANNUAL_REAL_GROWTH_RATE-VAL
  # this exponential will be repeated nrows-1 times, and skip the first index. 
  mat[i, 16]=cumprod(c(mat[1, 16], rep(exp(FIXED_OPERATING_COSTS_ANNUAL_REAL_GROWTH_RATE_VAL), nrows-1)))[-1]
  # Assign to the rows from 2 to nrows, and column 17 the following calculations below
  # The cummulative product of row 1 and column 17 times the exponential value raised to the exponent of negative FIXED_OPERATING_COSTS_ANNUAL_REAL_GROWTH_RATE-VAL
  # this exponential will be repeated nrows-1 times, and skip the first index. 
  mat[i, 17]=cumprod(c(mat[1, 17], rep(exp(FIXED_OPERATING_COSTS_ANNUAL_REAL_GROWTH_RATE_VAL), nrows-1)))[-1]
  # Assign to the rows from 2 to nrows, and column 18 the following calculations below
  # The cummulative product of row 1 and column 18 times the exponential value raised to the exponent of negative VARIABLE_OPERATING_COSTS_ANNUAL_REAL_GROWTH_RATE_VAL
  # this exponential will be repeated nrows-1 times, and skip the first index. 
  mat[i, 18]=cumprod(c(mat[1, 18], rep(exp(VARIABLE_OPERATING_COSTS_ANNUAL_REAL_GROWTH_RATE_VAL), nrows-1)))[-1]
  # Assign to the rows from 2 to nrows, and column 18 the following calculations below
  # The cummulative product of row 1 and column 18 times the exponential value raised to the exponent of negative VARIABLE_OPERATING_COSTS_ANNUAL_REAL_GROWTH_RATE_VAL
  # this exponential will be repeated nrows-1 times, and skip the first index. 
  mat[i, 19]=cumprod(c(mat[1, 19], rep(exp(VARIABLE_OPERATING_COSTS_ANNUAL_REAL_GROWTH_RATE_VAL), nrows-1)))[-1]
  
  #-----------------------------------------------------------------------------------------
  # names(mat)[20] -> Financing costs -> row 37
  # Assign from 2 to nrwos and column 20, the following calculations below
  mat[i, 20] = mat[i, 11] * INVESTMENT_COST_PER_NEW_CONNECTION_LOCAL_CURRENCY_VAL / 
    (mat[i, 4] * DENOMINATIONS_MONETARY_VAL)
  # Assign from 2 to nrows and column 21, the following calculations below
  mat[i, 21] = PERCENTAGE_FUNDED_BY_DEBT_VAL * mat[i, 20]
  #-----------------------------------------------------------------------------------------
  
  # If the value of LOAN_FOREIGN_CURRENCY_VAL is equal to 1
  if(LOAN_FOREIGN_CURRENCY_VAL==1){
    # Assign to d1 the value from the cell row 2 column 5
    d1=mat[2, 5]
    # Assign to dn the values from the cells from row 2 to nrow and column 5
    dn=mat[i, 5]
    # Else if the statement above is false
  }else{
    # Assign to d1 the value of 1
    d1=1
    # Assign to dn the value of d1
    dn=d1
  }
  # Assign to the cell row 1 and column 22, the total sum of 2 to nrows column 21 divided by dn
  mat[1, 22]=sum(mat[i, 21] / dn)
  # Assign from 2 to nrows and column 22, the cumulative sum of 2 to nrows column 21 divided by dn and the result multiplied by LOAN_PERIOD_VAL
  mat[i, 22]=cumsum((mat[i, 21] / dn)/LOAN_PERIOD_VAL)
  # Assign from 2 to nrows and column 23, the values obtained from the funciton Reduce
  mat[i, 23]=Reduce(
    # Creates a function that takes two inputs, and apply the following calculations
    function(prv,nxt) {prv+nxt}, 
    # Takes the vector below and applies the function created above
    # From 2 to nrows column 22 multiply by negative 1 and skip the indexes from -nrows + 2
    # To those values add the values from 2 to nrows from column 21 divided by dn and skipping its first index
    -(mat[i, 22])[-nrows+2]+(mat[i, 21]/dn)[-1], # next
    # Initial value 
    mat[2, 21] / d1, # initial value
    # Accumulate the results
    accumulate = TRUE
  )
  # Assing the values from 2 nrows and column 22 to the rows from 2 to nrows and column 24
  mat[i, 24]=mat[i, 22]
  # Assign the multiplication from 2 nrows and column 23 times INTEREST_RATE_VAL to 2 nrows and column 25
  mat[i, 25]=mat[i, 23] * INTEREST_RATE_VAL
  # names(mat)[22:24]
  
  #-----------------------------------------------------------------------------------------
  
  # Assign to j all the values from 2 to nrows minus 1
  j = i - 1
  
  # Assign all the values from j that are less than or equal to LOAN_PERIOD_GRACE_VAL
  # Then Assign all those values and additionally the value of LOAN_PERIOD_GRACE_VAL + 1
  l = which(j<=LOAN_PERIOD_GRACE_VAL); l=c(l, LOAN_PERIOD_GRACE_VAL+1)
  
  # Assign to all the rows l and column 26 to 31 the value of 8
  mat[l, 26:31]=0
  
  # Assign all the values from j that are greater than LOAN_PERIOD_GRACE VAL
  # Then Add a + 1 to all the numbers on the vector
  l=which(j>LOAN_PERIOD_GRACE_VAL); l=l+1
  
  # Assign only the indexes from the length of the vector l - 1 the values of dn, to the variable TEMP_ARR
  TEMP_ARR=dn[l-1]
  # Assign to the rows l and columns 26,27,28 the calculations done below
  # Take the rows l-LOAN_PERIOD_GRACE_VAL and Columns 26:28-3 and multiply by the variable TEMP_ARR AND FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_OPERATOR_VAL
  mat[l, 26:28]=mat[l - LOAN_PERIOD_GRACE_VAL, 26:28 - 3]*TEMP_ARR*
    FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_OPERATOR_VAL
  # Assign to the rows l and columns 29,30,31 the calculations done below
  # Take the rows l-LOAN_PERIOD_GRACE_VAL and Columns 29:31-6 and multiply by the variable TEMP_ARR AND FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_CONTRACTING_AUTHORITY_VAL
  mat[l, 29:31]=mat[l - LOAN_PERIOD_GRACE_VAL, 29:31 - 6]*TEMP_ARR*
    FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_CONTRACTING_AUTHORITY_VAL
  
  #-----------------------------------------------------------------------------------------
  # If the RISK_TOGGLE is not equal to true
  if (!RISK_TOGGLE){ # With No-Risk
    #-----------------------------------------------------------------------------------------
    # Assign that all the values for column 32 are NA_real
    mat[, 32] = NA_real_
    # Assign that all the values for column 33 are NA_real
    mat[, 33] = NA_real_
    
    # Assign to the cell row 1 and column 37 the result of the division below
    mat[1, 37] = ASSET_BASE_INITIAL_LOCAL_CURRENCY_VAL / DENOMINATIONS_MONETARY_VAL
    #mat[i, 37] = mat[i-1, 37] * exp(mat[i, 2])
    # Assign to the cells from 2 to nrows and column 37 the values obtain from the function Reduce
    mat[i, 37] = Reduce(
      # Create a function reduce that takes 2 inputs and does the following calculation
      function(prv,nxt) {prv * exp(nxt)}, 
      # Apply the function above to the vector the rows from 2 to nrows and column 2
      mat[i, 2], # next
      # Initial Value
      mat[1, 37],
      # Accumulate the results
      accumulate = TRUE
      # Skip first index 
    )[-1]
    
    # Assign to the rows from 2 to nrows and column 34, the values from rows 2 to nrows minus 1 and column 37
    mat[i, 34] = mat[i-1, 37]
    # Assign to the rows from 2 to nrows and column 35, the values from rows 2 to nrows minus 1 and column 37 multiplied by ASSET_BASE_DEPRECIATION_RATE_VAL
    mat[i, 35] = mat[i-1, 37] * ASSET_BASE_DEPRECIATION_RATE_VAL
    
    # Assign to the rows from 2 to nrows and column 39, the values from rows 2 to nrows and column 20
    mat[i, 39] = mat[i, 20]
    
    # Assign to the rows from column 42 the value of NA
    mat[, 42] = NA_real_
    # Assign to the rows from column 43 the value of NA
    mat[, 43] = NA_real_
    
    # Assign to the rows from 2 to nrows column 44, the calculation done below
    mat[i, 44] = mat[i, 13] * mat[i, 7]
    # Assign to the rows from 2 to nrows column 52, the calculation done below
    mat[i, 52] = (mat[i, 16] + mat[i, 18] * mat[i, 13] / (1 - mat[i, 14]))
    # Assign to the rows from 2 to nrows column 53, the calculation done below
    mat[i, 53] = (mat[i, 17] + mat[i, 19] * mat[i, 13] / (1 - mat[i, 14])) * mat[i, 5]
    
    # Assign to the rows from column 54 the value of NA
    mat[, 54] = NA_real_
    # Assign to the rows from column 55 the value of NA
    mat[, 55] = NA_real_
    
    # Assign to the variable TEMP_ARR, the values from the rows from 2 to nrows column 4
    TEMP_ARR = mat[i, 4]
    
    #-----------------------------------------------------------------------------------------
    # Else with risk
  } else { # with risk
    #-----------------------------------------------------------------------------------------
    
    # Assign the value of the cell row 1 column 2, to the cell of row 1 column 32
    mat[1, 32] = mat[1, 2]
    # Assign to the cell row 2 column 32, the calculations found below
    mat[2, 32] = mat[1, 32]*exp(INFLATION_SPEED_ADJUSTMENT_VAL*(INFLATION_LONG_TERM_VAL-mat[1, 32]))
    # Assign to the rows from 2:nrows skipping first index and column 32
    mat[k, 32] = Reduce(
      # Creates a function that takes two inputs
      function(prv,nxt) {
        # If the values are true from the indexes of the second input minus 1 
        if (TFLAG1_VAL[nxt-1]){
          # Take nxt rows and 2 column
          mat[nxt, 2]
          # If the values are false from the indexes of the second input minus 1
        } else {
          # Applyt the following calculation found below from the original inputs and inputs from this function
          prv*exp(INFLATION_SPEED_ADJUSTMENT_VAL*(INFLATION_LONG_TERM_VAL-prv))
        }
      }, 
      # Apply the function above to the vector found below that goes from 2:nrows skipping first index 
      k, # next
      # Initial Value
      mat[2, 32], 
      # Accumulate the results
      accumulate = TRUE
      # Skip first index
    )[-1]
    # Assign to the row 1 column 37 the following calculations
    mat[1, 37] = ASSET_BASE_INITIAL_LOCAL_CURRENCY_VAL / DENOMINATIONS_MONETARY_VAL
    # Assign to the row 2 column 37 the following calculations
    mat[2, 37] = mat[1, 37] * exp(mat[2, 32])
    # Assign to the rows from 2:nrows skipping first index and column 37
    mat[k, 37] = Reduce(
      # Creates a function that takes two inputs
      function(prv,nxt) {
        # If TFLAG2_VAL the indexes from nxt-1 is TRUE
        if (TFLAG2_VAL[nxt-1]){
          # Do the following calculation below
          prv*exp(mat[nxt, 2])
          # If TFLAG2_VAL the indexes from nxt-1 is FALSE
        } else {
          # Do the following calculation below
          prv*exp(mat[nxt, 32])
        }
      }, 
      # Apply the function above to the vector found below that goes from 2:nrows skipping first index  
      k, # next
      # Initial Value
      mat[2, 37],
      # Accumulate the results
      accumulate = TRUE
      # Skip the first index
    )[-1]
    
    # Assign 0 to the cell 1st row column 34
    mat[1, 34] = 0
    # Assign the value of row 1 column 37, to the row 2 column 34
    mat[2, 34] = mat[1, 37] 
    # Assign to the rows from 2:nrows skipping first index column 34, the values from rows from 2:nrows skipping first index and subtracting 1 
    # from each value and column 37
    mat[k, 34] = mat[k - 1, 37]
    
    # Assign to 0 to the cell row 1 column 35
    mat[1, 35] = 0 * ASSET_BASE_DEPRECIATION_RATE_VAL
    # Assign to the cell row 2 column 35, the following calculation below
    mat[2, 35] = mat[1, 37] * ASSET_BASE_DEPRECIATION_RATE_VAL
    # Assign to the rows from from 2:nrows skipping first index column 35, the values from rows from 2:nrows skipping first index and subtracting 1
    # from each value column 37 and multiply it by ASSET_BASE_DEPRECIATION_RATE_VAL
    mat[k, 35] = mat[k - 1, 37] * ASSET_BASE_DEPRECIATION_RATE_VAL
    
    # Total demand, expressed in millions of cubic meters is: Demand * 365 days / 1000 / 1000000
    # Assign the cumulative product of the following values found below that embodies
    # The first cell of row 1 column 12, and then e raised to the exponent of DEMAND_ANNUAL_GROWTH_VAL 
    # That value will be repeated nrows - 1 times
    DDEMAND_VAL=cumprod(c(mat[1, 12], rep(exp(DEMAND_ANNUAL_GROWTH_VAL), nrows-1)))
    # Assign to the TDEMAND_VAl the following calculation
    TDEMAND_VAL=DDEMAND_VAL * mat[, 10] * 365 / 1000 / DENOMINATIONS_WATER_VOLUMES_VAL
    
    # Assign 0 to the cell row 1 column 42
    mat[1, 42] = 0 # TDEMAND_VAL
    # Assign the second index from TDEMAND_VAL to the cell row 2 column 42
    mat[2, 42] = TDEMAND_VAL[2]
    # Assign to the rows from 2:nrows skipping first index and column 42
    mat[k, 42] = Reduce(
      # Create a function that takes two inputs
      function(prv,nxt) {
        # If the value from the index of the second input from TDEMAND_VAL is not equal to 0
        if (TDEMAND_VAL[nxt]!=0){
          # Apply the following calculations from the original inputs and the inputs from the function
          prv / mat[nxt-1, 10] * exp(DEMAND_ANNUAL_GROWTH_VAL) * mat[nxt, 10]
          # If the value from the index of the second input from TDEMAND_VAL is equal to 0
        } else {
          # Assign the value of the first input
          prv
        }
      }, 
      # Apply the function above to the vector found below that goes from 2:nrows skipping first index  
      k, # next
      # Initial Value
      mat[2, 42], 
      # Accumulate the results
      accumulate = TRUE
      # Skip the first index
    )[-1]
    
    # Assign the value from the cell row 1 column 42, to the cell of row 1 column 43
    mat[1, 43] = mat[1, 42]
    # Assign the value from the cell row 2 column 42, to the cell of row 2 column 43
    mat[2, 43] = mat[2, 42]
    # Assign to the rows from 2:nrows skipping first index and column 43
    mat[k, 43] = Reduce(
      # Creates a function that takes two inputs
      function(prv,nxt) {
        # If TFLAG2_VAL the indexes from nxt-1 is TRUE
        if (TFLAG2_VAL[nxt-1]){
          # Assign the values below
          mat[nxt, 13]
          # If TFLAG2_VAL the indexes from nxt-1 is False, then
        } else {
          # Do the following calculations below
          prv / mat[nxt-1, 10] * exp(DEMAND_ANNUAL_GROWTH_VAL) * mat[nxt, 10]
        }
      }, 
      # Apply the function above to the vector found below that goes from 2:nrows skipping first index  
      k, # next
      # Initial Value
      mat[2, 43], # initial value
      # Accumulate the results
      accumulate = TRUE
      # Skip the first index
    )[-1]
    
    # Assign 0 to the row 1 column 54
    mat[1, 54] = 0 # EXCHANGE_RATE_CURRENT_VAL * exp(EXCHANGE_RATE_DEPRECIATION_VAL)
    # Assign the following calculations to row 2 column 54
    mat[2, 54] = EXCHANGE_RATE_CURRENT_VAL * exp(EXCHANGE_RATE_DEPRECIATION_VAL)
    # Assign the cumulative product of the following values found below that embodies
    # The first cell of row 2 column 54, and then e raised to the exponent of EXCHANGE_RATE_DEPRECIATION_VAL 
    # That value will be repeated nrows - 2 times
    # Finally result skip the first index
    mat[k, 54] = cumprod(c(mat[2, 54], rep(exp(EXCHANGE_RATE_DEPRECIATION_VAL), nrows-2)))[-1]
    
    # Assign the value of row 1 column 54, to row 1 column 55
    mat[1, 55] = mat[1, 54]
    # Assign the value of row 2 column 54, to row 2 column 55
    mat[2, 55] = mat[2, 54]
    
    # Assign to the rows from 2:nrows skipping first index and column 55
    mat[k, 55] = Reduce(
      # Create a function that takes two inputs
      function(prv,nxt) {
        # If TFLAG2_VAL the indexes from nxt-1 is TRUE
        if (TFLAG2_VAL[nxt-1]){
          # Assign the following value
          mat[nxt, 5]
          # If TFLAG2_VAL the indexes from nxt-1 is FALSE
        } else {
          # Do the following calculations found below
          prv*exp(EXCHANGE_RATE_DEPRECIATION_VAL)
        }
      }, 
      # Apply the function above to the vector found below that goes from 2:nrows skipping first index  
      k, # next
      # Initial Value
      mat[2, 55], # initial value
      # Accumulate the results
      accumulate = TRUE
      # Skip the first index
    )[-1]
    
    #-----------------------------------------------------------------------------
    
    # Assign 1 to the cell row 1 column 33
    mat[1, 33] = 1
    # Assign from 2 to nrows and column 33, the values obtain from the Reduce function
    mat[i, 33] = Reduce(
      # Create a function that takes two inputs and performs the following calculations
      function(prv,nxt) {prv * exp(-nxt)}, 
      # Apply the function to the vector below
      mat[i, 32], # next
      # Initial Value
      mat[1, 33], # initial value
      # Accumulate the results
      accumulate = TRUE
      # Skip the first index
    )[-1]
    
    # Assign to the cells from 2 to nrows column 39, the following calculations below
    mat[i, 39] = mat[i, 11] * INVESTMENT_COST_PER_NEW_CONNECTION_LOCAL_CURRENCY_VAL / 
      (mat[i, 33] * DENOMINATIONS_MONETARY_VAL)
    
    # Assign to the cells from 2 to nrows column 44, the following calculations
    mat[i, 44] = mat[i, 43] * mat[i, 7]
    
    # Assign to the column 52, the value of NA
    mat[, 52] = NA_real_
    # Assign to the column 53, the value of NA
    mat[, 53] = NA_real_
    
    # Assign to the variable TEMP_ARR, the values from 2 to nrows column 33
    TEMP_ARR = mat[i, 33]
    
    #-----------------------------------------------------------------------------------------
  }
  #-----------------------------------------------------------------------------------------
  
  # Assign to TEMP1_SUM the total sum of 2 to nrows column 44
  TEMP1_SUM = sum(mat[2:nrows, 44])
  # Assign to the cells from 2 to nrow column 36, the following calculations
  mat[i, 36] = mat[i, 37] - mat[i, 34] + mat[i, 35]
  
  # mat[i, 38] = mat[i-1, 41]
  # mat[i, 40] = (mat[i, 38] + mat[i, 39]) * ASSET_BASE_DEPRECIATION_RATE_VAL
  # mat[i, 41] = mat[i, 38] + mat[i, 39] - mat[i, 40]
  
  # Assign to the cells from rows 2 to nrows, column 41
  mat[i, 41] = Reduce(
    # Create a function that takes two inputs, and applies the following calculation
    function(prv,nxt) {prv + mat[nxt, 39] - (prv + mat[nxt, 39]) * ASSET_BASE_DEPRECIATION_RATE_VAL}, 
    # Apply the following function to the vector below
    i, # next
    # Initial Value
    mat[1, 41], # initial value
    # Accumulate the results
    accumulate = TRUE
    # Skip the first index
  )[-1]
  
  # Assign to the cells from rows 2 to nrows, column 38. The values from rows 2 to nrows subtractted by 1, column 41
  mat[i, 38] = mat[i-1, 41]
  # Assign to the cells from rows 2 to nrows, column 40, the following calculations
  mat[i, 40] = (mat[i, 38] + mat[i, 39]) * ASSET_BASE_DEPRECIATION_RATE_VAL
  
  # Assign to the cells from rows 2 to nrows, column 47, the following calculations
  mat[i, 47] = DISCOUNT_RATE_VAL * (mat[i, 38] + mat[i, 39]) * 
    FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_CONTRACTING_AUTHORITY_VAL
  
  # Assign to the cells from rows 2 to nrows, column 48, the following calculations
  mat[i, 48] = DISCOUNT_RATE_VAL * (mat[i, 38] + mat[i, 39]) * 
    FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_OPERATOR_VAL
  
  # Assign to the cells from rows 2 to nrows, column 49, the following calculations
  mat[i, 49] = mat[i, 40] * FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_CONTRACTING_AUTHORITY_VAL * TEMP_ARR
  # Assign to the cells from rows 2 to nrows, column 50, the following calculations
  mat[i, 50] = mat[i, 40] * FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_OPERATOR_VAL * TEMP_ARR
  
  # Assign to the cells from rows 2 to nrows, column 51, the following calculations
  mat[i, 51] = mat[i, 36] * TEMP_ARR
  # Assign to the cells from rows 2 to nrows, column 56, the following variable obtained from the original input
  mat[i, 56] = ANNUAL_FIXED_FEES_PAYMENT_LOCAL_CURRENCY_VAL
  
  
  # Revenue requirement (real): Annual revenue requirement for the two parties. The contracting
  # authority collects the return and depreciation on existing assets and new assets. The
  # operator collects all other cost items.
  # Assign to the cells from rows 2 to nrows, column 40, the following calculations
  mat[i, 57] = mat[i, 47] + mat[i, 49] - mat[i, 56]
  #mat[i, 60] = mat[i, 48] + mat[i, 50] + mat[i, 51]
  
  #-----------------------------------------------------------------------------------------
  # If RISK_TOGGLE is FALSE
  if (!RISK_TOGGLE){ # 'With No-Risk
    #-----------------------------------------------------------------------------------------
    # Assign to the cells from rows 2 to nrows, column 60, the following calculations
    mat[i, 60] = mat[i, 48] + mat[i, 50] + mat[i, 51] + mat[i, 52] + mat[i, 53]
    
    #-----------------------------------------------------------------------------------------
    # If RISK_TOGGLE is TRUE
  } else { # with risk
    #-----------------------------------------------------------------------------------------
    # Assign to the cells from rows 2 to nrows, column 60, the following calculations
    mat[i, 60] = mat[i, 48] + mat[i, 50] + mat[i, 51] + mat[i, 16] + mat[i, 18] * 
      mat[i, 43] / (1 - mat[i, 14]) + mat[i, 17] * mat[i, 55] + 
      mat[i, 19] * (mat[i, 43] / (1 - mat[i, 14])) * mat[i, 55]
    #-----------------------------------------------------------------------------------------
  }  
  #-----------------------------------------------------------------------------------------
  # Assign to the cells from rows 2 to nrows, column 58, the following calculations
  mat[i, 58] = mat[i, 57] * mat[i, 7]
  # Assign to the variable TEMP2_SUM, the total sum of the rows from 2 to nrows, column 58
  TEMP2_SUM = sum(mat[2:nrows, 58])
  
  # Assign to the cells from rows 2 to nrows, column 61, the following calculations
  mat[i, 61] = mat[i, 60] * mat[i, 7]
  # Assign to the variable TEMP3_SUM, the total sum of the rows from 2 to nrows, column 61
  TEMP3_SUM = sum(mat[2:nrows, 61])
  
  #--------------
  
  # Assign to the cells from rows 2 to nrow, column 45 the results from the function Reduce
  mat[i, 45] = Reduce(
    # Create a function that takes two inptus
    function(prv,nxt) {
      # If TFLAG1_VAL the indexes from nxt-1 is TRUE
      if (TFLAG1_VAL[nxt-1]){
        # Assign the following values
        mat[nxt, 44]
        # If TFLAG1_VAL the indexes from nxt-1 is FALSE
      } else {
        # Assign the following values
        prv+mat[nxt, 44]
      }
    }, 
    # Apply the function above to the following vector that goes from 2 to nrows
    i, # next
    # Initial Value
    mat[1, 45], # initial value
    # Accumulate the results
    accumulate = TRUE
    # Skip the first index
  )[-1]
  
  # Assign to the cells from rows 2 to nrow, column 59 the results from the function Reduce
  mat[i, 59] = Reduce(
    # Creates a function that has two inputs
    function(prv,nxt) {
      # If TFLAG1_VAL the indexes from nxt-1 is TRUE
      if (TFLAG1_VAL[nxt-1]){
        # Assign the following values
        mat[nxt, 58]
        # If TFLAG1_VAL the indexes from nxt-1 is FALSE
      } else {
        # Assignt the following values
        prv+mat[nxt, 58]
      }
    }, 
    # Apply the function above to the following vector that goes from 2 to nrows
    i, # next
    # Initial Value
    mat[1, 59], # initial value
    # Accumulate the results
    accumulate = TRUE
    # Skip the first index 
  )[-1]
  
  # Assign to the cells from rows 2 to nrow, column 59 the results from the function Reduce
  mat[i, 62] = Reduce(
    # Create a function that has two inputs
    function(prv,nxt) {
      # If TFLAG1_VAL the indexes from nxt-1 is TRUE
      if (TFLAG1_VAL[nxt-1]){
        # Assign the following values
        mat[nxt, 61]
        # If TFLAG1_VAL the indexes from nxt-1 is FALSE
      } else {
        # Assign the following values
        prv+mat[nxt, 61]
      }
    }, 
    # Apply the function above to the following vector that goes from 2 to nrows
    i, # next
    # Initial Value
    mat[1, 62], # initial value
    # Accumulate the results
    accumulate = TRUE
    # Skip the first index 
  )[-1]
  
  #-----------------------------------------------------------------------------------------
  #-----------------------------------------------------------------------------------------
  
  # Change the values for i, j, and k 
  # Assign to i the values that goes from 1 to nrows
  # Assign to j the values of i and subtract 1
  # Assign to k the values of i and skip the first index
  i = 1:nrows; j = i - 1; k = i[-1]
  
  # If TARIFF_REVIEW_PERIOD_VAL is equal to 1 or the module between j and TARIFF_REVIEW_PERIOD_VAL is equal to 1. Return output TRUE or FALSE
  TFLAG1_VAL = ((TARIFF_REVIEW_PERIOD_VAL == 1) | (j %% TARIFF_REVIEW_PERIOD_VAL== 1))
  # If the module between j and TARIFF_REVIEW_PERIOD_VAL is equal to 0. Return output TRUE or FALSE
  TFLAG2_VAL = (j %% TARIFF_REVIEW_PERIOD_VAL == 0)
  
  # Assign to l, as an integer the following calculations
  l = as.integer(CONTRACT_LENGTH_VAL / TARIFF_REVIEW_PERIOD_VAL) * TARIFF_REVIEW_PERIOD_VAL
  
  # 46 -> Demand is discounted at the beginning of each reset period for the following
  # reset period. This is used as the denominator in the tariff reset calculations.
  
  # 63 & 64 -> Discounted revenue requirement: Revenue requirement is discounted at the
  # beginning of each reset period for the following reset period. This is used as the
  # numerator in the tariff reset calculations.
  
  #k -> j
  # Assign the values to the rows from 1 to nrows, column 46
  mat[i, 46] = ifelse(
    # If j is less than l
    j < l, 
    # If the initial statement above is true assign the result of the following
    # Inner Ifelse. If TFLAG2_VAL2 is TRUE, Assign the first value, if TFLAG2_VAL2 is FALSE assign 0
    ifelse(TFLAG2_VAL, mat[i + TARIFF_REVIEW_PERIOD_VAL, 45], 0), 
    # If the initial statement above is false assign the values below. 
    # Inner Ifelse. If j is equal to 0 is TRUE, Assign the first value. If j is not equal to 1 assign 0
    ifelse(j == l, mat[CONTRACT_LENGTH_VAL + 1, 45], 0)
  )
  
  # Assign the values to the rows from 1 to nrows, column 63
  mat[i, 63] = ifelse(
    # If j is less than l
    j < l, 
    # If the initial statement above is true assign the result of the following
    # Inner Ifelse. If TFLAG2_VAL2 is TRUE, Assign the first value, if TFLAG2_VAL2 is FALSE assign 0
    ifelse(TFLAG2_VAL, mat[i + TARIFF_REVIEW_PERIOD_VAL, 59], 0), 
    # If the initial statement above is false assign the values below. 
    # Inner Ifelse. If j is equal to 0 is TRUE, Assign the first value. If j is not equal to 1 assign 0
    ifelse(j == l, mat[CONTRACT_LENGTH_VAL + 1, 59], 0)
  )
  
  # Assign the values to the rows from 1 to nrows, column 64
  mat[i, 64] = ifelse(
    # If j is less than l
    j < l, 
    # If the initial statement above is true assign the result of the following
    # Inner Ifelse. If TFLAG2_VAL2 is TRUE, Assign the first value, if TFLAG2_VAL2 is FALSE assign 0
    ifelse(TFLAG2_VAL, mat[i + TARIFF_REVIEW_PERIOD_VAL, 62], 0), 
    # If the initial statement above is false assign the values below. 
    # Inner Ifelse. If j is equal to 0 is TRUE, Assign the first value. If j is not equal to 1 assign 0
    ifelse(j == l, mat[CONTRACT_LENGTH_VAL + 1, 62], 0)
  )
  
  #-----------------------------------------------------------------------------------------
  # Assign to cells from rows 1 to nrows, column 69 the following calculations
  mat[i, 69]=1 / mat[i, 4] # ifelse(mat[i, 4] != 0, 1 / mat[i, 4], Inf)
  
  # 65 to 68 -> Revenue required per m3: Revenue required per cubic meter is the
  # discounted revenue requirement divided by discounted demand.
  
  # Assign to cells from rows 1 to nrows, column 65 the following calculations
  mat[i, 65] = TEMP2_SUM / TEMP1_SUM
  # Assign to cells from rows 1 to nrows, column 66 the following calculations
  mat[i, 66] = TEMP3_SUM / TEMP1_SUM
  
  # Assign to cells from rows 1 to nrows, column 67 the following calculations
  mat[i, 67] = Reduce(
    # Create a function that takes two inputs
    function(prv,nxt) {
      # If TFLAG1_VAL in the index of nxt is TRUE
      if (TFLAG1_VAL[nxt]){
        # Do the following calculations with the following variables
        mat[nxt-1, 63] / mat[nxt-1, 46]
        # If TFLAG1_VAL in the index of nxt is FALSE
      } else {
        # Assign the following values
        prv
      }
    },
    # Apply the function to the vector below that goes from 1 to nrows
    i, # next
    # Initial Value
    0, # initial value
    # Accumulate the values
    accumulate = TRUE
  )[-1]
  # Assign to cells from rows 1 to nrows, column 68 the following calculations
  mat[i, 68] = Reduce(
    # Create a function that takes two inputs
    function(prv,nxt) {
      # If TFLAG1_VAL in the index of nxt is TRUE
      if (TFLAG1_VAL[nxt]){
        # Do the following calculations with the following variables
        mat[nxt-1, 64] / mat[nxt-1, 46]
        # If TFLAG1_VAL in the index of nxt is FALSE
      } else {
        # Assign the values
        prv
      }
    },
    # Apply the function to the vector below that goes from 1 to nrows
    i, # next
    # Initial Value
    0, # initial value
    # Accumulate True
    accumulate = TRUE
    # Skip index
  )[-1]
  
  #-----------------------------------------------------------------------------------------
  #-----------------------------------------------------------------------------------------
  
  # Assign to the cells from rows 1 to nrows and column 70 skipping the first index, the following calculations
  mat[k, 70] = mat[k, 65] * mat[k, 69]
  # Assign to the cells from rows 1 to nrows and column 71 skipping the first index, the following calculations
  mat[k, 71] = mat[k, 66] * mat[k, 69]
  # Assign to the cells from rows 1 to nrows and column 72 skipping the first index, the following calculations
  mat[k, 72] = mat[k, 67] * mat[k, 69]
  # Assign to the cells from rows 1 to nrows and column 73 skipping the first index, the following calculations
  mat[k, 73] = mat[k, 68] * mat[k, 69]
  
  # Assign to the cells from row 1 column 94, the following calculations
  mat[1, 94] = EXISTING_TARIFF_CONNECTED_VAL * mat[1, 92]
  # Assign to the cells from rows 1 to nrows and column 94 skipping the first index, the following calculations
  mat[k, 94] = (mat[k, 72] + mat[k, 73]) * mat[k, 92]
  
  # 74 to 75: Total revenue collected is Tariff (in nominal terms) * Billed demand * Collection rate.
  # Assign to the cells from rows 1 to nrows and column 74 skipping the first index, the following calculations
  mat[k, 74] = (mat[k, 70] + mat[k, 71]) * mat[k, 13] * mat[k, 15]
  # Assign to the cells from rows 1 to nrows and column 75 skipping the first index, the following calculations
  mat[k, 75] = (mat[k, 72] + mat[k, 73]) * mat[k, 13] * mat[k, 15]
  
  
  # 76 to 81: Cost items included in the revenue requirement.
  #-----------------------------------------------------------------------------------------
  # If RISK_TOGGLE is FALSE
  if (!RISK_TOGGLE){ # With No-Risk
    #-----------------------------------------------------------------------------------------
    # Assign to the cells from rows 1 to nrows and column 76 skipping the first index, the following calculations
    mat[k, 76] = (mat[k, 52] + mat[k, 53]) * mat[k, 69]
    #-----------------------------------------------------------------------------------------
    # If RISK_TOGGLE is TRUE
  } else { # With Risk
    #-----------------------------------------------------------------------------------------
    # Assign to the cells from rows 1 to nrows and column 76 skipping the first index, the following calculations
    mat[k, 76] =  (mat[k, 16] + mat[k, 18] * mat[k, 13] / 
                     (1 - mat[k, 14]) + (mat[k, 17] + mat[k, 19] * mat[k, 13] / 
                                           (1 - mat[k, 14])) * mat[k, 5]) * mat[k, 69]
    #-----------------------------------------------------------------------------------------
  }
  #-----------------------------------------------------------------------------------------
  # Assign to the cells from rows 1 to nrows and column 77 skipping the first index, the following calculations
  mat[k, 77] = mat[k, 51] * mat[k, 69]
  # Assign to the cells from rows 1 to nrows and column 78 skipping the first index, the following calculations
  mat[k, 78] = (mat[k, 47] + mat[k, 48]) * mat[k, 69]
  # Assign to the cells from rows 1 to nrows and column 79 skipping the first index, the following calculations
  mat[k, 79] = (mat[k, 49] + mat[k, 50]) * mat[k, 69]
  # Assign to the cells from rows 1 to nrows and column 80 skipping the first index, the following calculations
  mat[k, 80] = -mat[k, 56] * mat[k, 69]
  
  # Assign to the cells from rows 1 to nrows and column 81 skipping the first index
  # As vector the sum of the rows from 1 to nrows from the columns 76-80
  mat[k, 81] = as.vector(rowSums(mat[k, 76:80]))
  
  # 82 to 83: Check that the sum of NPV profit in real terms is equal to zero.
  # Assign to the cells from rows 1 to nrows and column 82 skipping the first index, the following calculations
  mat[k, 82] = (mat[k, 74] - mat[k, 81]) * mat[k, 4] * mat[k, 7]
  # Assign to the cells from rows 1 to nrows and column 83 skipping the first index, the following calculations
  mat[k, 83] = (mat[k, 75] - mat[k, 81]) * mat[k, 4] * mat[k, 7]
  
  
  # 84 to 87 -> Cash Flows: Present value of cash flows to the operator and the contracting
  # authority, expressed in millions of pesos.
  # Assign to TEMP_VAL the following calculations
  TEMP_VAL = mat[k, 7] * mat[k, 4]
  
  # Assign to the cells from rows 1 to nrows and column 84 skipping the first index, the following calculations
  mat[k, 84] = TEMP_VAL * (mat[k, 70] * mat[k, 13] * mat[k, 15] - 
                             mat[k, 30] - mat[k, 31] - (mat[k, 20] - mat[k, 21]) * 
                             FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_CONTRACTING_AUTHORITY_VAL)
  
  # Assign to the cells from rows 1 to nrows and column 86 skipping the first index, the following calculations
  mat[k, 86] = TEMP_VAL * (mat[k, 72] * mat[k, 13] * mat[k, 15] - mat[k, 30] - 
                             mat[k, 31] - (mat[k, 20] - mat[k, 21]) * 
                             FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_CONTRACTING_AUTHORITY_VAL)
  
  # Assign to the cells from rows 1 to nrows and column 85 skipping the first index, the following calculations
  mat[k, 85] = TEMP_VAL * (mat[k, 71] * mat[k, 13] * mat[k, 15] - 
                             (mat[k, 76] + mat[k, 77] + mat[k, 27] + mat[k, 28] + 
                                (mat[k, 20] - mat[k, 21]) * FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_OPERATOR_VAL))
  
  # Assign to the cells from rows 1 to nrows and column 87 skipping the first index, the following calculations
  mat[k, 87] = TEMP_VAL * (mat[k, 73] * mat[k, 13] * mat[k, 15] - 
                             (mat[k, 76] + mat[k, 77] + mat[k, 27] + mat[k, 28] + 
                                (mat[k, 20] - mat[k, 21]) * FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_OPERATOR_VAL))
  
  # 88 to 91 -> Debt-service ratio is the total cash flow of the operator divided by
  # financing costs. If the operator has no responsibility in the financing of coverage
  # extension, this ratio is undefined.
  # Assign to the cells from rows 1 to nrows and column 88 skipping the first index, the following calculations
  mat[k, 88] = mat[k, 71] * mat[k, 13] * mat[k, 15] - (mat[k, 76] + mat[k, 77] + mat[k, 27] + 
                                                         (mat[k, 20] - mat[k, 21]) * FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_OPERATOR_VAL)
  
  # Assign to the cells from rows 1 to nrows and column 89 skipping the first index, the following calculations
  mat[k, 89] = mat[k, 73] * mat[k, 13] * mat[k, 15] - (mat[k, 76] + mat[k, 77] + mat[k, 27] + 
                                                         (mat[k, 20] - mat[k, 21]) * FINANCING_RESPONSIBILITY_COVERAGE_EXTENSION_OPERATOR_VAL)
  
  # Assign to the row 1, column 90 the value of NaN
  mat[1, 90] = NaN
  # Assign to the cells from rows 1 to nrows and column 90 skipping the first index, the following calculations
  mat[k, 90] = mat[k, 88] / mat[k, 28]
  # Assign to the cells from rows 1 to nrows and column 90 skipping the first index, the following calculations
  # If the Values of the rows from 1 to nrows column 90 is infinite assign NaN
  # If the Values of the rows from 1 to nrows column 90 is not infinite assign the values
  mat[k, 90] = ifelse(is.infinite(mat[k, 90]), NaN, mat[k, 90])
  
  # Assign to the row 1, column 91 the value of NaN
  mat[1, 91] = NaN
  # Assign to the cells from rows 1 to nrows and column 91 skipping the first index, the following calculations
  mat[k, 91] = mat[k, 89] / mat[k, 28]
  # Assign to the cells from rows 1 to nrows and column 91 skipping the first index,
  # If the Values of the rows from 1 to nrows column 91 is infinite assign NaN
  # If the Values of the rows from 1 to nrows column 91 is not infinite assign the values
  mat[k, 91] = ifelse(is.infinite(mat[k, 91]), NaN, mat[k, 91])
  
  #-----------------------------------------------------------------------------------------
  #-----------------------------------------------------------------------------------------
  # Assign to the cells from rows 1 to nrows and column 95, the following calculations
  mat[i, 95] = EXISTING_TARIFF_CONNECTED_VAL * mat[i, 69] * mat[i, 92]
  # Assign to the cells from rows 1 to nrows and column 96, the following calculations
  mat[i, 96] = EXISTING_TARIFF_OTHER_VAL * mat[i, 93] / mat[i, 4]
  # Assign to the cells from rows 1 to nrows and column 97, the following calculations
  mat[i, 97] = EXISTING_TARIFF_COPYING_COST_VAL / mat[i, 4]
  
  #Willingness to pay
  # Assign to the cells from rows 1 to nrows and column 98, the following calculations
  mat[i, 98] = WTP_CONNECTION_VAL * EXISTING_TARIFF_CONNECTED_VAL * mat[i, 92] / mat[i, 4]
  # Assign to the cells from rows 1 to nrows and column 99, the following calculations
  mat[i, 99] = WTP_OTHER_VAL * EXISTING_TARIFF_OTHER_VAL * mat[i, 93] / mat[i, 4]
  # Assign to the cells from rows 1 to nrows and column 100, the following calculations
  mat[i, 100] = (WTP_COPYING_COST_VAL * EXISTING_TARIFF_COPYING_COST_VAL) / mat[i, 4]
  
  #-----------------------------------------------------------------------------------------
  #-----------------------------------------------------------------------------------------
  
  # Change in social welfare (Real)
  
  # Assign to the cells from rows 1 to nrows and column 101 skipping the first index, the following calculations
  mat[k, 101] = (mat[k, 95] - mat[k, 94]) * mat[1, 10] * mat[k, 4] * mat[k, 7] /
    DENOMINATIONS_MONETARY_VAL
  # Assign to the cells from rows 1 to nrows and column 102 skipping the first index, the following calculations
  mat[k, 102] = (mat[k, 99] + mat[k, 100] - mat[k, 96] - mat[k, 97] + 
                   mat[k, 98] - mat[k, 94]) * mat[k, 11] * mat[k, 4] * 
    mat[k, 7] / DENOMINATIONS_MONETARY_VAL
  
  # Assign to the cells from rows 1 to nrows and column 102 skipping the first index,
  # The cumulative sum from rows 1 to nrows and column 102 skipping the first index  
  mat[k, 102] = cumsum(mat[k, 102])
  #-----------------------------------------------------------------------------------------
  #-----------------------------------------------------------------------------------------
  # Return the matrix
  return(mat)
}


