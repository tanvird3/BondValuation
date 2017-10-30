library(openxlsx)
library(jrvFinance)
library(dplyr)
library(lubridate)
library(ycinterextra)
library(RQuantLib)
bond <-
  read.xlsx("bond.xlsx", detectDates = T) # The Excel File containing the data
setd <- as.Date("2016/12/31") # The valuation date
lvd <- as.Date("2016/12/22") # The Last Valuation Date
matd <- bond$Maturity # Maturity Date of the bond
facev <- bond$Face.Value # Face values of the bonds
coupon <- bond$Coupon # Coupon rates of the bonds

freq <- 2 # frequency of coupon payment
rate <-
  c(4.54 / 100, 6 / 100, 6.89 / 100, 7.64 / 100, 7.89 / 100) # Rate of 2, 5, 10, 15, 20 year bond

# Fraction Year to maturity
fytm <- c()
for (i in 1:length(matd)) {
  fytm[i] <-
    round(RQuantLib::yearFraction(setd, matd[i], 9), 2) # 9 is a daycounter type which is ActualActual.Bond
}


fytm11 <-
  round(as.numeric(((matd - setd) / 365.25)), 2) # Fraction year remaining to maturity

fytm1 <-
  round(daycount.actual(setd[1], matd) / (1 / 2 * daycount.actual(
    as.Date("2016/01/01"), as.Date("2017/12/31")
  )), 2)

# Current Market Yield is calculated based on the fraction year remaining to maturity of the bond
# if fytm is less than 2 (minimum bond maturity), then we have to interpolate cmy on 2 and 5 years
# otherwise the closest two difference are found and cmy is calculated based on it
x <- c(2, 5, 10, 15, 20)
cmy <- c()
for (i in 1:nrow(bond)) {
  if (fytm[i] <= 2) {
    # if the maturity is less than or equal to 2 years then cmy is interpolated on 2 and 5 years
    w = 1
    w1 = 2
  } else {
    # else the two nearest differences of maturity from fy remaining to maturity is calculated and cmy is calculated
    x <- c(2, 5, 10, 15, 20)
    re <- abs(x - fytm[i])
    w <- which(re == min(re))
    w1 <- which(re == min(re[-w]))
    wok <- c(w, w1)
    w <- min(wok)
    w1 <- wok[-which(wok == min(wok))]
  }
  cmy[i] <-
    round((rate[w] + (rate[w1] - rate[w]) / (x[w1] - x[w]) * (fytm[i] - x[w])), 6) #Current Market Yield
}

b <-
  round(round(
    bond.prices(
      rep(setd, nrow(bond)),
      matd,
      coupon = coupon,
      freq = 2,
      yield = cmy,
      convention = "ACT/ACT"
    ),
    4
  ) * facev / 100, 2) # Bond Price; divided by hundred as the value would be for each 100 Taka, so to find value in Taka div by 100
crvd <-
  as.numeric(facev * coupon * (setd - lvd) / 365) # coupon receivable since the last valuation date
di <-
  ceiling(as.numeric((matd - setd) / 180)) # diffeence between maturity date and last settlement date in terms of half year
# and rounded up to ceiling so that the date is before settlementdate
coupdays <-
  as.numeric(setd - (matd %m-% months(6 * di))) # no of days for which coupon has been accrued; with 'months(6 * di)'
# each value of 'di' is converted into 6 months term then its deducted from matd
# and finally from setd
accrued_coupon <- coupdays * facev * coupon / 365 # Accrued Coupon

# lastcouponpaymentdate-----lastvaluationdate-----settlementdate-------maturitydate
# here 6 months wise is deducted from maturity date to find the last couponpayment date and then it is deducted from settlementdate
# to find the no of days for which coupon has been accrued

bond$'Fraction Year to Maturity' <- fytm
bond$'Current Market Yield' <- cmy
bond$'Clean Price' <- b
bond$'Profit/Loss' <- bond[, 15] - bond[, 12]
bond$CouponR <- crvd
bond$CouponDays <- coupdays
bond$Accrued_Coupon <- accrued_coupon
bond$Duration <-
  bond.durations(setd, matd, coupon, 2, convention = "ACT/ACT", yield = cmy)
bond$WA.Duration <- bond[, 15] / sum(bond[, 15]) * bond[, 20]
bond$Modified_Duration <-
  bond.durations(
    setd,
    matd,
    coupon,
    2,
    convention = "ACT/ACT",
    yield = cmy,
    modified = T
  )