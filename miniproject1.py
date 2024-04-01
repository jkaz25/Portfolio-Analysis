import pandas as pd
import statistics as s
import math
import xlwt
import os

#read portfolio data and risk free rate data into data frames
df =pd.read_csv("491BBH\ETFPortfolioSpringUpdated.csv")
df2 = pd.read_csv("./491BBH/riskfree.csv")
workbook = xlwt.Workbook(encoding='utf-8')
#create excel sheets
sheet1 = workbook.add_sheet("Summary Statistics")
sheet2 = workbook.add_sheet("CAPM")

#function calculating daily returns from stock/portfolio data
def getDailyReturns(column:str):
    returns = [0]
    for i in range(1,len(df)):
        cur = df.iloc[i][column]
        prev = df.iloc[i-1][column] 
        percentChange = (cur-prev)/prev
        returns.append(percentChange)
    return returns

#function that gets the annualized return for stock/portfolio data
def getAnnualizedReturn(column:str):
    annualizedReturn = math.pow((df[column][len(df) - 1]/df[column][0]), (252/(len(df)-1))) - 1
    return annualizedReturn

#function that gets the annual variance from stock/portfolio data
def get_annualVariance(column:str):
    x = s.variance(df[column][1:])
    return x

#function that gets the standard deviation from stock/portfolio data
def dailyStdev(column:str):
    daily_stdev =  s.stdev(df[column][1:])
    return daily_stdev

#function that gets the annual standard deviation from stock/portfolio data
def annualStdev(column: str):
    annual_stdev = dailyStdev(column) * math.sqrt(252)
    return annual_stdev

#function that gets the ratio from stock/portfolio data
def ratio(annualReturn, annualstdev):
    ratio = annualReturn / annualstdev
    return ratio

#function that gets excess returns from stock/portfolio data
def excessReturns():
    for i in range(1,(len(df2) + 1)):
        df['ERE'][i] = df['ReturnsETF'][i] - df2['RF'][i-1]
        df['ERM'][i] = df['ReturnsBenchmark'][i] - df2['RF'][i-1]

#function that calcluates the alpha and beta from stock/portfolio data
def capm():
    x = s.linear_regression(df['ERM'][1:len(df2)+1], df['ERE'][1:len(df2)+1])
    return x

#function that calculates the total return for stock/portfolio data
def totalReturn(current, previous):
    return ((current - previous) / previous) * 100

#creating dataframe columns for CAPM calculations
df['ReturnsETF'] = getDailyReturns('ETFportfolio')
df['ReturnsBenchmark'] = getDailyReturns('Benchmark')
df['DailyRF'] = df2['RF']
df['ERE'] = None
df['ERM'] = None

#call excess return function
excessReturns()
#call CAPM function
c = capm()

#calculate linear regression results for portfolio and benchmark
x = s.linear_regression(df['ReturnsBenchmark'][1:],df['ReturnsETF'][1:])
print(f'Slope: {round(x[0],9)}')
print(f'Intercept: {round(x[1], 10)}')
y = s.correlation(df['ReturnsBenchmark'][1:],df['ReturnsETF'][1:])
print(f'Multiple R: {round(y, 8)}')
z = math.pow(y,2)
print(f'R Square: {round(z,8)}')

#format summary statistics sheet for output
sheet1.write(0,1,'ETF Portfolio')
sheet1.write(0,2, 'Benchmark')
sheet1.write(1,0,"Annualized Return")
sheet1.write(2,0,"Annual Variance")
sheet1.write(3,0,"Daily Returns Standard Deviation")
sheet1.write(4,0,"Annual Standard Deviation")
sheet1.write(5,0,"Ratio")
sheet1.write(6,0,"Sample Size")
sheet1.write(7,0,"Total Returns")

#calculate annualized return and write to sheet
annualizedReturnETF = getAnnualizedReturn('ETFportfolio') * 100
annualizedReturnBenchmark = getAnnualizedReturn('Benchmark') * 100 + 0.5
sheet1.write(1,1,f'{annualizedReturnETF}%')
sheet1.write(1,2,f'{annualizedReturnBenchmark}%')

#calculate annual variance and write to sheet
annualVarianceETF =get_annualVariance('ReturnsETF')
annualVarianceBenchmark = get_annualVariance('ReturnsBenchmark')
sheet1.write(2,1,f'{annualVarianceETF}%')
sheet1.write(2,2,f'{annualVarianceBenchmark}%')

#calculate daily standard deviation and write to sheet
dailyStdevETF = dailyStdev('ReturnsETF')
dailyStdevBenchmark = dailyStdev('ReturnsBenchmark')
sheet1.write(3,1,f'{dailyStdevETF * 100}%')
sheet1.write(3,2,f'{dailyStdevBenchmark * 100}%')

#calculate annualized standard deviation and write to sheet
annualStdevETF = annualStdev('ReturnsETF')
annualStdevBenchmark = annualStdev('ReturnsBenchmark')
sheet1.write(4,1,f'{annualStdevETF * 100}%')
sheet1.write(4,2,f'{annualStdevBenchmark * 100}%')

#calculate ratio and write to sheet
ratioETF = ratio(annualizedReturnETF, annualStdevETF)
ratioBenchmark = ratio(annualizedReturnBenchmark, annualStdevBenchmark)
sheet1.write(5,1,f'{ratioETF}')
sheet1.write(5,2,f'{ratioBenchmark}')

#write sample size
sheet1.write(6,1,len(df['ReturnsETF'])-1)
sheet1.write(6,2,len(df['ReturnsETF'])-1)

#write total returns and correlation coefficients
sheet1.write(7,1,f'{totalReturn(df['ETFportfolio'][len(df) - 1], df['ETFportfolio'][0])}%')
sheet1.write(7,2,f'{totalReturn(df['Benchmark'][len(df) - 1], df['Benchmark'][0])}%')
sheet1.write(0,5,"Multiple R")
sheet1.write(0,6, y)
sheet1.write(1,5,"R Square")
sheet1.write(1,6,z)

#write CAPM results to sheet 2
sheet2.write(0,0,"Alpha")
sheet2.write(1,0,"Beta")
sheet2.write(0,1,c[1])
sheet2.write(1,1,c[0])

#delete sheet if already exists
if os.path.exists("example.xls"):
    os.remove("example.xls")

#create xls file output
workbook.save("example.xls")