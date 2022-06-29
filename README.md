# stocks-analysis

## Analyzing performance of green stocks.

### Overview of Project:

#### The purpose of this project is to determine if a refactored VBA script performs more quickly than the original script. In this case, we are analyzing green stock data. In addition to analyzing the performance of each script, we will also view which green stocks performed well in the year 2017 and 2018. Measuring the performance speed of the original and refactored VBA script can show which is best to use and could be helpful in case a larger amount of stock data needs to be added to the data set and analyzed quickly.

### Results:

#### Based on the execution times, the refactored script is faster than the original script. The original script for 2017 data ran in about 1.035 seconds, whereas the refactored script ran in about 0.180 seconds. For 2018, the original script ran in about 1.063 seconds and the refactored script ran in about 0.242 seconds. 

#### Before the script was refactored:
#### ![BeforeRefactor2017](https://github.com/eoweed/stocks-analysis/blob/main/Resources/BeforeRefactor2017.png) 
#### ![BeforeRefactor2018](https://github.com/eoweed/stocks-analysis/blob/main/Resources/BeforeRefactor2018.png)

#### After the script was refactored:
#### ![Refactored2017](https://github.com/eoweed/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.png)
#### ![Refactored2018](https://github.com/eoweed/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.png)

#### The stock performance in 2017 was better than in 2018. Most stocks in 2017 had a positive return on investment by the end of the year (as shown above in green). In 2018, there were a few stocks that did well, but most had a negative return (as shown above in red). 

### Summary:

#### There are both advantages and disadvantages to refactoring code. One advantage is that refactoring may make the code perform faster by taking fewer steps or using less computer memory. It may also improve the logic of the code and make it easier to understand. However, disadvantages of refactoring code are that it is time-consuming and if it isn’t done well, it may slow down your computer. 

#### In this project, pros of refactoring the VBA script were to use less computer memory and improve the logic. The output arrays for “startingPrices” and “endingPrices” were initialized as single instead of double variables to hold fewer digits, and the “tickerIndex” variable was initialized to reference the tickers array. The “for” loops were also formatted to store the output, and then later create the output arrays at the end. This is quicker than creating the output within each “for” loop. 

#### link: [Refactored VBA Script](https://github.com/eoweed/stocks-analysis/blob/main/VBA_Challenge.vb)

#### The cons to refactoring were that it was extremely time consuming, and even though it did improve the performance speed, it is difficult to determine how useful that is unless we can compare it on a larger data set. 