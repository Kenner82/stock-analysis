# Stock Analysis

## Overview of Project
VBA script was written to analyze 2017 and 2018 stock data for 12 tickers to determine which stocks had the best returns over the course of a year. The original code performed calculations for starting price, ending price, total volume, and returns for each ticker individually, and then printed those results in a worksheet before moving on to the next ticker. Due to the length of time it took to run the code, it was refactored to store the starting price, ending price, and total volume for each ticker in arrays, and then perform the returns calculations and print all of the values at the end of the script. 

## Results

### Execution Time
The original code took approximately 1.39 seconds to run for both the 2017 and 2018 datasets. Formatting was applied to the results after the data had been entered into the worksheet and the timer had ended.

The refactored code performed the formatting while the timer was running, but still managed to run significantly faster. 

<img width="235" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/111674383/191434053-6560cbed-ed33-46e3-95ed-721b23e74074.png">

<img width="226" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/111674383/191434074-05f22a5d-5fea-472f-bc29-13b9c92a392b.png">
