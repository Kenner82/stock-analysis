# Stock Analysis

## Overview of Project
VBA script was written to analyze 2017 and 2018 stock data for 12 tickers to determine which stocks had the best returns over the course of a year. The original code performed calculations for starting price, ending price, total volume, and returns for each ticker individually, and then printed those results in a worksheet before moving on to the next ticker. Due to the length of time it took to run the code, it was refactored to store the starting price, ending price, and total volume for each ticker in arrays, and then perform the returns calculations and print all of the values at the end of the script. 

## Results

### Execution Time
The original code took approximately 1.39 seconds to run for both the 2017 and 2018 datasets. Formatting was applied to the results after the data had been entered into the worksheet and the timer had ended.

The refactored code performed the formatting while the timer was running, but still managed to run significantly faster. 

<img width="235" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/111674383/191434053-6560cbed-ed33-46e3-95ed-721b23e74074.png">

<img width="226" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/111674383/191434074-05f22a5d-5fea-472f-bc29-13b9c92a392b.png">

### Stock Performances

Conditional formatting was applied to the cells to make the results easier to interpret. Stocks with gains were highlighted in green, and stocks with losses were highlighted in red. This made it clear that 2017 was a better year for stock performances.

<img width="194" alt="Returns_2017" src="https://user-images.githubusercontent.com/111674383/191437984-4259595d-4a88-4259-b272-e7afcfdd5b72.png">       <img width="190" alt="Returns_2018" src="https://user-images.githubusercontent.com/111674383/191438215-b81fda3a-190a-4465-8182-72f4d6e4d243.png">

## Summary

### Pros And Cons

Refactoring allows improvements to be made to code that has already proven itself to be functional and logically sound. Instead of working out a solution to a problem, the author can focus on ways to make the code more flexible (so that it can easily be reused in other situations), easier for future coders to understand, or (as in this case) more efficient with the resources and time it takes to run. 

However, refactoring also involves tinkering with code that has been designed to run in a particular way, which means the odds are good that you will break something along the way. Properly formatted code with plenty of comments helps mitigate this risk, but care must be taken to check that no other applications rely on functions you might intend to change.

### Applications To The Original Script

While the changes in the logic of both the original and refactored code were minor, the structure of the code changed significantly. It was important to ensure that the same variable names as in the original code were declared and then treated as arrays instead of long variables. A designated index variable (tickerIndex) was also necessary to ensure that the same location in each array was being referenced. 

While the script could have been broken, in this instance the benefits of refactoring outweighed the risks, and steps were taken to ensure that the code was as well commented and easy to update for future coders as possible. 
