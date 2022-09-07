# Analysis of Green Stocks for client Steve
---

## Project Overview

### Stock Analysis Requirement

Steve is a recently graduated stock analyst, upon graduation he is in the process of reviewing his parents investment portfolio. With his parents keen on upcoming green energy stock, they had a special emphasis on investing in the DQ ticker. But, the goal is to diversify the investment portfolio, so this analysis looks at 12 selected greenstocks and assesses the following;

1. The total volume of stocks traded in 2017 and 2018
2. The annual return for each stock in 2017 and 2018

### Methodology - Data Analysis

To analyse the selected ticker data, we use VBA to write code and obtain the desired output. Our initial run of code has seperate loops running through the 2017 and 2018 data sheets, causing inefficiencies in output. But, a goal of this analysis is to refactor the code into a single loop and determine how effectively it runs compared to our code with multiple loops.

This sections below highlight the approach towards refactoring code in VBA for the stock analysis.

## Results

### Stock Performance (2017 & 2018)

The annual return for the tickers under review convey 2 different stories for 2017 and 2018. The annual return for 4 of the 12 stocks for 2017 were more than a 100% return rate. This means your investment at the start of the year would have doubled by the end of the year for those stocks. Furthermore, 11/12 stocks had a positive return rate, which indicates a diversified investment would give you a healthy return.

Although, in 2018 the story is different. Almost 1/3rd the stocks in 2018 has lost more than 40% of thier stock value, whilst a majority of the sotcks (10/12) have shown a negative return for 2018.

Given the high returns for 2017 and decline in 2018, the tickers analyzed show a higher degree of volatility and it is recommended to analyze additional stocks under green energy sector or to look at a longer timeframe of analysis.

The comparison of 2017 and 2018 stock data can be seen below;

![2017_VS_2018](/Other/2017_VS_2018.png)

### Execution Time - Original Script

Cosidering the output from the original script, the VBA script has multiple loops running which causes the code to be slightly inefficient. This results in the execution of the script to take almost a full second each (0.95 secs) to run and provide the complete output for each years output. 

When executing the code in the original script, we only initialize arrays for all the tickers using `Dim tickers(12) As String` and index each variable for all tickers starting from index = 0 to index = 11.

Then we loop over all rows using a `For Loop` and utilize the `If Then` statement for current tickers to obtain the totalVolume, tickerStartingPrices and tickerEndingPrices for current ticker. (Note examples of code used can be seen in attached VBA_Challenge.xlsm file under Module1 in VBA)

The outputs are then assigned to the "All Stocks Analysis" worksheet using the appropriate code and formatting.

The output format and timing can be seen on the image below;

![Original_Script](/Other/Original_Script.png)

### Execution Time - Refactored Script

Once we refactor the initial VBA code into a single loop the script was able to run more efficiently. This can be seen in the run time of the script in the VBA output where both years analysis output time almost cut by 75%, where each script ran in under (0.20 secs).

The output timing of the refactored script can be seen on the image below;

![Refactored_Script](/Other/Refactored_Script.png)

## Summary

### Advantages and Disadvantages of Refactored Code

### Application of Pros and Cons to Original VBA Script
