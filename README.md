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

### Execution Time - Original Script

### Execution Time - Refactored Script
