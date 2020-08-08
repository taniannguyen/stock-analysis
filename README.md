# Green Stocks with VBA Excel

## Overview of Project

### Purpose
The purpose of this analysis is to assist Steve in looking into DAQO stocks for his parents which they have decided to invest all of their money into. Steve is concerned about diversifying their funds and wants to analyze a handful of green energy stocks for alternative stock choices for his parents to choose from.  He wants to find the total daily volume and yearly return for each stock. This will be unraveled by using visual basic for application (VBA) applying on stock performance data between 2017 and 2018, which analyzes the dataset for daily volume and yearly return as well as outputting the execution times of the script. The dataset is a collection of green energy stocks from 2017 and 2018 with different tickers, closing prices, volumes, and etc. The goal will be gathering 12 tickers and display their corresponding total daily volume and percent yearly return with the respective year.  Steve also would like to include stock market over the past few years, this will require to refactor the script and we will compare this to the original script on the execution time. Finding the total daily volume per ticker will give the yearly volume and a rough idea of how often it gets traded. As for the return data this will help determine if a stock is traded often, then the price will accurately reflect the value of the stock. 

Conclusions will be reported based on the analyses that will be created by a table and the difference in execution times based on the original script and refactored script. All stocks (year) will be created from code inputted into VBA on stock performance dataset between 2017 and 2018. The table will represent either 2017 or 2018 with the tickers, total daily volume, and return. The overall results based on the table, DQ would not be a good stock to invest all of Steve’s parents all of their money in because in 2017 it had a return of 199.4% but in 2018 a return of -62.6% which shows poor performance.  Instead Steve should recommend his parents to diversify their investment into ENPH and RUN stocks because ENPH in 2017 had a return of 129.5% and in 2018 a return of 81.9% as for RUN in 2017 had a return of 5.5% and in 2018 a return of 84.0% which indicates a positive trend. The original script execution times for 2017 and 2018 were 0.5625 seconds and 0.5547 seconds respectively. With the refactored script execution times for 2017 and 2018 were 0.1133 seconds and 0.1133  From these results, it will help Steve give recommendations to his parents to where they should invest in particular stocks based on total daily volumes and yearly return, also the refactored script runs the VBA code quicker.


## Results

### Analysis of Stock Performance Between 2017 and 2018


[VBA_Challenge](Module 2 Challenge/VBA_Challenge.vbs)


### Analysis of Execution Times of Original Script and Refactored Script


![VBA_Challenge_2017](Resources/VBA_Challenge_2017.png)


## Summary
- What are the advantages or disadvantages of refactoring code?
  - The advantages of refactoring code are increase efficiency, minimize steps and not as complex, utilizing less memory, easier to maintain and read, or improving     the logic of the code creates an easier experience for future users to read.  Some disadvantages of refactoring code are time consuming, increasing the             execution time on the particular code, or ending up with the wrong results from the refactored code.
- How do these pros and cons apply to refactoring the original VBA script?
	- These points apply to refactoring the original VBA script in a positive aspect because it runs quicker, minimizes steps, increases efficiency, and returns a         larger set of data for the results as requested. The refactored gives opportunities to cover more data and helps support the results that is discovered from the     analysis. The negative aspect based from the pros and cons from refactoring the original VBA script is confusion may arise, time consuming, a lot of trial and       error, or not being able to return the result that the original VBA script asks for.



