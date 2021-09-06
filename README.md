# Stock-Analysis

## Overview of Project

The Client Steve loved the workbook that we prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute. So, taking all this into consideration & further determining whether refactoring the code successfully makes the VBA script run faster.

### Purpose
The purpose of this project is to determine whether refactoring the code written to analyze stocks for Steve, successfully makes the VBA script run faster. 

## Results & Analysis
Refactoring the code successfully made the VBA script run faster, the below table summarizes the runtimes and the improvements seen for both datasets.  

| Dataset  | Original Code  | Refactored Code  | Improvement Percentage |
|:----------|:----------|:----------|:----------|
| 2017    | 1.2 sec    | 0.14 sec   | 88.33% 
| 2018    | 1.14 sec    | 0.12 sec    | 89.47%


Here are the screenshot showing the runtimes for both the datasets with the original VBA script and the refactored VBA script.   

Original VBA script runtime

![](Resources/original_VBAScript_2017.png)
![](Resources/original_VBAScript_2018.png)

Refactored VBA script runtime

![](Resources/VBA_Challenge_2017.png)
![](Resources/VBA_Challenge_2018.png)


The below screenshots show the results/output of stock analysis, these are also available on a new worksheet called "All Stock Analysis". 

 
![](Resources/Output_stockanalysis_2017.png) 
![](Resources/Output_stockanalysis_2018.png)


The output is presented in a table form with 3 columns named as Ticker, Total Daily Volume & Return. Here are the steps followed to get the output for the stock-analysis:

**Step 1a:**

Create a tickerIndex variable and set it equal to zero before iterating over all the rows. This tickerIndex is used to access the correct index across the four different arrays we'll be using: the tickers array and the three output arrays are created in Step 1b.

**Step 1b:**

Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
The tickerVolumes array should be a Long data type.
The tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.

**Step 2a:**

Create a for loop to initialize the tickerVolumes to zero.

**Step 2b:**

Create a for loop that will loop over all the rows in the spreadsheet.

**Step 3a:**

Inside the for loop in Step 2b, wrote a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker. Use the tickerIndex variable as the index.

**Step 3b:**

Wrote an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices variable.

**Step 3c:**

Wrote an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.

**Step 3d:**

Wrote a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.

**Step 4:**

Used a for loop to loop through the arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to get output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.


Finally, we ran the stock analysis & 


verified the outputs were correct. We also looked at runtimes to make sure the refactored code runs faster. 

## Summary
Every piece of code is written from a perspective of the developer. Based on the perspective of a developer there can be advantage & disadvantages of refactoring a code. The purpose of refactoring a code is to make it run faster, making the code easier to understand & easier to maintain. There are other benefits of refactoring, it changes the way a developer thinks when implementing new code based on his/her leranings from refactoring code.  

There are also disadvantages when it comes to refactoring code, it can be time consuming, no idea to how much time it will take to finish the entire process & sometimes it may land the developer in situation where the developer has no idea what to do next.   

The original VBA script had the advantage that it works correctly.    
The advantage of the refactored VBA script is the refactored VBA script runs faster, than the original VBA script. The other advanatages of refactored VBA script was it's easier to understand & to maintain. 

The disadvantages of original VBA script was that the code was written in way that it read the excel sheet one time for every ticker. 
The disadvantages of refactored VBA script that it too additional time to refactor the code, There was also no idea to how much time it would take to finish the entire process or even if it would yeild any improvements.


Here is a table summarizing the runtime results for the original & refactored code.

| Dataset  | Original Code  | Refactored Code  | Improvement Percentage |
|:----------|:----------|:----------|:----------|
| 2017    | 1.2 sec    | 0.14 sec   | 88.33% 
| 2018    | 1.14 sec    | 0.12 sec    | 89.47%
