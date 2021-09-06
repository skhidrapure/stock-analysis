# stock-analysis

## Overview of Project

The Client Steve loved the workbook that we prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute. So taking all this into consideration & futher determining whether refactoring the code successfully made the VBA script run faster.

### Purpose
The purpose of this is project is to determine whether refactoring the code written to analyze stocks for steve, successfully made the VBA script run faster. 

## Results
On taking look at the code that was written, the need for refactoring was confirmed. So, the results of refactoring the code successfully made the VBA script run faster and also the outputs of stock-analysis on a new worksheet called All Stock Analysis. The output ha dto be presented in a table form with 3 columns named as Ticker, Total Daily Volume & Return. Here are the steps followed to get the output for the stock-analysis:

**Step 1a:**

Create a tickerIndex variable and set it equal to zero before iterating over all the rows. You will use this tickerIndex to access the correct index across the four different arrays you’ll be using: the tickers array and the three output arrays you’ll create in Step 1b.

**Step 1b:**

Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
The tickerVolumes array should be a Long data type.
The tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.

**Step 2a:**

Create a for loop to initialize the tickerVolumes to zero.

**Step 2b:**

Create a for loop that will loop over all the rows in the spreadsheet.

**Step 3a:**

Inside the for loop in Step 2b, write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker. Use the tickerIndex variable as the index.

**Step 3b:**

Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices variable.

**Step 3c:**

Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.

**Step 3d:**

Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.

**Step 4:**

Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.

Finally, we ran the stock analysis & checked whether the refactoring the code succefully made the VBA script run faster & it did. So the purpose of the project & the output for Stock-analysis was met.

## Summary
- Every piece of code is written from a prespective of the developer. So, based on the prespective of a different developer there can be advantage & disadvatages of refactoring a code. The purpose of refactoring a code is to make it run faster, making the code easier to understand & easier to maintain. There are other benefits of refactoring, it changes the way a developer thinks about the implentation when not refactoring. There also disadvantages when it comes to refactoring a code like it can be time consuming, no idea to how much time it will take to finish the entire process & sometimes it may land the developer in situation where the developer has no idea what to do next.

- The orignal VBA script had advantages like the user had freedom to give input on which dataset to select, Here is a image of the popup  the output worksheet was activated with the name of headers (Tickers, Total Daily Volume, Return), the vairable tickers defined & assigned with values for each ticker index & formatting for the output data was done. The advantages on the refactored VBA script are after succesfull completion of refactoring the main aim of the project was fulfilled, that is the refactored VBA script ran faster, the script was easier to understand & to maintain. The disadvantages of orginal VBA script is that the code was not fully completed & had errors & the final All Stocks Analysis worksheet though was active did'nt hold the final results whereas the disadvantages of refactored VBA script is it was time consuming, no idea to how much time it will take to finish the entire process & sometimes it may land the developer in situation where the developer has no idea what to do next.

