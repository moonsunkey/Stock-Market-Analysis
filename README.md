The spreadsheet you'll be working with in this activity is stock market data for 3 years! You'll use VBA to generate price, volume and performance information based on year and tickers.

## Instructions
Create a script that loops through all the stocks for one year and outputs the following information:

*The ticker symbol
*Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
*The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
*The total stock volume of the stock. The result should match the following image:
*Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
* Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.
*Make sure to use conditional formatting that will highlight positive change in green and negative change in red.


* There are three parts to this problem:

    * **Part 1:** 
  Break the problem down to 3 subroutines and 2 summary tables to hold the results: 
  
  1) Get total stock volume. There are two senarios-same or different ticker between the row and the next row. If the next row ticker is different to the prior row, loop through each row to sum up volume of the same ticker. Need to reset volume at 0 for the second scenario. 
  
  2) Get value of the open price and close price of first and last day of the year for each ticker respectively; perform mathematical functions to calculate price changes and percentage changes; Apply conditionall formatting for positive yearly changes highlight in green and negative in red. Need to use number format function to have a percentage with double decimal points. 

  * **Part 2:** Use for each worksheet loop to run all worksheets at once.
  * **Part 3:** Consolidate all subroutines to run at once. 

*** Notes: Make sure to test on the testing dataset, otherwise it takes a long time to run the codes.
