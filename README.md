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
  
  1) Get total stock volume. Solved this one first because if its similarity to the credit card activity. Find out the next ticker that's different from the previous one, then add up the total up to the row of the same ticker. Loop through the same process to find total volume for each ticker.
  
  2) The key to yearly change value and percentage is to find values of the opening price of the first day of the year and the closing price of the last day of the year. It takes me a long time to figure out. The previous approach was to try to grab the date and then find price for that date but I had no clue how to do that. Instead, with the help of the tutor, just like the previous part, first date started from row 2 and check if the next row is the same ticker to the previous one. If not, then the last value is at the row where it stopped being the same ticker. Then using loop these values can be found one by one. The mathematical function is the easier portion. But since there is division involved, this scenario was listed out to avoid error. The conditional formatting was the easiest part of the problem since there are only two scenarios with a simple ">0" or "0" logic.
  3) Finding greatest is to find max and min in terms of percentage and max volume. Apply loops to compare the value at each row. I made mistakes in the beginning to mess up the rows and columns count where the required values are supposed to be held.
  * **Part 2:** Use for each worksheet loop to run all worksheets at once.
  * **Part 3:** Consolidate all subroutines to run at once. Also created a button for the macro and a message box.

*** Notes: It was much easier to test on the testing dataset, otherwise it takes a long time to run the codes.
