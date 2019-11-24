# Challenge

# Stock-Analysis

This workbook uses VBA to compile stock data into a short summary that contains:

* Stock Name
* Day of purchase closing price
* Last closing price
* Total trading volume
  
The summary is automatically formatted to color losses red, gains green, and unchanged clear.

# Script Summary

1. Create arrays for ticker, volume, start price, end price

2. Set current ticker

3. Loop to every row
   * If ticker in current row equals current ticker then ->
      * Add volume to ticker's volume array
    
   * If row is first row for current ticker then ->
      * Determine start price
      * Save start price to ticker's start price array
      
    * If row is last row for current ticker then ->
       * Determine end price
       * Save end price to ticker's end price array
       * Increment ticker by 1 (begin work on next ticker)
       * Return volume to 0

4. Push results to worksheet

5. Format results
