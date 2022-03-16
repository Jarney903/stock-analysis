# Green Stock Analysis with VBA

## Overview of Project

####  The purpose of this project is to utilize VBA to retrieve and compare environmentally friendly stocks based on total volume and percent return performance in 2017 and 2018 respectively. The client, Steve, would like to use this analysis to help guide decisions in diversifying his parent’s environmentally friendly stock portfolio.  
<br />

## Results

#### The analysis was performed originally by creating for loops with nested if-then statements in VBA to loop through stock data to organize and produce the total volume and percent return performance in 2017 and 2018 respectively. This data was then published to an easy to consume workbook with conditional formatting to show the stock’s annual performance for these years. 

#### To improve the efficiency of the VBA code, thus improving the run-time of the macro, an indexing technique was applied to utilize index values to variable arrays. As seen in Figure 1, the index variable was used to trigger the stock ticker, totalize the trading volume for the year in question, and determine start of year and end of year closing price. 
<br />

### Figure 1: Index Technique Examples

```
    '1a) Create a ticker Index
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For j = 0 To 11
    tickerVolumes(tickerIndex) = 0
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    	For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value =    tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            End If
    
    Next i
    
   Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
  
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
```
#### The results were conclusive that by using the refactored code utilizing indexing techniques, the code runtime was improved by 1/10th of a second for 2017 and 2018 data as seen in Figure 2 and Figure 3 respectively. 
<br />

### Figure 1: Runtime for 2017 Data
![2017]()
<br />

### Figure 2: Runtime for 2018 Data
![2018]()


## Summary

- The advantages and disadvantages of refactoring code in general?
#### The general advantages of refactored code are that it allows code to be more efficient, easily read, adaptive to alterations, and generally cleaner and more organized. 
<br />

#### The general disadvantages are that it takes time to refactor code. This time can correlate to more money spent on a code that at the end of the refactoring process, will provide no additional functionality.  
<br />

## The advantages and disadvantages of the original and refactored VBA script?
#### The advantages of the refactored code, when compared to the original VBA script are as follows:

* ####    Cleaner and organized
* ####    More efficient run time (improved on average 1/10th of a second)
 <br />

## The advantages and disadvantages of the original and refactored VBA script?
#### The disadvantages of the refactored code, when compared to the original VBA script are as follows:

* #### Using the index for filtering through the stock ticker is dependent on the tickers being in the same order as the tickers() array. If the raw data ticker’s order did not match the tickers() array order, a simple tickerIndex = tickerIndex + 1 (as seen on ‘3d of the attached code) would not correctly corelate to the proper ticker. If the number of stocks were such that a manually ordered list of tickers would be impractical to produce, the original method would be advantageous. 
* #### The refactored method added several hours to the original coding time to complete. 

