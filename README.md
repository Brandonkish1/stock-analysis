# An Analysis of Green Energy Stocks

## Overview
The purpose of this analysis is analyze Green Energy stocks for Steve's parents. Steve's parents are invested heavily in DAQO New Energy (DQ) and Steve is concerned about their diversification. This analysis shows yearly trade volumes and returns, for 12 green energy stocks in 2017 and 2018. 

## Results

### Stock Performance
![Stock_Analysis_Results.pn](https://github.com/Brandonkish1/stock-analysis/blob/main/Stock_Analysis_Results.png)

DQ seems to be volatile with large gains, 199%, in 2017 but and losses, 62%, in 2018. Adding other stocks to the portfolio such as ENPH and RUN could decrease volatility because they increased in value year over year.

### Macro Performance

##### Macro Execution Times (Refactored)
![VBA_Challenge_2017.png](https://github.com/Brandonkish1/stock-analysis/blob/main/VBA_Challenge_2017.png)
![VBA_Challenge_2018.png](https://github.com/Brandonkish1/stock-analysis/blob/main/VBA_Challenge_2018.png)

In this analysis the refactoring of the code made it more efficient. In the original code the macro loops through the entire data set 12 times, once for each string in the array.

##### Sample of 12 Loop Macro
```
    'Loop Through Tickers
    For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0
        
       
        'Loop through Rows in Data
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
            
            '5a. Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
            
                totalVolume = totalVolume + Cells(j, 8).Value
                
            End If
            
            '5b. Find the starting price for the current ticker.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1) = ticker Then
            
                startingPrice = Cells(j, 6).Value
                
            End If
            
            
            '5c. Find the ending price for the current ticker.
             If Cells(j + 1, 1).Value <> ticker And Cells(j, 1) = ticker Then
            
                endingPrice = Cells(j, 6).Value
            End If
            
                        
        Next j
        
    'Output results
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    Next i
```




In the refactored code the macro only loops through the data set once. It checks the ticker in each row to see if it changes. When it changes it increments the tickerIndex to start summarizing the next ticker. This allows the macro to run much faster.


##### Sample of 1 Loops Macro
```
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
            ticker = tickers(tickerIndex)
    
            '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
        
            '3b) Check if the current row is the first row with the selected tickerIndex.
                       
            If Cells(i - 1, 1).Value <> ticker And Cells(i, 1) = ticker Then
            
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
            End If
            
                    
            '3c) check if the current row is the last row with the selected ticker
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
             If Cells(i + 1, 1).Value <> ticker And Cells(i, 1) = ticker Then
            
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            
            
            End If
    
        Next i
```



## Summary

### Refactoring Code
#### Advantages
Some advantages of refactored code are improved performance. This could be because the code runs faster, is less repetitive, or structured for a large volume of data rather than a hard coded range. 

#### Disadvantages
There is potential for mistakes. You are taking something that worked and changing it. 

### How it applies to this analysis
