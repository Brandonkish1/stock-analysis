# An Analysis of Green Energy Stocks

## Overview
The purpose of this analysis is analyze Green Energy stocks for Steve's parents. Steve's parents are invested heavily in DAQO New Energy (DQ) and Steve is concerned about their diversification. This analysis shows yearly trade volumes and returns, for 12 green energy stocks in 2017 and 2018. 

## Results




## Summary

### Refactoring Code
#### Advantages
Some advantages of refactored code are improved performance. This could be because the code runs faster, is less repetitive, or structured for a large volume of data rather than a hard coded range. 

#### Disadvantages
There is potential for mistakes. You are taking something that worked and changing it. 

#### How it applies to this analysis
In this anlysis the refactoring of the code made it more efficient. In the original code you are looping through the entire data set 12 times, once for each string in the array.

##### Sample of 12 Loops
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
