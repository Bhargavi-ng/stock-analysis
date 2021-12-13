# stock-analysis


## Overview of Project:

The purpose of this project is to refactor the VBA code to increase the efficiency. The code was originally written to collect certain stock information in the year 2017 and 2018 and determine whether or not the stocks are worth investing.


## Results:

The output of the refactored code matches with the output before the code was refactored for both 2017 and 2018 as seen in the screenshots below. The execution time for the year 2017 went down to 0.140 seconds from 0.738 seconds after the refactor and for year 2018, the execution time went down to 0.754 seconds to 0.222 seconds.

### Screenshots for year 2017:
#### Before Refactor:
![All Stocks Analysis Output for 2017 with Execution time before Refactor}(Resources/VBA_OriginalCode_Runtime_2017.PNG)

#### After Refactor:
![All Stocks Analysis Output for 2018 with Execution time after Refactor](Resources/VBA_Challenge_2017.png)

### Screenshots for year 2018:
#### Before Refactor:
![All Stocks Analysis Output for 2018 with Execution time before Refactor](Resources/VBA_OriginalCode_Runtime_2018.png)

#### After Refactor:
![All Stocks Analysis Output for 2018 with Execution time after Refactor](Resources/VBA_Challenge_2018.png)

### Refactored code snippet:
Below is the code snippet that is refactored with comments. Complete code for the All Stock Analysis can be found here [VBA Script](Resources/VBA_Challenge.vbs)

```
 '1a) Create a ticker Index
    tickerIndex = 0


   '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    
    
   '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        
    Next i
            
            
   '2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount
                      
       '3a) Increase volume for current ticker
        If Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        End If
        
       '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        End If
        
       '3c) check if the current row is the last row with the selected ticker
        If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
        End If
                         
       '3d Increase the tickerIndex.
        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
        
    Next j
            

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
               
    Next i
```

## Summary: 

### 1. What are the advantages or disadvantages of refactoring code?
    #### Advanages:
   - Make the code more efficient and improve performance
   - Improve the logic
   - Make the code easier to read for future users
   - Use less memory

    #### Disadvanages:
   - Additional resources like time and budget that is needed to do the refactor
   - Might introduce new bugs
   - Person doing the refactor might not understand the original code
        
### 2. How do these pros and cons apply to refactoring the original VBA script?
   The Pros of refactoring the original VBA script as part of this challenge improved the performance and made the code easier to read for future users.

   The Cons of refactoring the original VBA script as part of this challenge was waste of resources. The dataset is too small to see the benefit of refactor. But, in future if the dataset changes to something big this code would be beneficial.
