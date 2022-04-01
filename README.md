# Stock Analysis with VBA

Analyzing the performance of Green Energy stocks in 2017 and 2018 using VBA in Excel

## Overview

Some eager investors are excited about investing in green energy. Can we analyze the best performing stocks and show our findings in a way that's quick and easy to understand?

## Results

Green Energy stocks were wildly successful in 2017! Across the board, stocks soared, with some stocks gaining almost 200%!

![Stock Analysis 2017](/Resources/Stock_Analysis_2017.PNG)

2018, however, was a different story, with many losing 50% of their value over the year.

![Stock Analysis 2018](/Resources/Stock_Analysis_2018.png)

In short, green energy can be a lucrative field to invest in, but it is extremely volatile. Ideally, a diverse portfolio would limit risk, while still including industries important to the investor.

### Code

These tables were created by analyzing tons of code, and this was only looking at one industry! Initially, the code I made was clunky, and looking at the time it took to run, it would've been a poor choice to use to expand to a larger data set.

#### Original Method

Originally, the code was written to run through all the data for each individual ticker. This was done via a nested for loop:

```
'Loop through the tickers.
For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
'Loop through rows in the data.
        Worksheets(yearValue).Activate
        
            For j = rowStart To rowEnd
```

It was effective, but inefficient. Running through thousands of tickers 12 times slowed the code down. This was the time it took to run the code through the 2017 data:

![Original Code Time](/Resources/Original_Code_Timer.png)

That took almost a second to run through the green energy industry! If this was to be expanded to larger data sets, this would be unwieldy. Is it possible to get all the data without running through all the code several times?

#### Refactored Code

To slim down the code, arrays needed to be used.

```
'Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```

Now, using an index, the code can run through the data once, increasing the index when the ticker changes.

```
For i = 2 To RowCount
    
        'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        'Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        End If
        
        'check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            'Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
```

By running through the data only once, the run time of the code is way faster:

![Run time 2017](/Resources/VBA_Challenge_2017.png)

![Run time 2018](/Resources/VBA_Challenge_2018.png)

With speedier run times, this code is well positioned for bigger data sets!

## Summary

The refactored code runs significantly faster, however, it took some short cuts to get there. The biggest one is that the code runs on the assumption that the tickers are grouped and in order. If either of those assumptions are wrong, the whole code fails to work.

In general, refactoring is great for double checking the quality and efficiency of code, though it requires well organized code so that multiple coders can understand and update existing code. 

