# stock-analysis

Link to the Excel-file: https://github.com/Mishauwa/stock-analysis/blob/main/VBA_Challenge.xlsm

## Overview of Project: Explain the purpose  of this analysis

The purpose of the project is to edit/ refactor the solution code of a stock analysis.
In the stock analysis code stocks of the year 2017 and 2018 gets analyzed and it should be determined in which stocks it's worth investing. 
With refactoring the code we want to achieve that the code is more efficient and runs faster. No new functionalities will be added. 

## Results:

Below you can find pictures of the stock performances in the year 2017 and 2018. 
In 2017 the stocks performed in general very well. There are even 4 stocks (DQ, ENPH, FSLR, SEDG) which have 3 digits growths. Only one stock (TERP) made a loss. 

<img width="240" alt="Stocks_VBA_2017" src="https://user-images.githubusercontent.com/69826498/188000534-bc1086bb-82da-4517-8871-93eb5e4f5b11.png">

In 2018 the picture looks completely different. Only 2 stocks (ENPH and RUN) kept on growing. All the other stocks made losses. 
As a recommendation i would suggest to look at these 2 stocks first to anlayze in depth and consider an investment. For the other stocks I would check why the stocks made losses in 2018. It's important to know whether the negative trend is only temporary or will be permanent. E.g. it was perhaps only a general hype in the sector in 2017 and the companies are not really profitable or in future profitable. 

<img width="240" alt="Stocks_VBA_2018" src="https://user-images.githubusercontent.com/69826498/188001369-ab0e08b5-9d11-43f1-bfc2-c8ca3d14d01a.png">

The code in the original analysis is running in 0.38 seconds for the years 2017 and 2018. 
That means the refactored code is with 0.0625 for 2017 and 0.085 seconds for 2018 is between 4 and 6 times faster. The times vary a little bit by running the code multiple times. 

<img width="145" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/69826498/188002855-85be798c-8fe2-421a-8e57-dc7f9cbce213.PNG">

<img width="154" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/69826498/188002904-6fc583ca-2d66-4aa6-badb-34d255da4a6d.PNG">


## Summary

### What are the advantages and disadvantages of refactoring code?

Refactoring code helps to have a more organized and cleaner code. 
It's easier to read and better to understand. Also it's better to maintain the code. In addition the run time of the code is faster and can help especially running big analysis. 

It's risky to edit/ refactor the code in big applications. Especially if the original code is not well documented and hard to understand. 
It can be very time comsuming to refactor code and also it seems like there is a big risk by rewriting someone's other's code to get lost in the coding process.  

### How does these pros and cons apply to refactoring the original VBA script?

The biggest advantage of the refactored code is the decreased macro run time. 
In general the code is at the end also more organized and easier to read. By comparing the codes you can see that instead of having 2 nested for loops in the original code we have now 3 for loops after each other. With using arrays and indexes the code looks more clean and is faster.

A disadvantage is that it takes some time to figure out how the arrays are working with the indexes and it felt harder to implement it than the original script. 

'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet

    Worksheets(yearValue).Activate

    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    '1a) Create a ticker Index
    
    Dim tickerIndex As Integer
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.

    For i = 2 To RowCount
    
    '3a) Increase volume for current ticker
     
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value


    '3b) Check if the current row is the first row with the selected tickerIndex.

        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    
        End If
    
    
    '3c) check if the current row is the last row with the selected ticker
    'If the next row’s ticker doesn’t match, increase the tickerIndex.


        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then

        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If

    '3d Increase the tickerIndex.
    
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
        tickerIndex = tickerIndex + 1

        End If

    Next i


    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.

    For i = 0 To 11
    
    Worksheets("AllStocksAnalysis").Activate
    
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    

    Next i

