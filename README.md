# Stock-Analysis Project

Overview

In this project I refactored the Stock Market Dataet with VBA to loop through all the data. I used the data set available to refactor the code to make the vba script to run faster.  

## Results

I created a 4 different arrays; tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. 
The tickers array was used to establish the ticker symbol of a stock. 

### Original Code 

    2) Initialize array of all tickers

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

    '3a) Initialize variables for starting price and ending price

        Dim startingPrice As Double
        Dim endingPrice As Double

    '3b) Activate data worksheet

        Worksheets(yearValue).Activate

    '3c) Get the number of rows to loop over

        RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    '4) Loop through tickers

        For i = 0 To 11
        ticker = tickers(i)
        TotalVolume = 0
        Worksheets(yearValue).Activate

    '5) loop through rows in the data
            
    For j = 2 To RowCount

        '5a) Get total volume for current ticker

        If Cells(j, 2).Value = ticker Then

            'increase totalVolume by the value in the current row
            TotalVolume = TotalVolume + Cells(j, 9).Value

    End If

            '5b) get starting price for current ticker

        If Cells(j - 1, 2).Value <> ticker And Cells(j, 2).Value = ticker Then
            'set starting price
            startingPrice = Cells(j, 7).Value

        End If

            '5c) get ending price for current ticker
            
            If Cells(j + 1, 2).Value <> ticker And Cells(j, 2).Value = ticker Then
            'set ending price
            endingPrice = Cells(j, 7).Value

        End If

        Next j
    '6) Output data for current ticker

        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = TotalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

     Next i

#### Run-time for 2017 & 2017!

Run-times using the refactored code.

[VBA_Challenge_2017](https://user-images.githubusercontent.com/114794033/196592990-7933dab1-3980-4e30-919e-d1cd15c9e226.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/114794033/196593055-cb1f82fc-5d14-44d4-a5f0-1556df0b8ae8.png)

##### Advanatages and Disadvantages

The major disadvantage of refactoring code in VBA script is that if you do not know syntax, you will struggle to refactor your code. The major advantage of refactoring code in VBA script is that you can use as much as of the original code as you want to. 
The major disadvantage of refactoring code is that you are potentially making it unusable. The major advantage of refactoring code is making the code more efficient. 
