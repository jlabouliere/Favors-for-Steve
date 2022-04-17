# Favors-for-Steve
2017 and 2018 Analysis of 12 Stock Tickers
## Project Overview
Steve has elicited help in preparing a stock analyis for his parents with a set of 12 possible tickers for evaluation.  Utilizing VBA, we've attempted to create a functional tool that efficiently runs our script in the background to return a well formatted comparison of the 12 tickers with their Ticker Symbol, Total Daily Volume for the specified year and rate of return as a percentage.  

Steve's parents were particularly interested in Daqo (DQ), and this analysis attempts to show DQ's performance against 11 other options that meet Steve's and his parents' preference for Green Stocks.  After the initial analysis the code was refactored to improve performance and reprove the data set.

## Analysis
### DQ vs The Field
2018 proved to be a difficult year for the prospectus with only two tickers venturing into positive territory after a solid 2017.  DQ, coming off of a 2017 breakout with nearly 200% returns, faltered with losses exceeding 62%.  While DQ was the loss leader in 2018, a $1000 investment in the one stock would have garnered 137% ROI versus 59% returns had the investment been diversified across all 12 tickers.

![image](https://user-images.githubusercontent.com/98665941/163694606-7ae4bc76-cf07-47fe-a69f-8bc1115bda32.png),    ![image](https://user-images.githubusercontent.com/98665941/163694592-7d8838c7-c198-4fd8-953a-3fc890079421.png)

### Code Refactored
Refactoring the code provided a modicum of efficiency with the limited datasets when the macros were run identically.  An average of 82% faster runs after the refactor is indicitave of material improvement should a larger swath of data be analyzed.  One drawback to refactoring, and only in the sense of a lack of experience on my part, was the time input in refactoring the code.

![image](https://user-images.githubusercontent.com/98665941/163696448-69ed3307-0754-4e4a-83d2-2a72215e19ab.png)


![Screen Shot 2022-04-16 at 5 56 14 PM Small](https://user-images.githubusercontent.com/98665941/163696463-f0c6955e-652c-48a4-a10b-9f3d10005421.png),   ![Screen Shot 2022-04-16 at 6 08 03 PM](https://user-images.githubusercontent.com/98665941/163696503-8e5381c2-e9dd-4157-854c-4eda1ae72845.png)

![Screen Shot 2022-04-16 at 5 59 28 PM Small](https://user-images.githubusercontent.com/98665941/163696455-bf823cea-0dba-4190-bbe5-902444cc50cf.png),   ![Screen Shot 2022-04-16 at 6 08 25 PM](https://user-images.githubusercontent.com/98665941/163696511-f942cb25-1961-42e3-b06a-426f0de95565.png)

## Summary
Refactoring the code was an excellent exercise in cementing the learnings by having to repeat the results with new nuances to the syntax.  Removal of the And qualifiers to the tickerStarting/EndingPrices versus the original code provided substantial improvements in the preformance of the script and eliminated what appeared to be extraneous filtering to the data.

![Screen Shot 2022-04-16 at 8 44 54 PM](https://user-images.githubusercontent.com/98665941/163696744-cf851aa7-e4c2-4a24-bad9-f7dd301381e3.png)

As previously stated, given a lack of experience in code, the amount of time undertaken to refactor working code was exponentially higher than the time savings in the run after refactor.  Applying this to a much larger data set and practice on this side will doubtless absorb the inefficiencies.  Since this application is provided as a tool for a third party, refactoring would be recommended.  The code becomes simplified and provides less opportunity for failure.

### Refactored code
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

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
    RowStart = 2
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
        
    Next i
        
        ''2b) Loop over all the rows in the spreadsheet.
        For i = RowStart To RowCount
    
            '3a) Increase volume for current ticker
         
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
        
            '3b) Check if the current row is the first row with the selected tickerIndex.
        
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
            '3c) check if the current row is the last row with the selected ticker
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
            End If
           
    
        Next i
        
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
