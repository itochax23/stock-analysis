Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'get organized. Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    Worksheets("2018").Activate

    'dynamically set the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    totalVolume = 0
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    'loop all those rows
    For i = 2 To RowCount
    
        'increase totalVolume only if ticker is DQ
        If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
        End If
        'tell it what to do when it finds the first DQ value in column A
        If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then

        startingPrice = Cells(i, 6).Value

        End If
        'tell it what to do when it hits the last one in column A
        If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then

        endingPrice = Cells(i, 6).Value

        End If
    'tell it to keep doing the loop through the DQ values until there aren't any more
    Next i
    'make it readable for humans, and saved somewhere instead of just in a msgbox or something
    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = "2018"
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = endingPrice / startingPrice - 1
    

End Sub

Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer

'Format the output sheet on the "All Stocks Analysis" worksheet
Worksheets("All Stocks Analysis").Activate
    'set to take input from user about which year to analyze, using concatenation of yearValue variable
    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'shamelessly copy organizational code from before. Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize an array of all tickers. tedious but better than doing it repeatedly. make sure they're spelled right...
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
    'Prepare for the analysis of tickers. what data are we looking for? if it doesn't exist give it a new place to live.
    'Initialize variables for the starting and ending prices

    Dim startingPrice As Double
    Dim endingPrice As Double

    'Activate the worksheet with the data to loop through/where to look
    Worksheets(yearValue).Activate

    'Dynamically set the start and end of the rows to loop through. I guess we just memorize this.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'Loop through the 12 tickers in the list above
   For i = 0 To 11

        ticker = tickers(i)
        totalVolume = 0

        'Activate the worksheet with the data to loop through
         Worksheets(yearValue).Activate

        'loop over all the rows of that specific i ticker we're looking at
        For j = 2 To RowCount

            If Cells(j, 1).Value = ticker Then
        'again with the stealing of code we already wrote for the other thing
            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(j, 8).Value

            End If
            'find starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

            startingPrice = Cells(j, 6).Value

            End If
            'find ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

            endingPrice = Cells(j, 6).Value

            End If
        Next j

    'output the data for the current ticker somewhere for humans to read
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1

    Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)


End Sub
'formatting shenanigans
Sub ResultsFormat()
    Range("A3:C3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
    End With
    With Selection.Font
        .Name = "Calibri"
        .Size = 14
    End With
    Selection.Font.Italic = True
    'honestly it was easier to format in the regular GUI and record it as a macro to understand how to write this
    
    'Number Formats
    Range("C4:C15").Select
    Selection.Style = "Percent"
    Range("B4:B15").Select
    Selection.Style = "Currency"
    Range("C4:C15").Select
    Selection.NumberFormat = "0.00%"
    
    'column format - don't copy this in future, it's ugly
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    'conditional formatting loop of the results on the all stocks worksheet
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Greater than make the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Less than, make the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'otherwise clear the cell color (keep it set to nothing)
            Cells(i, 3).Interior.Color = xlNone

        End If
    'go to the next row until you run into the end
    Next i
End Sub

'Clear the worksheet so we can run the macros again
Sub ClearWorksheet()

Cells.Clear
End Sub

Sub yearValueAnalysis()

yearValue = InputBox("What year would you like to run the analysis on?")

End Sub
