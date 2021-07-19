# VBA of Wall Street

## Overview of Project
Our client had a need to:
* analyze stock data for DQ over time
* analyze multiple stock as potential alternatives
* analyze stock data for multiple years
* allow him to input a year to review into the macro if he wanted to compare the performance of multiple stocks
* add buttons and visual cues like conditional formatting to improve readability

This work was done in Excel using VBA macros.

## Results
Using images and examples of the code, we can compare the stock performance between 2017 and 2018, as well as the time it took to run the original script and the refactored script.

### Original code example
When the script was first created, it went through an array of 12 specific stock tickers, not all of them. It looped through each ticker, then through each row to determine the total volume, starting price, and ending price, and would output it with formatting to a separate sheet.

```
'Loop through the 12 tickers in the list above
   For i = 0 To 11

        ticker = tickers(i)
        totalVolume = 0

        'Activate the worksheet with the data to loop through
         Worksheets(yearValue).Activate

        'loop over all the rows of that specific i ticker we're looking at
        For j = 2 To RowCount

            If Cells(j, 1).Value = ticker Then
            'reuse code and increase totalVolume by the value in the current row
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
```

### Refactored code example
```
here
```

## Summary
This summary will address the advantages and disadvantages of refactoring code and how these pros and cons applied to refactoring the original script written for this project.

### Advantanges and disadvantages of refactoring code
Some advantages of refactoring are that it can improve the readability of the code for Future You and other people to update going forward. It may also expand the usability of the script, and allow it to run in a faster amount of time than before.

A disadvantage of refactoring may be that touching the code can result in errors unintentionally; additionally, refactoring could add time to the development of both re-writing the code and fixing new bugs. 

### How these apply to refactoring the original VBA script
We created a VBA macro to trigger pop-ups and get inputs from the user about which data to run the macro on, read and change cell values, and format cells for ease of use.

It was built to originally use for loops and conditionals. This eventually included using nested for loops as well. With refactoring, we were able to use our existing code as a framework for an enhanced script that was faster and allowed for analysis of a much larger dataset. 

This was potentially more time consuming than was originally needed by the client, and there was potential for a more difficult degree of maintainability going forward if someone who is less familiar with nested for loops and the variables we initialized needs to make changes. However, overall it has good comments and structure that is logical, so that in addition to being faster and able to handle larger datasets, it should be more than sufficient in the future.
