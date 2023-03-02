Option Explicit
Sub TickerYearlychangepercentchangeandTotalstockvolume():

    'Defining the variables
    Dim ws As Worksheet
    'Variable for opening price at first row for a particular year of a particular stock
    Dim openpricefr As Double
    'Variable for closing price at last row for a particular year of a particular stock
    Dim closepricelr As Double
    Dim inputrow As Long
    Dim summaryrow As Long
    Dim lastRow As Long
    Dim Ticker As String
    Dim Yearlychange As Double
    Dim percentchange As Double
    Dim volume As LongLong
    Dim greatesttotalvolume As Variant
    
    'Activating eachworksheet
    For Each ws In Worksheets
        ws.Activate
        'Selecting the Last row
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        summaryrow = 2
        ' Add the word Ticker, yearly change, percentage change and total stock volume
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change($)"
        Cells(1, 11).Value = "Percentage Change"
        Cells(1, 12).Value = "Total Stock Volume"
        'Adding the words for Greater % Calculations
        'Add the words Ticker, Value, Greatest total volume, Greater % increase, Greater % decrease
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
        Cells(2, 14).Value = "Greater % Increase"
        Cells(3, 14).Value = "Greater % Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
        
        ' Autofit to display data
        Columns("A:P").AutoFit
        
        'Starting the for loop for Ticker, Yearly change, Percentage change and Total stock volume
        For inputrow = 2 To lastRow
            Ticker = Range("A" & inputrow).Value
            'Grabbed the close price last row of particular year
            If Ticker <> Range("A" & (inputrow + 1)).Value Then
                'Input
                closepricelr = Range("F" & inputrow).Value
                'Calculation the Yearly change and percentage change
                Yearlychange = closepricelr - openpricefr
                percentchange = (Yearlychange / openpricefr) * 100
                ' Calculating the total stock volume if the condition meets
                volume = volume + Range("G" & inputrow).Value
                'Output
                Cells(summaryrow, 9) = Ticker
                Cells(summaryrow, 10).Value = Yearlychange
                Cells(summaryrow, 11).Value = percentchange
                Cells(summaryrow, 12).Value = volume
                'prepare for the next stock
                summaryrow = summaryrow + 1
                volume = 0
            ' Grabbed the first price first row of a particular year
            ElseIf Ticker <> Range("A" & (inputrow - 1)).Value Then
                openpricefr = Range("C" & inputrow).Value
            ' Calculating the total stock volume if the condition doesn't meet
            Else
                volume = volume + Range("G" & inputrow).Value
            End If
        Next inputrow
        
        'Conditional formatting loop work for column J (Percentage Change)
        'Selecting the Last row as the row count would be smaller for this one so column J has been inserted at column A
        lastRow = Cells(Rows.Count, 10).End(xlUp).Row
        ' Add the conditional formatting to column J and Column K
        For inputrow = 2 To lastRow
                'Conditional statement of red color for negative values)
                If Cells(inputrow, 10) < 0 Then
                    Cells(inputrow, 10).Interior.Color = vbRed
                'Conditional statement of green color for positive values
                Else
                    Cells(inputrow, 10).Interior.Color = vbGreen
                End If
        Next inputrow
        
        'Greater percent Calculations
        For inputrow = 2 To lastRow
            'Input
            greatesttotalvolume = Range("L" & inputrow).Value
            percentchange = Range("K" & inputrow).Value
            'Get the maximum value in column total stock volume
                If greatesttotalvolume = WorksheetFunction.Max(Range("L2:L" & lastRow)) Then
                    'Output
                    Cells(4, 15).Value = Cells(inputrow, 9).Value
                    Cells(4, 16).Value = greatesttotalvolume
                    'Get the maximum value in column Percentage Change
                ElseIf percentchange = WorksheetFunction.Max(Range("K2:K" & lastRow)) Then
                    'Output
                    Cells(2, 15).Value = Cells(inputrow, 9).Value
                    Cells(2, 16).Value = percentchange
                    'Get the minimum value in column Percentage Change
                ElseIf percentchange = WorksheetFunction.Min(Range("K2:K" & lastRow)) Then
                    'Output
                    Cells(3, 15).Value = Cells(inputrow, 9).Value
                    Cells(3, 16).Value = percentchange
                End If
            Next inputrow
    
    Next ws
    
    MsgBox ("All done")
        
End Sub

