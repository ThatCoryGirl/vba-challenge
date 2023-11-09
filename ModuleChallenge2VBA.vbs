Sub VBA_Challenge()

'Create a script that loops through all the stocks for one year and outputs the following information
    'The ticker symbol
    'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year
    'The percentage change from the Opening price at the Beginning of a given year to the Closing price at the End of that year.
    'The total stock volume of the stock.
    'The result should match an image given to the class.
    'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
    'The solution should match the image given to the class.
    
    'I am declaring the variables to use later.
    Dim ticker As String
    Dim dailyopen As Double
    Dim dailyclose As Double
    Dim change As Single
    Dim dailydate As String
    Dim i As Long
    Dim j As Integer
    Dim k As Integer
    Dim endnumber As Long
    Dim yearlyopen As Double
    Dim yearlyclose As Double
    Dim yearlychange As Single
    Dim percentchange As Double
    Dim dailyvolume As Double
    Dim yearlyvolume As Double
    'Function to find last row thespreadsheetguru.com
    Dim LastRow As Long
    Dim GreatestPercentIncrease As Single
    Dim GreatestPercentDecrease As Single
    Dim GreatestTotalVolume As Double
    Dim ws As Integer
    
    
    ws = ActiveWorkbook.Worksheets.Count
    For k = 1 To ws
    'Activate = Focus
    Worksheets(k).Activate
    
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    
    
    
    i = 1
    j = 2
    endnumber = LastRow
    yearlyvolume = 0
    GreatestPercentIncrease = 0
    GreatestPercentDecrease = 0
    GreatestTotalVolume = 0
    
    'We're makin' headers now guys
    Cells(i, 9).Value = "Ticker"
    Cells(i, 10).Value = "Yearly Change"
    Cells(i, 11).Value = "Percent Change"
    Cells(i, 12).Value = "Total Stock Volume"
    
    'Greatest headers
    Cells(i, 16).Value = "Ticker"
    Cells(i, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    'Start ticker loop
    For i = 2 To endnumber
    
    
        'fill buckets
        ticker = Cells(i, 1).Value
        dailyopen = Cells(i, 3).Value
        dailyclose = Cells(i, 6).Value
        change = dailyclose - dailyopen
        dailydate = Cells(i, 2).Value
        dailyvolume = Cells(i, 7).Value
        yearlyvolume = yearlyvolume + dailyvolume
        
        'only pertaining to bucket info
        'Advised to use Right Function by Sean Houston
        If Right(dailydate, 4) = "0102" Then
            yearlyopen = dailyopen
        End If
        
        If Right(dailydate, 4) = "1231" Then 'If then Loop starts here
            yearlyclose = dailyclose
            yearlychange = yearlyclose - yearlyopen
            'Used stackoverflow for percentage formatting
            percentchange = yearlychange / yearlyopen
            If percentchange > GreatestPercentIncrease Then
                GreatestPercentIncrease = percentchange
                Cells(2, 16).Value = ticker
                Cells(2, 17).Value = Format(GreatestPercentIncrease, "0.00%")
            End If
            If percentchange < GreatestPercentDecrease Then
                GreatestPercentDecrease = percentchange
                Cells(3, 16).Value = ticker
                Cells(3, 17).Value = Format(GreatestPercentDecrease, "0.00%")
            End If
            If yearlyvolume > GreatestTotalVolume Then
                GreatestTotalVolume = yearlyvolume
                Cells(4, 16).Value = ticker
                Cells(4, 17).Value = GreatestTotalVolume
            End If
                
            
    
            Cells(j, 9).Value = ticker
            Cells(j, 10).Value = yearlychange
            'Used excelhowto.com for number formatting
            Cells(j, 10).NumberFormat = "0.00"
            'Used excel-easy.com/vba to find color formatting
            'Red is 3, Green is 10
            If yearlychange < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
            If yearlychange >= 0 Then
                Cells(j, 10).Interior.ColorIndex = 10
            End If
            Cells(j, 11).Value = Format(percentchange, "0.00%")
            Cells(j, 12).Value = yearlyvolume
    
            j = j + 1
            yearlyvolume = 0
            
        End If 'If then Loop ends here
        
    Next i
    
    'Autofit columns from learn.microsoft.com
    ActiveSheet.Range("I:Q").Columns.AutoFit
    
    Next k
    
    'MsgBox (GreatestPercentIncrease)
    'yearlychange = yearlyclose - yearlyopen
    'used stackoverflow for percentage formatting
    'percentchange = Format(yearlychange / yearlyopen, "0.00%")
    
    'Cells(j, 9).Value = ticker
    'Cells(j, 10).Value = yearlychange
    'Used excelhowto.com for number formatting
    'Cells(j, 10).NumberFormat = "0.00"
    'Used excel-easy.com/vba to find color formatting
    'Red is 3, Green is 10
    'If yearlychange < 0 Then
        'Cells(j, 10).Interior.ColorIndex = 3
    'End If
    'If yearlychange >= 0 Then
        'Cells(j, 10).Interior.ColorIndex = 10
    'End If
    'Cells(j, 11).Value = percentchange
    'Cells(j, 12).Value = yearlyvolume
    
    
    
    'Make the message box say the ticker, open, and close.
    'Yearly change of the ticker AAB from the open of 23.43 to the close of 23.57
    'If dailydate = "20200102" Then
        'MsgBox ("Yearly change of the ticker " & ticker & " on " & dailydate & " from the open of " & dailyopen & " to the close of " & dailyclose & " was " & change)
    
    'Else
        'MsgBox ("This date was " & dailydate & " but its incorrect")
    
    'End If
    'Calculation between the open and the close
    'MsgBox ("Ticker " & ticker & " had a yearly change of " & yearlychange & " which is " & percentchange)
    'MsgBox (LastRow)
    ' Data for this dataset was generated by EdX Boot Camps LLC, and is intended for educational purposes only.
    
    
    
End Sub