Sub marketscan():

' Declare Variables and Baseline Variables
    Dim tickervolume As Double
    Dim i As Long
    Dim ychange As Double
    Dim j As Integer
    Dim lastcount As Long
    Dim lastRow As Double
    Dim percentchange As Double
    Dim dates As Integer
    Dim dchange As Double
    Dim change As Double
    
    
' Column headers
    
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
' Headers for additional table


        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
' Values

        j = 0
        tickervolume = 0
        ychange = 0
        lastcount = 2
        
' Determine last row
        
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
    
' ticker changes, then print results

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    
    tickervolume = tickervolume + Cells(i, 7).Value
    
' 0 total ticker volume

    If tickervolume = 0 Then
    
    Range("I" & 2 + j).Value = Cells(i, 1).Value
    Range("J" & 2 + j).Value = 0
    Range("K" & 2 + j).Value = "%" & 0
    Range("L" & 2 + j).Value = 0
    
    Else
    

'  Find first non zero values in the previous amounts

    If Cells(lastcount, 3) = 0 Then
    For nonvalues = lastcount To i
    If Cells(nonvalues, 3).Value <> 0 Then
    lastcount = nonvalues
    
    Exit For
    
    End If
    
    Next nonvalues
    
    End If
    
' Determine change

   ychange = (Cells(i, 6) - Cells(lastcount, 3))
    percentchange = Round((ychange / Cells(lastcount, 3) * 100), 2)
    
    
' previous amount of stock ticker, then results


    lastcount = i + 1
    
    Range("I" & 2 + j).Value = Cells(i, 1).Value
    Range("J" & 2 + j).Value = Round(ychange, 2)
    Range("K" & 2 + j).Value = "%" & percentchange
    Range("L" & 2 + j).Value = tickervolume
    
    
' Conditional formatting, colors green and red if negative

    Select Case ychange
    
    Case Is > 0
    Range("J" & 2 + j).Interior.ColorIndex = 4
    
    Case Is < 0
    Range("J" & 2 + j).Interior.ColorIndex = 3
    
    Case Else
    Range("J" & 2 + j).Interior.ColorIndex = 0
    
    End Select
    
    End If
    
    
    
' reset ticker , then add results if ticker is the same

    tickervolume = 0
    ychange = 0
    j = j + 1
    dates = 0
    
        Else
        
        tickervolume = tickervolume + Cells(i, 7).Value
        
        End If
        
        Next i
        
        
' max and min values separated

    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastRow)) * 100
    
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastRow)) * 100
    
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & lastRow))
    
    
    GreatestIncrease = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastRow)), Range("K2:K" & lastRow), 0)
    GreatestDecrease = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastRow)), Range("K2:K" & lastRow), 0)
    GreatestVolume = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastRow)), Range("L2:L" & lastRow), 0)
    
    
' ticker symbol for Greatest Increase, Decrease and Volume

    Range("P2") = Cells(GreatestIncrease + 1, 9)
    Range("P3") = Cells(GreatestDecrease + 1, 9)
    Range("P4") = Cells(GreatestVolume + 1, 9)
    
    
End Sub

    
    
    
    
    
    
        
        
        
        
        
