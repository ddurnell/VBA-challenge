Attribute VB_Name = "Module1"
Sub StockLoop()
    Start = Timer
'Create a script that will loop through all the stocks for one year for each run and take the following information.
 ' The ticker symbol.
 ' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
 ' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
 ' The total stock volume of the stock.
 
 ' Do it for every worksheet in the book
    Dim Current As Worksheet
    For Each Current In Worksheets
        'Choose the workbook
        'MsgBox Current.Name
        Worksheets(Current.Name).Activate
        
        Dim totalRows As Long
        totalRows = Cells(Rows.Count, 1).End(xlUp).Row
          
        Dim ticker As String
        Dim startPrice As Currency
        Dim endPrice As Currency
        Dim percent As Double
        Dim totalVolume As Double
        Dim tickerNum As Long
        Dim minPercent As Double
        Dim maxPercent As Double
        Dim maxVolume As Double
        Dim minPercentTicker As String
        Dim maxPercentTicker As String
        Dim maxVolumeTicker As String
        
        ticker = "blank"
        tickerNum = 2
        
        ' Set the headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        ' Loop through unique tickers
        For I = 2 To totalRows
            ' if we have changed the ticker on the next loop
            If ticker = "blank" Then
                ticker = Cells(I, 1)
                startPrice = Cells(I, 3)
                totalVolume = 0
            'new ticker
            ElseIf Cells(I, 1) <> ticker Then
                endPrice = Cells(I - 1, 6)
                Cells(tickerNum, 9).Value = ticker
                Cells(tickerNum, 10).Value = (endPrice - startPrice)
                If startPrice = 0 Then
                    Cells(tickerNum, 11).Value = 100#
                Else
                    percent = (endPrice - startPrice) / startPrice
                    If percent < minPercent Then
                        minPercent = percent
                        minPercentTicker = ticker
                    ElseIf percent > maxPercent Then
                        maxPercent = percent
                        maxPercentTicker = ticker
                    End If
                    
                    Cells(tickerNum, 11).Value = FormatPercent(percent)
                End If
                Cells(tickerNum, 12).Value = totalVolume
                ticker = "blank"
                tickerNum = tickerNum + 1
            ' add up total volume
            Else
                totalVolume = totalVolume + Cells(I, 7)
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                End If
            End If
        Next I
        
        'Summarize Headers
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        'Summarize Values
        Range("P2").Value = maxPercentTicker
        Range("P3").Value = minPercentTicker
        Range("P4").Value = maxVolumeTicker
        Range("Q2").Value = FormatPercent(maxPercent)
        Range("Q3").Value = FormatPercent(minPercent)
        Range("Q4").Value = Str(maxVolume)
        
        Range("=A:Z").Columns.AutoFit
    
    Call Color(totalRows)
    
    Next

    Finish = Timer
    TotalTime = Finish - Start
    MsgBox "That took a whopping " & Round(TotalTime, 1) & " seconds"
End Sub

' Conditionally format positive and negative
' Gleaned from StackOverflow and YouTube examples
Sub Color(totalRows As Long)
    ' (1) Make negative values red
    With Range("J:J").FormatConditions.Add(xlCellValue, xlLess, "=0")
        .Interior.ColorIndex = 3
        .StopIfTrue = False
    End With

    ' (2) Make positive values green
    With Range("J:J").FormatConditions.Add(xlCellValue, xlGreater, "=0")
        .Interior.ColorIndex = 4
        .StopIfTrue = False
    End With
    
    'We don't want the header green
    Range("J1").FormatConditions.Delete
    
End Sub
