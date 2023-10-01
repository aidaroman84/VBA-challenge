Attribute VB_Name = "Module1"
Sub Stocks()
    
    For Each ws In Worksheets
        If ws.Name <> "Summary" Then ' Skip the Summary sheet
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"

            ' Declaring variables
            Dim tickerName As String
            Dim openYearly As Double
            Dim closeYearly As Double
            Dim totalVolume As Double
            Dim yearlyChange As Double
            Dim percentChange As Double
            Dim tickerRow As Long
            tickerRow = 2
            Dim lastRow As Long
            lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

            openYearly = ws.Cells(2, 3).Value

            ' Loop through the data
            For i = 2 To lastRow
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                    tickerName = ws.Cells(i, 1).Value
                    ws.Cells(tickerRow, 9).Value = tickerName
                End If
                
                ' Last day of a year for the current ticker
                If ws.Cells(i, 2).Value Like "*1231" Then
                    closeYearly = ws.Cells(i, 6).Value
                    
                    ' Calculate yearly change
                    yearlyChange = closeYearly - openYearly
                    ws.Cells(tickerRow, 10).Value = yearlyChange
                   
                    ' Calculate percent change
                    If openYearly <> 0 Then
                        percentChange = (yearlyChange / openYearly)
                        ws.Cells(tickerRow, 11).Value = percentChange
                        ws.Cells(tickerRow, 11).Style = "Percent"
                    Else
                        ws.Cells(tickerRow, 11).Value = 0
                    End If
                    
                    ' Reset
                    totalVolume = 0
                    tickerRow = tickerRow + 1
                    openYearly = ws.Cells(i + 1, 3).Value
                End If
                
                ' Calculate total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            Next i
        End If
        
        'Cell formatting
    Dim yearLastRow As Long
    yearLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row


    For i = 2 To yearLastRow
         If ws.Cells(i, 10).Value >= 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
         Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
         End If
    Next i
   
    'Find Max Percent and Min Percent
     Dim percentLastRow As Long
     percentLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
     Dim percent_max As Double
     percent_max = 0
     Dim percent_min As Double
     percent_min = 0

    For i = 2 To percentLastRow
          If percent_max < ws.Cells(i, 11).Value Then
             percent_max = ws.Cells(i, 11).Value
             ws.Cells(2, 17).Value = percent_max
             ws.Cells(2, 17).Style = "Percent"
             ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
          ElseIf percent_min > ws.Cells(i, 11).Value Then
             percent_min = ws.Cells(i, 11).Value
             ws.Cells(3, 17).Value = percent_min
             ws.Cells(3, 17).Style = "Percent"
             ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
          End If
     Next i

    'Greatest total volume
    Dim totalVolumeRow As Long
    totalVolumeRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
    Dim totalVolumeMax As Double
    totalVolumeMax = 0

    For i = 2 To totalVolumeRow
         If totalVolumeMax < ws.Cells(i, 12).Value Then
            totalVolumeMax = ws.Cells(i, 12).Value
            ws.Cells(4, 17).Value = totalVolumeMax
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
         End If
    Next i
   
Next ws

End Sub
    

