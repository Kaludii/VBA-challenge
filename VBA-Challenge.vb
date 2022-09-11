Sub Stock_Data()
    
    'Message Box when macro starting
    MsgBox ("Starting!")

    For Each ws In Worksheets
    
        Dim WSName As String
        'Current row in sheet
        Dim i As Long
        'Start row of ticker symbol
        Dim j As Long
        'Counter to fill Ticker symbol row
        Dim TickerTotalCount As Long
        Dim LRowA As Long
        Dim LRowH As Long
        'The variable for % change
        Dim PercentChanged As Double
        'The variable for greatest % increase
        Dim GreatestPercentIncreased As Double
        'The variable for greatest % decrease
        Dim GreatestPercentDecreased As Double
        'Variable for greatest total volume
        Dim GreatestTotalVolume As Double
        
        'Get the Worksheet Name
        WSName = ws.Name
        
        'Automatically create column headers
        Worksheets("2018").Range("H1").Value = "Ticker"
        Worksheets("2018").Range("I1").Value = "Yearly Change"
        Worksheets("2018").Range("J1").Value = "Percent Change"
        Worksheets("2018").Range("K1").Value = "Total Stock Volume"
        Worksheets("2018").Range("O1").Value = "Ticker"
        Worksheets("2018").Range("P1").Value = "Value"
        Worksheets("2018").Range("N2").Value = "Greatest % Increase"
        Worksheets("2018").Range("N3").Value = "Greatest % Decrease"
        Worksheets("2018").Range("N4").Value = "Greatest Total Volume"
        Worksheets("2019").Range("H1").Value = "Ticker"
        Worksheets("2019").Range("I1").Value = "Yearly Change"
        Worksheets("2019").Range("J1").Value = "Percent Change"
        Worksheets("2019").Range("K1").Value = "Total Stock Volume"
        Worksheets("2019").Range("O1").Value = "Ticker"
        Worksheets("2019").Range("P1").Value = "Value"
        Worksheets("2019").Range("N2").Value = "Greatest % Increase"
        Worksheets("2019").Range("N3").Value = "Greatest % Decrease"
        Worksheets("2019").Range("N4").Value = "Greatest Total Volume"          
        Worksheets("2020").Range("H1").Value = "Ticker"
        Worksheets("2020").Range("I1").Value = "Yearly Change"
        Worksheets("2020").Range("J1").Value = "Percent Change"
        Worksheets("2020").Range("K1").Value = "Total Stock Volume"
        Worksheets("2020").Range("O1").Value = "Ticker"
        Worksheets("2020").Range("P1").Value = "Value"
        Worksheets("2020").Range("N2").Value = "Greatest % Increase"
        Worksheets("2020").Range("N3").Value = "Greatest % Decrease"
        Worksheets("2020").Range("N4").Value = "Greatest Total Volume"        
        
        'Set Ticker Counter to first row
        TickCount = 2
        
        'Set start row to 2
        j = 2
        
        'Find the last non-blank cell in column A
        LRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox ("Last row in column A is " & LRowA)
        
            'Loop through all rows
            For i = 2 To LRowA
            
                'Check if ticker name changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Write ticker in column H (#9)
                ws.Cells(TickCount, 8).Value = ws.Cells(i, 1).Value
                
                'Calculate and write Yearly Change in column I (#9)
                ws.Cells(TickCount, 9).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'Conditional formating
                    If ws.Cells(TickCount, 9).Value < 0 Then
                
                    'Set cell background color to red
                    ws.Cells(TickCount, 9).Interior.ColorIndex = 3
                
                    Else
                
                    'Set cell background color to green
                    ws.Cells(TickCount, 9).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculate and write percent change in column K (#11)
                    If ws.Cells(j, 3).Value <> 0 Then
                    PercentChanged = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    'Percent formating
                    ws.Cells(TickCount, 10).Value = Format(PercentChanged, "Percent")
                    
                    Else
                    
                    ws.Cells(TickCount, 10).Value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate and write total volume in column L (#12)
                ws.Cells(TickCount, 11).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase TickCount by 1
                TickCount = TickCount + 1
                
                'Set new start row of the ticker block
                j = i + 1
                
                End If
            
            Next i
            
        'Find last non-blank cell in column H
        LRowH = ws.Cells(Rows.Count, 8).End(xlUp).Row
        'MsgBox ("Last row in column H is " & LRowH)
        
        'Prepare for summary
        GreatestTotalVolume = ws.Cells(2, 11).Value
        GreatestPercentIncreased = ws.Cells(2, 10).Value
        GreatestPercentDecreased = ws.Cells(2, 10).Value
        
            'Loop for summary
            For i = 2 To LRowH
            
                'For greatest total volume--check if next value is larger--if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 11).Value > GreatestTotalVolume Then
                GreatestTotalVolume = ws.Cells(i, 11).Value
                ws.Cells(4, 15).Value = ws.Cells(i, 8).Value
                
                Else
                
                GreatestTotalVolume = GreatestTotalVolume
                
                End If
                
                'For greatest increase--check if next value is larger--if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 10).Value > GreatestPercentIncreased Then
                GreatestPercentIncreased = ws.Cells(i, 10).Value
                ws.Cells(2, 15).Value = ws.Cells(i, 8).Value
                
                Else
                
                GreatestPercentIncreased = GreatestPercentIncreased
                
                End If
                
                'For greatest decrease--check if next value is smaller--if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 11).Value < GreatestPercentDecreased Then
                GreatestPercentDecreased = ws.Cells(i, 10).Value
                ws.Cells(3, 15).Value = ws.Cells(i, 8).Value
                
                Else
                
                GreatestPercentDecreased = GreatestPercentDecreased
                
                End If
                
            'Write summary results in ws.Cells
            ws.Cells(2, 16).Value = Format(GreatestPercentIncreased, "Percent")
            ws.Cells(3, 16).Value = Format(GreatestPercentDecreased, "Percent")
            ws.Cells(4, 16).Value = Format(GreatestTotalVolume, "Scientific")
            
            Next i
            
            
    Next ws
    
    'Message Box when macro ending
    MsgBox ("Done!")
        
End Sub




