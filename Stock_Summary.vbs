Option Explicit

Sub StockChange()

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    
    
        'Declare variables
        Dim i As Long
        Dim SumRow As Long
        SumRow = 2
        Dim LastRow As Long
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        'Add headers for summary row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        
        For i = 2 To LastRow
            Dim Volume As Double
            Dim Ticker As String
            Dim Count As Long
            Ticker = Cells(i, 1).Value
            Count = Count + 1
            Volume = Volume + ws.Cells(i, 7).Value
            
            Dim BoYOpen As Double
            Dim EoYClose As Double
            
            Dim YearChange As Double
            Dim PercChange As Double
            
        
            'Get Values when the ticker changes
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                EoYClose = ws.Cells(i, 6).Value
                'calculate values
                    YearChange = EoYClose - BoYOpen
                
                    If BoYOpen = 0 Then
                        PercChange = 0
                    Else
                        PercChange = YearChange / BoYOpen
                    End If
                
                
                'add values and strings to summary row
                ws.Range("I" & SumRow).Value = Ticker
                ws.Range("J" & SumRow).Value = YearChange
                ws.Range("K" & SumRow).Value = PercChange
                ws.Range("L" & SumRow).Value = Volume
                
                
                
                'reset values and move to next summary row
                SumRow = SumRow + 1
                Volume = 0
                Count = 0
                           
            'Get opening amount from first instance of ticker
            ElseIf ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value And Count = 1 Then
                BoYOpen = ws.Cells(i, 3).Value
            
                    
            End If
        
        Next i
            
        
        'Format summary rows
        Dim i2 As Long
        Dim LastSumRow As Long
        
        LastSumRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
        
        For i2 = 2 To LastSumRow
            ws.Range("K2:K" & LastSumRow).NumberFormat = "0.00%"
            
            If ws.Cells(i2, 10) > 0 Then
               ws.Cells(i2, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(i2, 10) < 0 Then
                ws.Cells(i2, 10).Interior.ColorIndex = 3
            Else
            End If
        
        Next i2
        
    Next ws
    
End Sub
