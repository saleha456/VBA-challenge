Option Explicit

Sub StockChange()
    
    'Declare variables
    Dim i As Long
    Dim SumRow As Long
    SumRow = 2
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Add headers for summary row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    
    For i = 2 To LastRow
        Dim Volume As Double
        Dim Ticker As String
        Dim Count As Long
        Ticker = Cells(i, 1).Value
        Count = Count + 1
        Volume = Volume + Cells(i, 7).Value
        
        Dim BoYOpen As Double
        Dim EoYClose As Double
        
        Dim YearChange As Double
        Dim PercChange As Double
        
    
        'Get Values when the ticker changes
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            EoYClose = Cells(i, 6).Value
            'calculate values
                YearChange = EoYClose - BoYOpen
            
                If BoYOpen = 0 Then
                    PercChange = 0
                Else
                    PercChange = YearChange / BoYOpen
                End If
            
            
            'add values and strings to summary row
            Range("I" & SumRow).Value = Ticker
            Range("J" & SumRow).Value = YearChange
            Range("K" & SumRow).Value = PercChange
            Range("L" & SumRow).Value = Volume
            
            'reset values and move to next summary row
            SumRow = SumRow + 1
            Volume = 0
            Count = 0
                       
        'Get opening amount from first instance of ticker
        ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value And Count = 1 Then
            BoYOpen = Cells(i, 3).Value
        
                
        End If
    
    Next i


End Sub
