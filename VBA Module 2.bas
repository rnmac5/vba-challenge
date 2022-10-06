Attribute VB_Name = "Module1"
Sub Multi_Yr_Stck()

For Each ws In Worksheets
    Dim WorksheetLoop As String
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Dimension variables
Dim total As Double
    total = 0
Dim Ticker As String
Dim TableRow As Integer
    TableRow = 2
Dim YearlyChange As Double
    YearlyChange = 0
Dim OpenVal As Double
    OpenVal = ws.Range("C2").Value
Dim CloseVal As Double
    CloseVal = 0
Dim TotalVol As Double
    TotalVol = 0

    

'Create title row
ws.Range("K1").Value = "Ticker"
ws.Range("L1").Value = "Yearly Change"
ws.Range("M1").Value = "Percentage Change"
ws.Range("N1").Value = "Total Stock Volume"


'Start loop
For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
     
        'Ticker name
        Ticker = ws.Cells(i, 1).Value
        ws.Range("K" & TableRow).Value = Ticker
     
     
     'Yearly Change
     CloseVal = ws.Cells(i, 6).Value
      YearlyChange = CloseVal - OpenVal
      ws.Range("L" & TableRow).Value = YearlyChange
      
            'Color Index
            If YearlyChange > 0 Then
                ws.Range("L" & TableRow).Interior.ColorIndex = 4
          
                Else
                ws.Range("L" & TableRow).Interior.ColorIndex = 3
        
            End If
    
    'Percent Change
    PercentChange = CloseVal - OpenVal
    PercentChange = PercentChange / OpenVal
     PercentChange = Format(PercentChange, "0.0000%")
        ws.Range("M" & TableRow) = PercentChange
    
    CloseVal = 0
    OpenVal = ws.Cells(i + 1, 3).Value
    YearlyChange = 0
    PercentChange = 0
    

        
        'Total Stock Volume
        
       TotalVol = TotalVol + ws.Cells(i, 7).Value
        ws.Range("N" & TableRow) = TotalVol
        TotalVol = 0
    

        
        TableRow = TableRow + 1
        
   Else
        
           
   TotalVol = TotalVol + ws.Cells(i, 7).Value
   

   End If
   



    Next i

Next ws

End Sub


