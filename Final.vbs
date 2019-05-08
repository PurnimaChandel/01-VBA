Sub Stock_Volume()

Dim Summary_row_number As Double
Dim Intitial_price As Double
Dim Last_row As Double
Dim Last_Row1 As Double
Dim Total_Volume As Double
Dim Close_Price As Double
Dim i As Long


For Each ws In Worksheets
    Last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    
    With ws.Sort
     .SortFields.Add Key:=Range("A1"), Order:=xlAscending
     .SortFields.Add Key:=Range("B1"), Order:=xlAscending
     .Header = xlYes
     .Apply
    End With
    
    
    Summary_row_number = 2
    Total_Volume = 0
    Intitial_price = ws.Cells(2, 3).Value
    
    For i = 2 To Last_row
       If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
       
            ws.Cells(Summary_row_number, 9).Value = ws.Cells(i, 1).Value
            
            
            Close_Price = ws.Cells(i, 6).Value
            ws.Cells(Summary_row_number, 11).Value = Close_Price - Intitial_price
            If ws.Cells(Summary_row_number, 11).Value <> "" Then
               If ws.Cells(Summary_row_number, 11).Value > 0 Then
                  ws.Cells(Summary_row_number, 11).Interior.Color = RGB(0, 255, 0)
               Else
                 ws.Cells(Summary_row_number, 11).Interior.Color = RGB(255, 0, 0)
            End If
          End If
            
            If Intitial_price = 0 Then
                ws.Cells(Summary_row_number, 12).Value = "NA"
                Else
                    ws.Cells(Summary_row_number, 12).Value = (ws.Cells(Summary_row_number, 11).Value / Intitial_price)
                    ws.Cells(Summary_row_number, 12).NumberFormat = "0.00%"
            
            End If
                
       
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        ws.Cells(Summary_row_number, 10).Value = Total_Volume
          
        Summary_row_number = Summary_row_number + 1
        Intitial_price = ws.Cells(i + 1, 3).Value
        Total_Volume = 0
              
         Else
            Total_Volume = Total_Volume + Cells(i, 7).Value
              
            
        End If
                       
               
    Next i
    
    
Next ws


End Sub
