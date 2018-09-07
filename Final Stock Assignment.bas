Attribute VB_Name = "Module1"
Sub stock()
    
For Each ws In Worksheets
ws.Activate


    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Value"
    
    Dim Ticker As String
    Dim Total_Volume As Double
    Dim Summary As Integer
    Dim last_row As Long
    
    
    Total_Volume = 0
    Summary = 2
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For i = 2 To last_row
       
       If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
           Total_Volume = Cells(i, 7).Value + Total_Volume
           Cells(Summary, 10).Value = Total_Volume
       
       ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           Total_Volume = Cells(i, 7).Value
           Ticker = Cells(i, 1).Value
           
           Cells(Summary, 9).Value = Ticker
           Summary = Summary + 1
       
       End If
        
     Next i
     
Next ws
    
End Sub

