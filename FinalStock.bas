Attribute VB_Name = "FinalStock"
Sub stockmarket()
    Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        
   
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
   
    Dim ticker_name As String
    Dim open_price As Double
        open_price = Cells(2, 3).Value
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim volume As Double
        volume = 0
    Dim row As Double
        row = 2
    Dim i As Long
   
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    For i = 2 To lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker_name = Cells(i, 1).Value
            Cells(row, 9).Value = ticker_name
            close_price = Cells(i, 6).Value
            yearly_change = close_price - open_price
            Cells(row, 10).Value = yearly_change
        If (open_price = 0 And close_price = 0) Then
            percent_change = 0
        ElseIf (open_price = 0 And close_price <> 0) Then
            percent_change = 1
        Else
            percent_change = yearly_change / open_price
            Cells(row, 11).Value = percent_change
            Cells(row, 11).NumberFormat = "0.00%"
        End If
    
            volume = volume + Cells(i, 7).Value
            Cells(row, 12).Value = volume
            row = row + 1
            open_price = Cells(i + 1, 3).Value
            volume = 1
        Else
            volume = volume + Cells(i, 7).Value
        End If
    Next i
    
    yearly_changelastrow = ws.Cells(Rows.Count, 9).End(xlUp).row
    
    For j = 2 To yearly_changelastrow
        If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
            Cells(j, 10).Interior.Color = vbGreen
        ElseIf Cells(j, 10).Value < 0 Then
            Cells(j, 10).Interior.Color = vbRed
        End If
    Next j
    
    Next ws
End Sub




