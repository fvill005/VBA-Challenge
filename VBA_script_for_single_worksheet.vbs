Attribute VB_Name = "Module1"
Sub multi_year_analysis()

Dim ticker As String
Dim year_open As Double
Dim year_close As Double
Dim year_change As Double
Dim percent_change As Double
Dim total_vol As Double
Dim lastrow As Long
Dim summary_table_row As Double
Dim greatest_inc_ticker As String
Dim greatest_inc_percent As Double
Dim greatest_dc_ticker As String
Dim greatest_dc_percent As Double
Dim greatest_vol_ticker As String
Dim greatest_vol As Double


summary_table_row = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
total_vol = 0

Cells(1, 8).Value = "<Ticker>"

Cells(1, 9).Value = "<Yearly Change>"

Cells(1, 10).Value = "<Percent Change>"

Cells(1, 11).Value = "<Total Stock Volume>"

Cells(2, 15).Value = "<Greatest % Increase>"

Cells(3, 15).Value = "< Greatest % Decrease>"

Cells(4, 15).Value = "<Greatest Total Volume>"

Cells(1, 16).Value = "<Ticker>"

Cells(1, 17).Value = "<Value>"

Range("O:O").Columns.AutoFit



For i = 2 To lastrow
    total_vol = total_vol + Cells(i, 7).Value
   
    If year_open = 0 Then
            year_open = Cells(i, 3).Value
        End If
   
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
         ticker = Cells(i, 1).Value
 
        Cells(summary_table_row, 8).Value = ticker
        Cells(summary_table_row, 11).Value = total_vol
 
        year_close = Cells(i, 6).Value
        year_change = year_close - year_open
       
        Cells(summary_table_row, 9) = year_change
       
      If (year_open = 0 And year_close = 0) Then
        percent_change = 0
    
    ElseIf (year_open = 0 And year_close <> 0) Then
        percent_change = 1
        
    Else
        year_open = Cells(i, 3).Value
        year_close = Cells(i, 6).Value
        percent_change = (year_change / year_open)
    End If
    
    Cells(summary_table_row, 10).NumberFormat = "0.00%"
    Cells(summary_table_row, 10).Value = percent_change
   
           
       
summary_table_row = summary_table_row + 1
        total_vol = 0
        year_open = 0
       
       
    End If
   
        If Cells(i, 9) >= 0 Then
            Cells(i, 9).Interior.ColorIndex = 4
           
        Else
            Cells(i, 9).Interior.ColorIndex = 3
        End If
       
    Next i
   
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
   
    greatest_inc_ticker = Cells(2, 8).Value
    greatest_inc_percent = Cells(2, 10).Value
    greatest_dc_ticker = Cells(2, 8).Value
    greatest_dc_percent = Cells(2, 10).Value
    greatest_vol_ticker = Cells(2, 8).Value
    greatest_vol = Cells(2, 11).Value
   
    For i = 2 To lastrow
        If Cells(i, 10).Value = Application.WorksheetFunction.Max(Range("J2:J" & lastrow)) Then
            greatest_inc_percent = Cells(i, 10).Value
            greatest_inc_ticker = Cells(i, 8).Value
            Cells(2, 16).Value = greatest_inc_ticker
            Cells(2, 17).Value = greatest_inc_percent
           
        End If
       
        If Cells(i, 10).Value = Application.WorksheetFunction.Min(Range("J2:J" & lastrow)) Then
            greatest_dc_percent = Cells(i, 10).Value
            greatest_dc_ticker = Cells(i, 8).Value
            Cells(3, 16).Value = greatest_dc_ticker
            Cells(3, 17).Value = greatest_dc_percent
       
        End If
       
         If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow)) Then
            greatest_vol = Cells(i, 11).Value
            greatest_vol_ticker = Cells(i, 8).Value
            Cells(4, 16).Value = greatest_vol_ticker
            Cells(4, 17).Value = greatest_vol
        End If
       
        Next i
   
   
   
   









    
   

End Sub
Sub multi_year_analysis_all_ws()

End Sub
