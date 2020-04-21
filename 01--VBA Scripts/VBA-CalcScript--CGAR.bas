Attribute VB_Name = "Module1"

Sub VBA_Challenge()

    Dim ticker As String
    
    Dim tick_initial_price As Double
    Dim tick_total_change As Double
    Dim tick_percent_change As Double
    Dim tick_stock_volume As Double
    
    Dim MAX_pct as Integer
    Dim MIN_pct as Integer
    Dim MAX_vol as Integer

    Dim tick_string_calcs As Integer
    tick_string_calcs = 2
    
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "% Change"
    Cells(1, 13).Value = "Total Stock Volume"
    
    Cells(1, 15).Value = "Max % Change"
    Cells(1, 16).Value = "Min % Change"
    Cells(1, 17).Value = "Max Stock Volume"

    
    Application.ScreenUpdating = False
    n = Range("A1", Range("A1").End(xlDown)).Rows.Count
    Range("A1").Select
    
    For i = 2 To n
        
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        
            ticker = Cells(i, 1).Value
            
            tick_initial_price = Cells(i, 3).Value
            tick_total_change = tick_total_change + (Cells(i, 3).Value - Cells(i, 6).Value)
            tick_stock_volume = tick_stock_volume + Cells(i, 7).Value

            
            Range("J" & tick_string_calcs).Value = ticker
            Range("K" & tick_string_calcs).Value = tick_total_change
            Range("L" & tick_string_calcs).Value = tick_percent_change
            Range("M" & tick_string_calcs).Value = tick_stock_volume
            
            tick_string_calcs = (tick_string_calcs + 1)
            
            tick_stock_volume = 0
            tick_total_change = 0
            tick_percent_change = 0
            
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            tick_closing_price = Cells(i, 6).Value
            
            tick_total_change = tick_initial_price - tick_closing_price
                
        Else
            
            tick_stock_volume = tick_stock_volume + Cells(i, 7).Value
        
        End If
        
        
        If tick_initial_price <> 0 Then
        
            tick_percent_change = (tick_total_change / tick_initial_price) * 100
            
        Else
            
            tick_percent_change = 0
        
        End If
    
                 
        ActiveCell.Offset(1, 0).Select
    
    
        If Cells(i, 11).Value > 0 Then
            
            Cells(i, 11).Interior.ColorIndex = 4
            
        ElseIf Cells(i, 11).Value < 0 Then
            
            Cells(i, 11).Interior.ColorIndex = 3
        
        End If
    
    Next i
    
MAX_pct = WorksheetFunction.Max(i, 12)
If Cells(i, 12).Value = MAX_pct Then
    Cells(2, 15).Value = Cells(i, 10).Value
    Cells(3, 15).Value = Cells(i, 12).Value
End If

MIN_pct = WorksheetFunction.Min(i, 12)
If Cells(i, 12).Value = MIN_pct Then
    Cells(2, 16).Value = Cells(i, 10).Value
    Cells(3, 15).Value = Cells(i, 12).Value
End If

MAX_vol = WorksheetFunction.Max(i, 13)
If Cells(i, 13).Value = MAX_vol 
    Cells(2, 17).Value = Cells(i, 10).Value
    Cells(3, 17).Value = Cells(i, 13).Value

    Application.ScreenUpdating = True
    
End Sub
Sub Reset()
Attribute Reset.VB_ProcData.VB_Invoke_Func = " \n14"

    Columns("J:Z").Select
    Selection.Delete Shift:=xlToLeft
    Range("H1").Select
End Sub


