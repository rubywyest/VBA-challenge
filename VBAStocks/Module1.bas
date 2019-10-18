Attribute VB_Name = "Module1"
Sub Docalculation()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call stock_basic
        xSh.Range("A1:Q" & Cells(Rows.Count, 1).End(xlUp).Row).Columns.AutoFit
    Next
    Application.ScreenUpdating = True
End Sub
 
      
Sub stock_basic()
 
  Dim ticker As String
  Dim open_price As Double
  Dim close_price As Double
  Dim yearly_change As Double
  Dim percent_change As Double
  Dim total_stock_volume As Double
  Dim lastrow As Long
  
    yearly_change = 0
    percent_change = 0
    total_stock_volume = 0
   
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Dim summary_total_row As Long
    summary_total_row = 2
  
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
   
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
  
    
  
  ' Loop
   For I = 2 To lastrow
   
    If Cells(I - 1, 1).Value <> Cells(I, 1).Value Then
    
        If (I <> 2) Then
            Range("I" & summary_total_row).Value = ticker
            Range("J" & summary_total_row).Value = close_price - open_price
            
            
            If (Range("J" & summary_total_row).Value >= 0) Then
                Range("J" & summary_total_row).Interior.ColorIndex = 4
            Else
                Range("J" & summary_total_row).Interior.ColorIndex = 3
            End If
            
            
            
            If open_price > 0 Then
                Range("K" & summary_total_row).Value = (close_price - open_price) / open_price
                Range("K" & summary_total_row).NumberFormat = "0.00%"
            End If
            Range("l" & summary_total_row).Value = total_stock_volume
            summary_total_row = summary_total_row + 1
            total_stock_volume = 0
            
            
        End If
    
           ticker = Cells(I, 1).Value
            open_price = Cells(I, 3).Value
            close_price = Cells(I, 6).Value
            total_stock_volume = total_stock_volume + Cells(I, 7).Value
        
            
        Else
            close_price = Cells(I, 6).Value
           total_stock_volume = total_stock_volume + Cells(I, 7).Value
           
              
            If (I = lastrow) Then
                Range("I" & summary_total_row).Value = ticker
                If open_price > 0 Then
                Range("j" & summary_total_row).Value = close_price - open_price
                End If
                Range("k" & summary_total_row).Value = (close_price - open_price) / open_price
                Range("l" & summary_total_row).Value = total_stock_volume
            End If
            
        End If
        
  Next I


Dim lastrowsum As Long
    
    lastrowsum = Cells(Rows.Count, 9).End(xlUp).Row
    

Dim grt_percent_increase As Double
Dim grt_percent_decrease As Double
Dim grt_volume As Double
    Dim grt_percent_ticker As String
    Dim grt_percent_decrease_ticker As String
    Dim grt_volume_ticker As String
    
    grt_percent_increase = 0
    grt_percent_decrease = 0
    grt_volume = 0

    

    For j = 2 To lastrowsum

    If (Cells(j, 11) > grt_percent_increase) Then
        grt_percent_increase = Cells(j, 11)
        grt_percent_ticker = Cells(j, 9)
    End If
    If (Cells(j, 11) < grt_percent_decrease) Then
        grt_percent_decrease = Cells(j, 11)
        grt_percent_decrease_ticker = Cells(j, 9)
    End If
    If (Cells(j, 12) > grt_volume) Then
        grt_volume = Cells(j, 12)
        grt_volume_ticker = Cells(j, 9)
    End If
    Next j
    
   Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("P2").Value = grt_percent_ticker
    Range("P3").Value = grt_percent_decrease_ticker
    Range("P4").Value = grt_volume_ticker
    Range("Q1").Value = "Value"
    Range("Q2").Value = grt_percent_increase
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").Value = grt_percent_decrease
    Range("Q3").NumberFormat = "0.00%"
    Range("Q4").Value = grt_volume
    
    
End Sub



