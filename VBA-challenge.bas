Attribute VB_Name = "Module1"
Sub testing()
    Dim ticker_symbol As String
    Dim yearly_change As Double
        yearly_change = 0
        
     Dim open_price As Double
         open_price = Cells(2, 3).Value
         
    Dim close_price As Double
        close_price = 0
        
        
    Dim percent_change As Double
        percent_change = 0
             
             
    Dim volume As Double
        volume = 0
        
    Dim LastRow As Long
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        

    
    Dim ticker_row As Integer
        ticker_row = 2
    
    
    For i = 2 To LastRow
    

   
        
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker_symbol = Cells(i, 1).Value
            close_price = Cells(i, 6).Value
            yearly_change = close_price - open_price
            percent_change = yearly_change / open_price
            volume = volume + Cells(i, 7).Value
            
            open_price = Cells(i + 1, 3).Value
            
            
            Range("I" & ticker_row).Value = ticker_symbol
            Range("J" & ticker_row).Value = yearly_change
            Range("K" & ticker_row).Value = percent_change
            Range("L" & ticker_row).Value = volume
         
   
                If yearly_change > 0 Then
                    Range("J" & ticker_row).Interior.ColorIndex = 4
                Else: Range("J" & ticker_row).Interior.ColorIndex = 3
                End If
            
            ticker_row = ticker_row + 1
            volume = 0
                
                
                
                

            
            
        Else
            volume = volume + Cells(i, 7).Value
            
            
            
            
            
        End If
        
    Next i
    



End Sub

Sub results()

    Dim LastRowsResults As Long
        LastRowsResults = Cells(Rows.Count, 9).End(xlUp).Row
                
    Dim MaxTicker As String
        MaxTicker = Cells(2, 9).Value
        
    Dim MaxPercent As Double
        MaxPercent = Cells(2, 11).Value
        
    Dim MinTicker As String
        MinTicker = Cells(2, 9).Value
        
    Dim MinPercent As Double
        MinPercent = Cells(2, 11).Value
        
    Dim MaxVTicker As String
        MaxVTicker = Cells(2, 9).Value
        
    Dim MaxVolume As Double
        MaxVolume = Cells(2, 12).Value
        
    For j = 2 To LastRowsResults
    
        If Cells(j + 1, 11).Value > MaxPercent Then
            MaxPercent = Cells(j + 1, 11).Value
            MaxTicker = Cells(j + 1, 9).Value
            
        End If
        
        If Cells(j + 1, 11).Value < MinPercent Then
            MinPercent = Cells(j + 1, 11).Value
            MinTicker = Cells(j + 1, 9).Value
            
        End If
        
        If Cells(j + 1, 12).Value > MaxVolume Then
            MaxVolume = Cells(j + 1, 12).Value
            MaxVTicker = Cells(j + 1, 9).Value
            
        End If
        
    
    Next j
    
    Range("O" & 2).Value = MaxTicker
    Range("P" & 2).Value = MaxPercent
    Range("O" & 3).Value = MinTicker
    Range("P" & 3).Value = MinPercent
    Range("O" & 4).Value = MaxVTicker
    Range("P" & 4).Value = MaxVolume

End Sub


