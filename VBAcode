Attribute VB_Name = "Module1"
Sub dataloop():

    Dim ticker As String
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim totalstock As Double
    totalstock = 0
    Dim openprice As Double
    Dim closeprice As Double
    Dim previousamount As Long
    previousamount = 2
    
    Dim tablerow As Integer
    tablerow = 2
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    totalstock = totalstock + Cells(i, 7).Value
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            ticker = Cells(i, 1).Value
            Range("I" & tablerow).Value = ticker
            
            
            
            Range("L" & tablerow).Value = totalstock
            totalstock = 0
            
            openprice = Range("C" & previousamount)
            closeprice = Range("F" & i)
            
            yearlychange = closeprice - openprice
            Range("J" & tablerow).Value = yearlychange
            
            If openprice = 0 Then
                percentchange = 0
                    
                Else
                YearlyOpen = Range("C" & previousamount)
                percentchange = yearlychange / openprice
                        
            End If
            
            
            Range("K" & tablerow).Value = percentchange
            
        If Range("J" & tablerow).Value >= 0 Then
            Range("J" & tablerow).Interior.ColorIndex = 4
                    
                Else: Range("J" & tablerow).Interior.ColorIndex = 3
                
        End If
        
        tablerow = tablerow + 1
        previousamount = i + 1
        
            
        End If

        Next i
    
End Sub
