Sub Stock()

'Displaying headers in row 1
Dim headers() As Variant
Dim Main As Worksheet
Dim wb As Workbook

Set wb = ActiveWorkbook

headers() = Array("<ticker>", "<date>", "<open>", "<high>", "<low>", "<close>", "<vol>", " ", "Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", " ", " ", " ", "Ticker", "Value")

For Each Main In wb.Sheets
    With Main
    .Rows(1).Value = " "
    For i = LBound(headers()) To UBound(headers())
    .Cells(1, 1 + i).Value = headers(i)
    
    Next i
    .Rows(1).Font.Bold = True
    End With
Next Main

    For Each Main In Worksheets
    'Variables
    Dim Ticker As String
    Ticker = " "
    Dim Total_Vol As Double
    Total_Vol = 0
    Dim Start_Price As Double
    Start_Prce = Main.Cells(2, 3).Value
    Dim End_Price As Double
    End_Price = 0
    Dim Yearly_Change As Double
    Yearly_Change = 0
    Dim Percent_Change As Double
    Percent_Change = 0
    Dim Lastrom As Long
    Lastrow = Main.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Bonus Variables
    Dim Big_Ticker As String
    Big_Ticker = " "
    Dim Lil_Ticker As String
    Lil_Ticker = " "
    Dim Vol_Ticker As String
    Vol_Ticker = " "
    Dim Big_Percent As Double
    Big_Percent = 0
    Dim Lil_Percent As Double
    Lil_Percent = 0
    Dim Big_Vol As Double
    Big_Vol = 0
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
    'For loop to run through all data
    For i = 2 To Lastrow
        
        'If statement to find unequal tickers and there data
        If Main.Cells(i + 1, 1).Value <> Main.Cells(i, 1).Value Then
            
            Ticker = Main.Cells(i, 1).Value
            End_Price = Main.Cells(i, 6).Value
            Yearly_Change = End_Price - Start_Price
            
            'If statement to varify the starting price is not zero
            If Start_Price <> 0 Then
                
                Percent_Change = (Yearly_Change / Start_Price) * 100
            
            End If
            
            'displaying data
            Total_Vol = Total_Vol + Main.Cells(i, 7).Value
            Main.Range("I" & Summary_Table_Row).Value = Ticker
            Main.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            'changing color of table
            If (Yearly_Change > 0) Then
                
                Main.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
            ElseIf (Yearly_Change <= 0) Then
                
                Main.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
            End If
            
            'displaying data
            Main.Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
            Main.Range("L" & Summary_Table_Row).Value = Total_Vol
            
            'resetting variables
            Summary_Table_Row = Summary_Table_Row + 1
            Start_Price = Main.Cells(i + 1, 3).Value
        
            'Bonus
            'see if we have the greatest % increase
             If (Percent_Change > Big_Percent) Then
            
                Big_Percent = Percent_Change
                Big_Ticker = Ticker
            
            'see if we have the greatest % decrease
            ElseIf (Percent_Change < Lil_Percent) Then
        
                Lil_Percent = Percent_Change
                Lil_Ticker = Ticker
            
            End If
        
            'see if we have the greatest total volume
            If (Total_Vol > Big_Vol) Then
            
                Big_Vol = Total_Vol
                Vol_Ticker = Ticker
            
            End If
        
            'Reset Variables
            Percent_Change = 0
            Total_Vol = 0
        
        Else
            
            Total_Vol = Total_Vol + Main.Cells(i, 7).Value
        
        End If
    
    Next i

    'Display Bonus information
    Main.Range("Q2").Value = (CStr(Big_Percent))
    Main.Range("Q3").Value = (CStr(Lil_Percent))
    Main.Range("Q4").Value = Big_Vol
    Main.Range("P2").Value = Big_Ticker
    Main.Range("P3").Value = Lil_Ticker
    Main.Range("P4").Value = Vol_Ticker
    Main.Range("O2").Value = "Greatest % Increased"
    Main.Range("O3").Value = "Greatest % Decreased"
    Main.Range("O4").Value = "Greatest Total Volume"
    
Next Main
End Sub

