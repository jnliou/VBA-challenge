
Sub TickerTesting()


    'Create a Variable to Hold File Name, and Last Row
    Dim LastRowT As Long
    Dim Ticker As String
    
    Dim Summary_Table As Long
    Summary_Table = 2
    
    Dim openy As Double
    Dim closey As Double
    Dim Yearly As Double
    Dim PercentageC As Double
    Dim TotalV As Double
    
    TotalV = 0
    
    'Add titles to the columns
    
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
    'Determine the Last Row
        LastRowT = Cells(Rows.Count, "A").End(xlUp).Row
         
    
    'Add the ticker cells from column A into column I by looping
    
    For i = 2 To LastRowT

        TotalV = TotalV + Cells(i, 7).Value
        
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        
        openy = Cells(i, 3).Value
        
        End If
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        Ticker = Cells(i, 1).Value
        
        'Print the Ticker in the Summary Table
        
        Range("I" & Summary_Table).Value = Ticker
        
        closey = Cells(i, 6).Value
        
        Yearly = closey - openy
        
        Cells(Summary_Table, 10).Value = Yearly
        Cells(Summary_Table, 12).Value = TotalV
        
        'Add red or green backgrounds based on the value of the Yearly Change
            
            If Cells(Summary_Table, 10).Value > 0 Then
            Cells(Summary_Table, 10).Interior.ColorIndex = 4
        
        Else
            Cells(Summary_Table, 10).Interior.ColorIndex = 3
        
        End If
        
        
        'Calculate the Percentage Change
        
        PercentageC = Yearly / openy
        
        Cells(Summary_Table, 11).Value = PercentageC
        
        
        'Change the value from a decimal point to a percentage
    
        Cells(Summary_Table, 11).NumberFormat = "0.00%"
        
        
        'Add one to the summary table row
    
        Summary_Table = Summary_Table + 1

        'Reset Total Vol

        TotalV = 0
        
        
        End If
        
        
            
            Next i
    

End Sub
