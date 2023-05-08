Sub WorkBookLoop()

    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Select
        Call TickerTesting
        Call Greatest
        Call Autofit
    Next

End Sub



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

Sub Greatest()

'Declare variables for the greatest % and total volume

Dim MaxP As Double
Dim MinP As Double
Dim GreatV As Double

'Insert titles into columns
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Volume"


'Find max value within column k
    MaxP = Application.WorksheetFunction.Max(Range("K:k"))

'display max value into separate column, convert value to percentage
        Cells(2, 16) = MaxP
        Cells(2, 16).NumberFormat = "0.00%"

'Find the min Value of column k
    MinP = Application.WorksheetFunction.Min(Range("K:k"))

'display min value into separate column, converet value to percentage
        Cells(3, 16) = MinP
        Cells(3, 16).NumberFormat = "0.00%"

'Find the greatest total volume in column L
    GreatV = Application.WorksheetFunction.Max(Range("L:l"))

'display the greatest toal volume in separate column
        Cells(4, 16) = GreatV

'find the location of the ticker that corresponds to the above values

'declare variables that will be used for the Ticker
Dim inc_t As Integer
Dim dec_t As Integer
Dim totalvolt As Integer

'use match function, to find value corresponding to our Max Percentage, Min Percentage, and Greatest Total Volume
inc_t = WorksheetFunction.Match(MaxP, Range("K:K"), 0)
dec_t = WorksheetFunction.Match(MinP, Range("K:K"), 0)
totalvolt = WorksheetFunction.Match(GreatV, Range("L:L"), 0)

        ' assign the Ticker values to column O
        Range("O2") = Cells(inc_t, 9)
        Range("O3") = Cells(dec_t, 9)
        Range("O4") = Cells(totalvolt, 9)
        

End Sub

Sub Autofit()

'Autofit all the title columns in the worksheet
    Range("A1:P4").Columns.Autofit
End Sub