
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