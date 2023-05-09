Sub WorkBookLoop()

    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Select
        Call TickerTesting
        Call Greatest
        Call Autofit
    Next

End Sub
