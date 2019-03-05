Sub easy()

    # 'Define ws as Worksheet to loop through each worksheet
    Dim ws As Worksheet
    # 'Define vol as Double to hold larger integers than Integer or Long
    Dim vol As Double
    
    # 'Activate each worksheet at the start of each loop with .Activate
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
        # 'Add headers to column I (Ticker) and J (Total Stock Volume)
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Total Stock Volume"
        
        # 'Set starting row as 2 with j variable for column I and J
        j = 2
        
        # 'Loop through each row of the data
        For i = 2 To 70926
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                
                # 'Add the name of each unique ticker from column A to column I
                Cells(j, 9).Value = Cells(i, 1).Value
                
                # 'Add 1 to j to proceed to the next row in column I
                j = j + 1
            End If
        Next i
    
        # 'Loop through each row of unique ticker from column I
        For a = 2 To 290
        
            # 'SumIf to calculate the total stock volume of each ticker
            # 'I found this to be a much more efficient approach as another for-loop took too long to compute
            Cells(a, 10).Value = Application.WorksheetFunction.SumIf(Range("A:A"), Cells(a, 9), Range("G:G"))
        Next a
        
    # 'Loop through the next worksheet until there are no more worksheets
    Next ws
    
End Sub
