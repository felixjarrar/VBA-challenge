Sub challenge1():

    ' Dimensions
    Dim total As Double

    ' see row number of last row with data
    countRows = Cells(Rows.Count, "A").End(xlUp).Row

    ' Have the title row get set up 
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"

    For i = 2 To countRows

        ' Print results is ticker changes
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Results stored in variable
            total = total + Cells(i, 7).Value

            ' Ticker symbol is printed 
            Range("I" & 2 + j).Value = Cells(i, 1).Value

            ' Total printed
            Range("J" & 2 + j).Value = total

            ' Reset Total
            total = 0

            ' Move to the next row
            j = j + 1

        ' Here, the else adds to total volume
        Else
            total = total + Cells(i, 7).Value

        End If

    Next i

End Sub
