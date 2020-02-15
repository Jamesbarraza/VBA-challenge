Sub Fresh_Start()
For Each ws In Worksheets
Dim WorksheetName As String
WorksheetName = ws.Name
Sheets(ws.Name).Select
Columns("I:L").EntireColumn.AutoFit
Cells(1, 1).Select
Dim DateMinOpen As Variant
Dim DateMaxClose As Variant
Dim i As Double
Dim x As Double
Dim TotalV As Double
TotalV = 0
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 9).Value = Cells(1, 1).Value
i = 2
x = 2
Cells(x, 9).Value = Cells(i, 1).Value
DateMinOpen = Cells(i, 3).Value
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRow
        If Cells(i, 1).Value = Cells(x, 9).Value Then
            TotalV = TotalV + Cells(i, 7).Value
            DateMaxClose = Cells(i, 6).Value
           
        Else
            Cells(x, 10).Value = DateMaxClose - DateMinOpen
                If DateMinOpen <> 0 Then
                    Cells(x, 11).Value = (DateMaxClose - DateMinOpen) / DateMinOpen
                   
                Else
                    Cells(i, 11).Value = "Error"
                
                End If
            Cells(x, 11).Style = "Percent"
                If Cells(x, 10).Value > 0 Then
                    Cells(x, 10).Interior.ColorIndex = 4
                Else
                    Cells(x, 10).Interior.ColorIndex = 3
                End If
        Cells(x, 12).Value = TotalV
        DateMinOpen = Cells(i, 3).Value
        TotalV = Cells(i, 7).Value
        x = x + 1
        Cells(x, 9).Value = Cells(i, 1).Value

        End If

            Next i
            

Next ws
End Sub
