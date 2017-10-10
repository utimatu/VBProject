Function GetHealthValue(Row As String, Column As String, Worksheet As String)
    Dim i As Integer
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If ActiveWorkbook.Worksheets(i).Name = Worksheet Then
            Dim x As Integer: x = 1
            Dim LastDate As Integer: LastDate = 1
            Do While True
                If ActiveWorkbook.Worksheets(i).Cells(x, 1) = Row Then
                    Dim y As Integer: y = 1
                    Do While True
                        If ActiveWorkbook.Worksheets(i).Cells(LastDate, y) = Column Then
                            GetHealthValue = ActiveWorkbook.Worksheets(i).Cells(x, y)
                        ElseIf ActiveWorkbook.Worksheets(i).Cells(LastDate, y) = "" Then
                            Exit Do
                        End If
                        y = y + 1
                    Loop
                ElseIf ActiveWorkbook.Worksheets(i).Cells(x, 1) = "Name" Then
                    LastDate = x
                End If
                x = x + 1
            Loop
        End If
    Next i
End Function

Function SUMTWONUMBERS(x As Integer, y As Integer)
    SUMTWONUMBERS = x + y
End Function

