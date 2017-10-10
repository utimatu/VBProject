Public Function TotalEmployeeWeeks(StartDate As String, EndDate As String)
    Dim nameArray() As String
    Dim countArray() As String
    ReDim Preserve nameArray(100)
    ReDim Preserve countArray(100)
    Dim x As Integer: x = 5
    Dim y As Integer: y = 0
    For y = 0 To 100
        nameArray(y) = ""
        countArray(y) = 0
    Next
    y = 0
    Dim i As Integer
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If DateInside(ActiveWorkbook.Worksheets(i).Name, StartDate, EndDate) Then
            x = 5
            Do While ActiveWorkbook.Worksheets(i).Cells(x, 1) <> "Totals"
                If ActiveWorkbook.Worksheets(i).Cells(x, 10).Value <> "" Then
                    For y = 0 To UBound(nameArray, 1)
                        If y = UBound(nameArray, 1) Then
                            ReDim Preserve nameArray(2 * UBound(nameArray, 1))
                            ReDim Preserve countArray(2 * UBound(countArray, 1))
                            nameArray(y) = ActiveWorkbook.Worksheets(i).Cells(x, 1)
                            countArray(y) = 1
                            Exit For
                        ElseIf nameArray(y) = ActiveWorkbook.Worksheets(i).Cells(x, 1) Then
                            countArray(y) = countArray(y) + 1
                            Exit For
                        ElseIf countArray(y) = 0 Then
                            nameArray(y) = ActiveWorkbook.Worksheets(i).Cells(x, 1)
                            countArray(y) = 1
                            Exit For
                        End If
                    Next
                End If
                x = x + 1
            Loop
        End If
    Next
    x = x + 1
    Dim s As String: s = ""
    For y = 0 To UBound(nameArray, 1)
        s = s + nameArray(y) + ": " + CStr(countArray(y)) + vbCrLf
    Next
    PrintToFile (s)
End Function

Public Function TotalEmployeeDays(StartDate As String, EndDate As String)
    Dim nameArray() As String
    Dim countArray() As String
    ReDim Preserve nameArray(100)
    ReDim Preserve countArray(100)
    Dim x As Integer: x = 5
    Dim y As Integer: y = 0
    Dim z As Integer: z = 0
    Dim DayCount As Integer: DayCount = 0
    For y = 0 To 100
        nameArray(y) = ""
        countArray(y) = 0
    Next
    y = 0
    Dim i As Integer
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If DateInside(ActiveWorkbook.Worksheets(i).Name, StartDate, EndDate) Then
            DayCount = DayCount + 5
            x = 5
            Do While ActiveWorkbook.Worksheets(i).Cells(x, 1) <> "Totals"
                For z = 0 To 6
                    If ActiveWorkbook.Worksheets(i).Cells(x, 3 + z).Value <> "" Then
                        For y = 0 To UBound(nameArray, 1)
                            If y = UBound(nameArray, 1) Then
                                ReDim Preserve nameArray(2 * UBound(nameArray, 1))
                                ReDim Preserve countArray(2 * UBound(countArray, 1))
                                nameArray(y) = ActiveWorkbook.Worksheets(i).Cells(x, 1)
                                countArray(y) = 1
                                Exit For
                            ElseIf nameArray(y) = ActiveWorkbook.Worksheets(i).Cells(x, 1) Then
                                countArray(y) = countArray(y) + 1
                                Exit For
                           ElseIf countArray(y) = 0 Then
                                nameArray(y) = ActiveWorkbook.Worksheets(i).Cells(x, 1)
                                countArray(y) = 1
                                Exit For
                            End If
                        Next
                    End If
                Next
                x = x + 1
            Loop
        End If
    Next
    x = x + 1
    Dim s As String: s = "DayCount: " + CStr(DayCount) + vbCrLf
    For y = 0 To UBound(nameArray, 1)
        s = s + nameArray(y) + ": " + CStr(countArray(y)) + vbCrLf
    Next
    PrintToFile (s)
End Function

Public Function AverageNumberOfEmployees(StartDate As String, EndDate As String) As Integer
    Dim nameArray() As String
    Dim countArray() As Boolean
    ReDim Preserve nameArray(100)
    ReDim Preserve countArray(100, 12)
    Dim x As Integer: x = 5
    Dim y As Integer: y = 0
    Dim z As Integer: z = 0
    For y = 0 To 100
        nameArray(y) = ""
        For z = 0 To 11
            countArray(y, z) = False
        Next
    Next
    y = 0
    z = 0
    Dim i As Integer
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If DateInside(ActiveWorkbook.Worksheets(i).Name, StartDate, EndDate) Then
            x = 5
            Do While ActiveWorkbook.Worksheets(i).Cells(x, 1) <> "Totals"
                If ActiveWorkbook.Worksheets(i).Cells(x, 10).Value <> "" Then
                    For y = 0 To UBound(nameArray, 1)
                        If y = UBound(nameArray, 1) Then
                            ReDim Preserve nameArray(2 * UBound(nameArray, 1))
                            ReDim Preserve countArray(2 * UBound(countArray, 1), 12)
                            nameArray(y) = ActiveWorkbook.Worksheets(i).Cells(x, 1)
                            countArray(y, GetMonthFromDate(ActiveWorkbook.Worksheets(i).Name) - 1) = True
                            Exit For
                        ElseIf nameArray(y) = ActiveWorkbook.Worksheets(i).Cells(x, 1) Then
                            countArray(y, GetMonthFromDate(ActiveWorkbook.Worksheets(i).Name) - 1) = True
                            Exit For
                        ElseIf nameArray(y) = "" Then
                            nameArray(y) = ActiveWorkbook.Worksheets(i).Cells(x, 1)
                            countArray(y, GetMonthFromDate(ActiveWorkbook.Worksheets(i).Name) - 1) = True
                            Exit For
                        End If
                    Next
                End If
                x = x + 1
            Loop
        End If
    Next
    Dim total As Integer: total = 0
    For y = 0 To 11
        For z = 0 To UBound(nameArray, 1)
            If countArray(z, y) Then
                total = total + 1
            End If
        Next
    Next
    AverageNumberOfEmployees = total
End Function

