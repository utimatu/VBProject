Sub UnhideAllSheets()
    Dim wsSheet As Worksheet
    For Each wsSheet In ActiveWorkbook.Worksheets
        wsSheet.Visible = xlSheetVisible
    Next wsSheet
End Sub

Sub CopyLastSheet()
    Dim i As Integer: i = ActiveWorkbook.Worksheets.Count
    Dim j As Integer
    Dim regEx As New RegExp
    Dim Dat As Date: Dat = NextMonday()
    Dim Da As String: Da = Format(Dat, "m-d-yy")
    Dim OldDa As String
    If WorksheetExists(Da) Then
        Da = Da & "(1)"
    End If
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "^(?:1[0-2]|0?[1-9])-(?:0?[1-9]|[12]\d|3[01])-(?:\d{2}|\d{4})$"
    End With
    For j = i To 0 Step -1
        If regEx.Test(ActiveWorkbook.Worksheets(j).Name) Then
            OldDa = ActiveWorkbook.Worksheets(j).Name
            Exit For
        End If
    Next
    ActiveWorkbook.Worksheets("Template").Copy After:=ActiveWorkbook.Worksheets(j)
    ActiveWorkbook.Worksheets("Template (2)").Activate
    ActiveSheet.Visible = -1
    ActiveSheet.Name = Da
    ActiveSheet.Range("$C$2").Value = Da
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With
    EnterEmployees
    FixTotals
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If (ActiveWorkbook.Worksheets(i).Name = "Truck") Then
            For j = 1 To 32767
                If (ActiveWorkbook.Worksheets(i).Cells(j, 1).Value = "") Then
                    ActiveWorkbook.Worksheets(i).Cells(j, 1).Value = Da
                    Exit For
                End If
            Next j
        ElseIf (ActiveWorkbook.Worksheets(i).Name = "Health") Then
            Call AddColumnToInsurance(i, Da, OldDa)
        ElseIf (ActiveWorkbook.Worksheets(i).Name = "Post") Then
            Call AddColumnToInsurance(i, Da, OldDa)
        ElseIf (ActiveWorkbook.Worksheets(i).Name = "Pre") Then
            Call AddColumnToInsurance(i, Da, OldDa)
        ElseIf (ActiveWorkbook.Worksheets(i).Name = "Loan") Then
            Call AddColumnToLoan(i, Da, OldDa)
        ElseIf (ActiveWorkbook.Worksheets(i).Name = "Bank") Then
            Call AddColumnToBank(i, Da, OldDa)
        End If
    Next i
    ActiveWorkbook.Worksheets("Payroll").Cells(1, 1).Value = Da
End Sub

Sub FixTotals()
    Dim i As Integer
    For i = 1 To ActiveSheet.UsedRange.Rows.Count + 1
        If ActiveSheet.Cells(i, 1).Value = "Totals" Then
            Dim Column As Integer
            For Column = 10 To 19
                ActiveSheet.Cells(i, Column).Formula = "=SUM(" & ActiveSheet.Range(Cells(5, Column), Cells(i - 1, Column)).Address & ")"
            Next
        End If
    Next
End Sub

Sub AddColumnToBank(index As Integer, Da As String, OldDa As String)
    Dim i As Integer
    Dim j As Integer
    Dim K As Integer
    Dim Temp As String
    Dim EmployeesList(50) As String
    Dim BankList(50) As String
    For i = 1 To ActiveWorkbook.Worksheets(index).UsedRange.Rows.Count + 1
        If (ActiveWorkbook.Worksheets(index).Cells(i, 1) = "Name") Then
            For j = 1 To 4096
                If (ActiveWorkbook.Worksheets(index).Cells(i, j) = OldDa) Then
                    ActiveWorkbook.Worksheets(index).Cells(i, j + 2).NumberFormat = "@"
                    ActiveWorkbook.Worksheets(index).Cells(i, j + 2).Value = Da
                    ActiveWorkbook.Worksheets(index).Cells(i, j + 3).Value = "Balance"
                    For K = 1 To 50
                        If ActiveWorkbook.Worksheets(index).Cells(i + K, j) = "" Then
                            Exit For
                        ElseIf ActiveWorkbook.Worksheets(index).Cells(i + K, j) <> "" Then
                            Temp = "=IFERROR(IF(VLOOKUP(" & ActiveWorkbook.Worksheets(index).Cells(i + K, 1).Address(False, True) & ", INDIRECT(""'""&"
                            Temp = Temp & ActiveWorkbook.Worksheets(index).Cells(i, j + 2).Address(True, False) & "&""'!$A$5:$P$100""), 14, FALSE)="""", 0, VLOOKUP("
                            Temp = Temp & ActiveWorkbook.Worksheets(index).Cells(i + K, 1).Address(False, True) & ", INDIRECT(""'""&"
                            Temp = Temp & ActiveWorkbook.Worksheets(index).Cells(i, j + 2).Address(True, False) & "&""'!$A$5:$P$100""), 14, FALSE)), 0)"
                            ActiveWorkbook.Worksheets(index).Cells(i + K, j + 2).Formula = Temp
                            Temp = "=IFERROR(" & ActiveWorkbook.Worksheets(index).Cells(i + K, j + 1).Address(False, False) & "+"
                            Temp = Temp & ActiveWorkbook.Worksheets(index).Cells(i + K, j + 2).Address(False, False) & ", 0)"
                            ActiveWorkbook.Worksheets(index).Cells(i + K, j + 3).Formula = Temp
                        End If
                    Next K
                End If
            Next j
        End If
    Next i
End Sub

Public Function AddColumnToLoan(index As Integer, Da As String, OldDa As String)
    Dim i As Integer
    Dim j As Integer
    Dim K As Integer
    Dim Temp As String
    For i = 1 To ActiveWorkbook.Worksheets(index).UsedRange.Rows.Count + 1
        If (ActiveWorkbook.Worksheets(index).Cells(i, 1) = "Name") Then
            For j = 1 To 4096
                If (ActiveWorkbook.Worksheets(index).Cells(i, j) = OldDa) And (ActiveWorkbook.Worksheets(index).Cells(i + 1, j + 1).Value <> 0) Then
                    ActiveWorkbook.Worksheets(index).Cells(i, j + 2).NumberFormat = "@"
                    ActiveWorkbook.Worksheets(index).Cells(i, j + 2).Value = Da
                    ActiveWorkbook.Worksheets(index).Cells(i, j + 3).Value = "Balance"
                    For K = 1 To 50
                        On Error Resume Next:
                            If ActiveWorkbook.Worksheets(index).Cells(i + K, j).Value = "" Then
                                Exit For
                        ElseIf ActiveWorkbook.Worksheets(index).Cells(i + K, j) <> "" Then
                            Temp = "=IF(VLOOKUP(" & ActiveWorkbook.Worksheets(index).Cells(i + K, 1).Address(False, True) & ", INDIRECT(""'""&"
                            Temp = Temp & ActiveWorkbook.Worksheets(index).Cells(i, j + 2).Address(True, False) & "&""'!""&""$A$4:$S$100""), MATCH(""Pay"", INDIRECT(""'""&"
                            Temp = Temp & ActiveWorkbook.Worksheets(index).Cells(i, j + 2).Address(True, False) & "&""'!""&""$A$4:S$4""), 0), FALSE) > 20, "
                            Temp = Temp & ActiveWorkbook.Worksheets(index).Cells(i + K, 2).Address(False, True) & ", 0)"
                            ActiveWorkbook.Worksheets(index).Cells(i + K, j + 2).Formula = Temp
                            Temp = "= " & ActiveWorkbook.Worksheets(index).Cells(i + K, j + 1).Address(False, False) & " - "
                            Temp = Temp & ActiveWorkbook.Worksheets(index).Cells(i + K, j + 2).Address(False, False)
                            ActiveWorkbook.Worksheets(index).Cells(i + K, j + 3).Formula = Temp
                        End If
                    Next K
                End If
            Next j
        End If
    Next i
End Function

Public Function AddColumnToInsurance(index As Integer, Da As String, OldDa As String)
    Dim i As Integer
    Dim j As Integer
    Dim K As Integer
    Dim Temp As String
    For i = (ActiveWorkbook.Worksheets(index).UsedRange.Rows.Count + 1) To 1 Step -1
        If (ActiveWorkbook.Worksheets(index).Cells(i, 1) = "Name") Then
            For j = 1 To 4096
                If (ActiveWorkbook.Worksheets(index).Cells(i, j) = OldDa) Then
                    ActiveWorkbook.Worksheets(index).Cells(i, j + 2).NumberFormat = "@"
                    ActiveWorkbook.Worksheets(index).Cells(i, j + 2).Value = Da
                    ActiveWorkbook.Worksheets(index).Cells(i, j + 3).Value = "Balance"
                    For K = 1 To 50
                        If ActiveWorkbook.Worksheets(index).Cells(i + K, j) = "" Then
                            Exit For
                        ElseIf ActiveWorkbook.Worksheets(index).Cells(i + K, j) <> "" Then
                            Temp = "=IF(" & ActiveWorkbook.Worksheets(index).Cells(i + K, 1).Address(False, True)
                            Temp = Temp & "=""Carter, Tim"", " & ActiveWorkbook.Worksheets(index).Cells(i + K, 2).Address(False, True) & ", IF(" & ActiveWorkbook.Worksheets(index).Cells(i + K, 1).Address(False, True)
                            Temp = Temp & " = ""Romano, Nancy"", " & ActiveWorkbook.Worksheets(index).Cells(i + K, 2).Address(False, True) & ", IF(VLOOKUP(" & ActiveWorkbook.Worksheets(index).Cells(i + K, 1).Address(False, True)
                            Temp = Temp & ", INDIRECT(""'""&" & ActiveWorkbook.Worksheets(index).Cells(i, j + 2).Address(True, False) & "&""'!""&""$A$1:$Z$55""), MATCH(""Pay"", INDIRECT(""'""&"
                            Temp = Temp & ActiveWorkbook.Worksheets(index).Cells(i, j + 2).Address(True, False) & "&""'!""&""$A$4:$K$4""), 0), FALSE)< " & "IF(" & ActiveWorkbook.Worksheets(index).Cells(i + K, 1).Address(False, True) & "=""Lane, Everett"", 5, 10)"
                            Temp = Temp & ", 0, IF(VLOOKUP("
                            Temp = Temp & ActiveWorkbook.Worksheets(index).Cells(i + K, 1).Address(False, True) & ", INDIRECT(""'""&" & ActiveWorkbook.Worksheets(index).Cells(i, j + 2).Address(True, False)
                            Temp = Temp & "&""'!""&""$A$1:$Z$55""), MATCH(""Pay"", INDIRECT(""'""&" & ActiveWorkbook.Worksheets(index).Cells(i, j + 2).Address(True, False) & "&""'!""&""$A$4:$K$4""), 0), FALSE) < 20, "
                            Temp = Temp & ActiveWorkbook.Worksheets(index).Cells(i + K, 2).Address(False, True) & ", " & ActiveWorkbook.Worksheets(index).Cells(i + K, 2).Address(False, True) & "*2))))"
                            ActiveWorkbook.Worksheets(index).Cells(i + K, j + 2).Formula = Temp
                            Temp = "=IF(" & ActiveWorkbook.Worksheets(index).Cells(i + K, j + 2).Address(False, False) & ">" & ActiveWorkbook.Worksheets(index).Cells(i + K, 2).Address(False, True)
                            Temp = Temp & "+" & ActiveWorkbook.Worksheets(index).Cells(i + K, j + 1).Address(False, False) & ", 0, "
                            Temp = Temp & ActiveWorkbook.Worksheets(index).Cells(i + K, j + 1).Address(False, False) & " - (" & ActiveWorkbook.Worksheets(index).Cells(i + K, j + 2).Address(False, False)
                            Temp = Temp & "-" & ActiveWorkbook.Worksheets(index).Cells(i + K, 2).Address(False, True) & "))"
                            ActiveWorkbook.Worksheets(index).Cells(i + K, j + 3).Formula = Temp
                        End If
                    Next K
                    Exit For
                End If
            Next j
        End If
    Next i
End Function

Public Function NextMonday() As Date
    Dim D As Integer
    Dim N As Date
    Dim regEx As New RegExp
    Dim i As Integer: i = ActiveWorkbook.Worksheets.Count
    Dim A As Integer
    For A = 1 To i
        Debug.Print (ActiveWorkbook.Worksheets(A).Name)
    Next A
    With regEx
        .Global = True
        .MultiLine = False
        .IgnoreCase = False
        .Pattern = "^(?:1[0-2]|0?[1-9])-(?:0?[1-9]|[12]\d|3[01])-(?:\d{2}|\d{4})$"
    End With
    Dim LastDateString As String
    For A = i To 1 Step -1
        If regEx.Test(ActiveWorkbook.Worksheets(A).Name) Then
            LastDateString = ActiveWorkbook.Worksheets(A).Name
            Exit For
        End If
    Next
    LastDateString = Replace(LastDateString, "-", "/")
    Dim LastDate As Date: LastDate = Format(CDate(LastDateString), "m,d,yy")
    D = Weekday(LastDate)
    N = LastDate + (8 - D)
    NextMonday = N
End Function

Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function

Public Sub EnterEmployees()
    Dim totalCount As Integer: totalCount = 0
    Dim CurrentRowDestination As Integer: CurrentRowDestination = 5
    Dim CurrentRowSource As Integer: CurrentRowSource = 2
    Do While ActiveWorkbook.Worksheets("Employees").Range("$A$" & CurrentRowSource) <> ""
        If ActiveWorkbook.Worksheets("Employees").Range("$C$" & CurrentRowSource).Value = "1" Then
            totalCount = totalCount + 1
        End If
        CurrentRowSource = CurrentRowSource + 1
    Loop
    CurrentRowSource = 2
    Dim CopyOfTotalCount As Integer: CopyOfTotalCount = totalCount
    Do While ActiveWorkbook.Worksheets("Employees").Range("$A$" & CurrentRowSource) <> ""
        If ActiveWorkbook.Worksheets("Employees").Range("$C$" & CurrentRowSource).Value = "1" Then
            ActiveSheet.Range("$A$" & CurrentRowDestination).Value = ActiveWorkbook.Worksheets("Employees").Range("$A$" & CurrentRowSource)
            ActiveSheet.Range("$B$" & CurrentRowDestination).Value = ActiveWorkbook.Worksheets("Employees").Range("$B$" & CurrentRowSource)
            CurrentRowDestination = CurrentRowDestination + 1
            totalCount = totalCount - 1
            If totalCount <> 0 Then
                CopyLastRow CurrentRowDestination
            End If
        End If
        CurrentRowSource = CurrentRowSource + 1
    Loop
    CurrentRowSource = 5
    CurrentRowDestination = 4
    VerifyCorrectRows (CopyOfTotalCount)
    For x = 0 To CopyOfTotalCount - 1
        ActiveWorkbook.Worksheets("Payroll").Cells(CurrentRowDestination, 3).Value = ActiveSheet.Cells(CurrentRowSource, 1).Value
        CurrentRowSource = CurrentRowSource + 1
        CurrentRowDestination = CurrentRowDestination + 1
    Next
End Sub

Public Sub VerifyCorrectRows(totalCount As Integer)
    Dim x As Integer: x = 4
    Dim RowCount As Integer: RowCount = 0
    Do While ActiveWorkbook.Worksheets("Payroll").Cells(x, 1) <> "Totals" And ActiveWorkbook.Worksheets("Payroll").Cells(x, 1) <> ""
        RowCount = RowCount + 1
        x = x + 1
    Loop
    If RowCount > totalCount Then
        x = 4 + totalCount
        Do While ActiveWorkbook.Worksheets("Payroll").Cells(x, 1) <> "Totals"
            ActiveWorkbook.Worksheets("Payroll").Rows(x).Delete
        Loop
    ElseIf RowCount < totalCount Then
        x = 4 + RowCount
        Do While x < 4 + totalCount
            ActiveWorkbook.Worksheets("Payroll").Rows(x).EntireRow.Insert Shift:=slUp, CopyOrigin:=xlFormatFromRightOrAbove
            ActiveWorkbook.Worksheets("Payroll").Rows(x - 1).EntireRow.Copy
            ActiveWorkbook.Worksheets("Payroll").Rows(x).EntireRow.PasteSpecial xlPasteFormats
            ActiveWorkbook.Worksheets("Payroll").Rows(x).EntireRow.PasteSpecial xlPasteFormulas
            x = x + 1
        Loop
    End If
    Call EnterPayrollTotals(5 + totalCount)
End Sub

Public Sub CopyLastRow(CurrentRow)
    ActiveSheet.Rows(CurrentRow).EntireRow.Insert Shift:=slUp, CopyOrigin:=xlFormatFromRightOrAbove
    ActiveSheet.Rows(CurrentRow - 1).EntireRow.Copy
    ActiveSheet.Rows(CurrentRow).EntireRow.PasteSpecial xlPasteFormats
    ActiveSheet.Rows(CurrentRow).EntireRow.PasteSpecial xlPasteFormulas
    Application.CutCopyMode = False
End Sub

Public Sub EnterPayrollTotals(RowNumber As Integer)
    Dim Columns() As Variant: Columns = Array(10, 11, 17, 18, 19, 20, 24)
    Dim Column As Integer
    Dim Temp As String
    For Column = 0 To 6
        Temp = "=SUM(" & ActiveWorkbook.Worksheets("Payroll").Range(Cells(4, Columns(Column)).Address, Cells(RowNumber - 2, Columns(Column)).Address).Address & ")"
        ActiveWorkbook.Worksheets("Payroll").Cells(RowNumber - 1, Columns(Column)).Formula = Temp
    Next
    ActiveWorkbook.Worksheets("Payroll").Rows(RowNumber - 2).EntireRow.Copy
    ActiveWorkbook.Worksheets("Payroll").Rows(RowNumber - 1).EntireRow.PasteSpecial xlPasteFormats
    ActiveWorkbook.Worksheets("Payroll").Cells(RowNumber - 1, 1).Value = "Totals"
End Sub



