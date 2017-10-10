Public Function GetHoursForWeek(Name As String, WeekEnding As String) As Double
    Dim BeginIncludesDate As String: BeginIncludesDate = GetWeekStartDate(WeekEnding)
    Dim BeginIndex As Long: BeginIndex = BinDateSearch(BeginIncludesDate, "HoursDB", 5, False)
    Dim EndIndex As Long: EndIndex = BinDateSearch(WeekEnding, "HoursDB", 5, True)
    Dim x As Long
    Dim Hours As Double: Hours = 0
    For x = BeginIndex To EndIndex
        If ActiveWorkbook.Worksheets("HoursDB").Cells(x, 1).Value = GetQBName(Name) Then
            Hours = Hours + ActiveWorkbook.Worksheets("HoursDB").Cells(x, 3).Value
        End If
    Next
    GetHoursForWeek = Hours
End Function

Public Function GetQBName(Name As String) As String
    Dim x As Integer
    For x = 2 To ActiveWorkbook.Worksheets("Employees").UsedRange.Rows.Count
        If ActiveWorkbook.Worksheets("Employees").Cells(x, 1).Value = Name Then
            GetQBName = ActiveWorkbook.Worksheets("Employees").Cells(x, 5).Value
            Exit Function
        End If
    Next
End Function

Public Function CalculateBankHours(Row As String, sheet As String, PayString As String, TotalHours As Double) As Double
    Dim CurrentBankHours As Double: CurrentBankHours = GetCurrentBankHours(Row, sheet)
    If IsBank(PayString) Then
        Dim PayEquivalentHours As Double: PayEquivalentHours = GetEquivalentHours(PayString)
        If Row = "Barber, Dave" Then
            If TotalHours > PayEquivalentHours Then
                CalculateBankHours = TotalHours - PayEquivalentHours
            ElseIf TotalHours = PayEquivalentHours Then
                CalculateBankHours = 0
            Else
                If CurrentBankHours < PayEquivalentHours - TotalHours Then
                    CalculateBankHours = -1 * CurrentBankHours
                Else
                    CalculateBankHours = TotalHours - PayEquivalentHours
                End If
            End If
            Exit Function
        End If
        If TotalHours < PayEquivalentHours Then
            Dim NeededHours As Double: NeededHours = CalcNeededHours(TotalHours, PayEquivalentHours)
            If NeededHours > CurrentBankHours Then
                CalculateBankHours = -1 * CurrentBankHours
                Exit Function
            Else
                CalculateBankHours = -1 * NeededHours
                Exit Function
            End If
        ElseIf TotalHours = PayEquivalentHours Then
            CalculateBankHours = 0
        Else
            Dim ExcessHours As Double: ExcessHours = CalcExcessHours(TotalHours, PayEquivalentHours)
            CalculateBankHours = ExcessHours
            Exit Function
        End If
    Else
        CalculateBankHours = 0
    End If
End Function

Public Function CalcExcessHours(TotalHours, PayEquivalentHours) As Double
    Dim RGtotal As Double: RGtotal = RG(TotalHours)
    Dim RGEquivalent As Double: RGEquivalent = RG(PayEquivalentHours)
    CalcExcessHours = RGtotal - RGEquivalent
End Function

Public Function CalcNeededHours(TotalHours, PayEquivalentHours) As Double
    Dim RGNeeded As Double
    Dim RGtotal As Double: RGtotal = RG(TotalHours)
    Dim RGEquivalent As Double: RGEquivalent = RG(PayEquivalentHours)
    RGNeeded = RGEquivalent - RGtotal
    CalcNeededHours = RGNeeded
End Function

Public Function RG(Hours) As Double
    Dim Returned As Double
    If Hours > 40 Then
        Returned = Hours + (Hours - 40) * 0.5
    Else
        Returned = Hours
    End If
    RG = Returned
End Function


Public Function GetCurrentBankHours(Row, sheet) As Double
    Dim x As Integer: x = 1
    Dim y As Integer: y = 4
    Dim LastNameRow As Integer: LastNameRow = 1
    For x = 1 To ActiveWorkbook.Worksheets("Bank").UsedRange.Rows.Count
        If ActiveWorkbook.Worksheets("Bank").Cells(x, 1).Value = "Name" Then
            LastNameRow = x
        ElseIf ActiveWorkbook.Worksheets("Bank").Cells(x, 1).Value = Row Then
            Do While ActiveWorkbook.Worksheets("Bank").Cells(LastNameRow, y).Value <> ""
                If ActiveWorkbook.Worksheets("Bank").Cells(LastNameRow, y).Value = sheet Then
                    GetCurrentBankHours = ActiveWorkbook.Worksheets("Bank").Cells(x, y - 1).Value
                    Exit Function
                End If
                y = y + 2
            Loop
        End If
    Next
End Function

Public Function GetEquivalentHours(PayString) As Double
    Dim Returned As String: Returned = ""
    Dim x As Integer: x = 1
    For x = 1 To Len(PayString)
        If IsNumeric(Mid(PayString, x, 1)) Or Mid(PayString, x, 1) = "." Then
            Returned = Returned & Mid(PayString, x, 1)
        End If
    Next
    GetEquivalentHours = Returned
End Function

Public Function IsBank(PayString As String) As Boolean
    If InStr(PayString, "Bank") <> 0 Then
        IsBank = True
    Else
        If InStr(PayString, "bank") <> 0 Then
            IsBank = True
        Else
            IsBank = False
        End If
    End If
End Function


Public Function CalculateInsuranceCost(Row, Column, sheet) As Double
    Dim Balance As Double: Balance = GetHealthValue(Row, Column, sheet, -1)
    Dim WeeklyDeduction As Double: WeeklyDeduction = GetLastHealthValue(Row, "Weekly Deduction", sheet)
    Dim MaxCost As Double: MaxCost = GetHealthValue(Row, Column, sheet)
    If (Balance + WeeklyDeduction) >= MaxCost Then
        CalculateInsuranceCost = MaxCost
        Exit Function
    ElseIf Balance > 0 Then
        CalculateInsuranceCost = Balance + WeeklyDeduction
        Exit Function
    ElseIf MaxCost > WeeklyDeduction Then
        CalculateInsuranceCost = WeeklyDeduction
        Exit Function
    End If
End Function

Public Function GetLastHealthValue(Row, Column, sheet, Optional ByVal offset As Integer) As Double
    Dim i As Integer
    For i = 1 To Application.Caller.Parent.Parent.Worksheets.Count
        If Application.Caller.Parent.Parent.Worksheets(i).Name = sheet Then
            Exit For
        End If
    Next i
    Dim TableHead As Integer: TableHead = 1
    Dim killSwitch As Integer: killSwitch = 0
    Dim y As Integer: y = 1
    Dim x As Integer: x = 1
    Do While killSwitch < 10
        If Application.Caller.Parent.Parent.Worksheets(i).Cells(TableHead, 1).Value = "Name" Then
            killSwitch = 0
            x = TableHead
            Do While Application.Caller.Parent.Parent.Worksheets(i).Cells(x, 1).Value <> ""
                If Application.Caller.Parent.Parent.Worksheets(i).Cells(x, 1).Value = Row Then
                    Do While Application.Caller.Parent.Parent.Worksheets(i).Cells(TableHead, y).Value <> ""
                        If Application.Caller.Parent.Parent.Worksheets(i).Cells(TableHead, y).Value = Column Then
                            GetLastHealthValue = Application.Caller.Parent.Parent.Worksheets(i).Cells(x, y + offset).Value
                        End If
                        y = y + 1
                    Loop
                    y = 1
                End If
                x = x + 1
            Loop
            TableHead = x
        Else
            killSwitch = killSwitch + 1
            TableHead = TableHead + 1
        End If
    Loop
End Function

Public Function GetHealthValue(Row, Column, sheet, Optional ByVal offset As Integer) As Double
    Dim i As Integer
    For i = 1 To Application.Caller.Parent.Parent.Worksheets.Count
        If Application.Caller.Parent.Parent.Worksheets(i).Name = sheet Then
            Exit For
        End If
    Next i
    Dim TableHead As Integer: TableHead = 1
    Dim killSwitch As Integer: killSwitch = 0
    Dim y As Integer: y = 1
    Dim x As Integer: x = 1
    Do While killSwitch < 10
        If Application.Caller.Parent.Parent.Worksheets(i).Cells(TableHead, 1).Value = "Name" Then
            killSwitch = 0
            x = TableHead
            Do While Application.Caller.Parent.Parent.Worksheets(i).Cells(x, 1).Value <> ""
                If Application.Caller.Parent.Parent.Worksheets(i).Cells(x, 1).Value = Row Then
                    Do While Application.Caller.Parent.Parent.Worksheets(i).Cells(TableHead, y).Value <> ""
                        If Application.Caller.Parent.Parent.Worksheets(i).Cells(TableHead, y).Value = Column Then
                            GetHealthValue = Application.Caller.Parent.Parent.Worksheets(i).Cells(x, y + offset).Value
                            Exit Function
                        End If
                        y = y + 1
                    Loop
                    y = 1
                End If
                x = x + 1
            Loop
            TableHead = x
        Else
            TableHead = TableHead + 1
            killSwitch = killSwitch + 1
        End If
    Loop
End Function


