Sub CalcBonuses()
    Dim Names() As String
    ReDim Preserve Names(100)
    Dim Costs() As Double
    ReDim Preserve Costs(100)
    Dim Start As Long
    Dim Finish As Long
    Dim Entered As Boolean: Entered = False
    Dim ShouldBeEntered As Boolean: ShouldBeEntered = True
    Dim BeginDate As String
    Dim EndDate As String
    Dim TempString As String
    Call BuildNames
    Call ValidateJobs
    For x = 2 To ActiveWorkbook.Worksheets("Bonuses").UsedRange.Columns.Count
        If ActiveWorkbook.Worksheets("Bonuses").Cells(1, x).Value = "" Then
            Exit For
        End If
        Dim Jobs() As String
        Jobs = Split(ActiveWorkbook.Worksheets("Bonuses").Cells(1, x).Value, "||")
        For Job = 0 To UBound(Jobs)
            BeginDate = ""
            EndDate = ""
            If InStr(Jobs(Job), "(") > 0 Then
                TempString = Split(Jobs(Job), "(")(1)
                TempString = Split(TempString, ")")(0)
                BeginDate = Split(TempString, "|")(0)
                EndDate = Split(TempString, "|")(1)
                Jobs(Job) = Split(Jobs(Job), "(")(0)
            End If
            For y = 1 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count
                If ActiveWorkbook.Worksheets("MetadataDB").Cells(y, 1).Value = Jobs(Job) Then
                    Start = ActiveWorkbook.Worksheets("MetadataDB").Cells(y, 2).Value
                    If BeginDate <> "" Then
                        Do While True
                            If DateCompare2(BeginDate, ActiveWorkbook.Worksheets("HoursDB").Cells(Start, 5).Value) Then
                                Exit Do
                            Else
                                Start = Start + 1
                            End If
                        Loop
                    End If
                    Finish = ActiveWorkbook.Worksheets("MetadataDB").Cells(y, 3).Value
                    If EndDate <> "" Then
                        Do While True
                            If DateCompare2(ActiveWorkbook.Worksheets("HoursDB").Cells(Finish, 5).Value, EndDate) Then
                                Exit Do
                            Else
                                Finish = Finish - 1
                            End If
                        Loop
                    End If
                    Exit For
                End If
            Next
            For y = Start - 1 To Finish + 1
                If ActiveWorkbook.Worksheets("HoursDB").Cells(y, 8).Value = Jobs(Job) Then
                    For z = 2 To ActiveWorkbook.Worksheets("Employees").UsedRange.Rows.Count
                        If ActiveWorkbook.Worksheets("Employees").Cells(z, 5) = ActiveWorkbook.Worksheets("HoursDB").Cells(y, 1).Value Then
                            If ActiveWorkbook.Worksheets("Employees").Cells(z, 3).Value <> "0" Then
                                Dim index As Integer: index = -1
                                For i = 0 To UBound(Names)
                                    If Names(i) = ActiveWorkbook.Worksheets("HoursDB").Cells(y, 1).Value Then
                                        index = i
                                        Exit For
                                    End If
                                Next
                                If index = -1 Then
                                    For i = 0 To UBound(Names)
                                        If Names(i) = "" Then
                                            Names(i) = ActiveWorkbook.Worksheets("HoursDB").Cells(y, 1).Value
                                            index = i
                                            Dim TempName As String
                                            Dim TempCost As Double
                                            For j = i To 1 Step -1
                                                If StrComp(Names(j - 1), Names(j), vbTextCompare) = 1 Then
                                                    TempName = Names(j - 1)
                                                    Names(j - 1) = Names(j)
                                                    Names(j) = TempName
                                                    TempCost = Costs(j - 1)
                                                    Costs(j - 1) = Costs(j)
                                                    Costs(j) = TempCost
                                                    index = j - 1
                                                Else
                                                    index = j
                                                    Exit For
                                                End If
                                            Next
                                            Exit For
                                        End If
                                    Next
                                End If
                                Dim Wage As Double: Wage = DetermineWage(ActiveWorkbook.Worksheets("HoursDB").Cells(y, 1), ActiveWorkbook.Worksheets("HoursDB").Cells(y, 5).Value)
                                If IsOvertime(ActiveWorkbook.Worksheets("HoursDB").Cells(y, 4).Value) Then
                                    Wage = Wage * 1.5
                                End If
                                Costs(index) = Costs(index) + Wage * ActiveWorkbook.Worksheets("HoursDB").Cells(y, 3).Value
                            End If
                        End If
                    Next
                End If
            Next
        Next
        Dim TotalCost As Double: TotalCost = 0
        For y = 0 To UBound(Costs)
            TotalCost = TotalCost + Costs(y)
        Next
        Dim Ind As Integer
        For y = 0 To UBound(Costs)
            Dim K As Integer: K = 4
            Do While ActiveWorkbook.Worksheets("Bonuses").Cells(K, 1).Value <> ""
                If ActiveWorkbook.Worksheets("Bonuses").Cells(K, 1).Value = Names(y) Then
                    ActiveWorkbook.Worksheets("Bonuses").Cells(K, x).Value = Round((Costs(y) / TotalCost) * ActiveWorkbook.Worksheets("Bonuses").Cells(3, x).Value * 0.6, 2)
                End If
                K = K + 1
            Loop
        Next
        K = 4
        Do While ActiveWorkbook.Worksheets("Bonuses").Cells(K, 1).Value <> ""
            If ActiveWorkbook.Worksheets("Bonuses").Cells(K, 1).Value = "Barber, David J" Then
                ActiveWorkbook.Worksheets("Bonuses").Cells(K, x).Value = Round(ActiveWorkbook.Worksheets("Bonuses").Cells(3, x).Value * 0.052, 2)
            ElseIf ActiveWorkbook.Worksheets("Bonuses").Cells(K, 1).Value = "Carter, Timothy J" Then
                ActiveWorkbook.Worksheets("Bonuses").Cells(K, x).Value = Round(ActiveWorkbook.Worksheets("Bonuses").Cells(3, x).Value * 0.078, 2)
            ElseIf ActiveWorkbook.Worksheets("Bonuses").Cells(K, 1).Value = "Friedline, Kyle" Then
                ActiveWorkbook.Worksheets("Bonuses").Cells(K, x).Value = Round(ActiveWorkbook.Worksheets("Bonuses").Cells(3, x).Value * 0.008, 2)
            ElseIf ActiveWorkbook.Worksheets("Bonuses").Cells(K, 1).Value = "Houston, Audra" Then
                ActiveWorkbook.Worksheets("Bonuses").Cells(K, x).Value = Round(ActiveWorkbook.Worksheets("Bonuses").Cells(3, x).Value * 0.008, 2)
            ElseIf ActiveWorkbook.Worksheets("Bonuses").Cells(K, 1).Value = "Hull, Curtis R" Then
                ActiveWorkbook.Worksheets("Bonuses").Cells(K, x).Value = Round(ActiveWorkbook.Worksheets("Bonuses").Cells(3, x).Value * 0.004, 2)
            End If
            K = K + 1
        Loop
        For y = 0 To UBound(Costs)
            Costs(y) = 0
            Names(y) = ""
        Next
    Next
End Sub


Sub TimBonuses()
    Dim Names() As String: Names = BuildNamesReturned()
    Dim Wages() As Double: ReDim Wages(UBound(Names))
    Dim Hours() As Double: ReDim Hours(UBound(Names))
    Dim Jobs() As String
    Dim x As Integer
    Dim y As Integer
    Dim z As Long
    Dim i As Long
    Dim j As Integer
    Dim BeginDate As String
    Dim EndDate As String
    Dim TotalBonus As Double: TotalBonus = ActiveWorkbook.Worksheets("Bonuses").Cells(3, 2).Value
    Dim TotalCost As Double
    Call ValidateJobs
    For x = 2 To ActiveWorkbook.Worksheets("Bonuses").UsedRange.Columns.Count
        Jobs = Split(ActiveWorkbook.Worksheets("Bonuses").Cells(1, x).Value, "||")
        For y = 0 To UBound(Jobs)
            BeginDate = ""
            EndDate = ""
            If InStr(Jobs(y), "(") > 0 Then
                TempString = Split(Jobs(y), "(")(1)
                TempString = Split(TempString, ")")(0)
                BeginDate = Split(TempString, "|")(0)
                EndDate = Split(TempString, "|")(1)
                Jobs(y) = Split(Jobs(y), "(")(0)
            End If
            For z = 1 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count
                If ActiveWorkbook.Worksheets("MetadataDB").Cells(z, 1).Value = Jobs(y) Then
                    For i = ActiveWorkbook.Worksheets("MetadataDB").Cells(z, 2).Value To ActiveWorkbook.Worksheets("MetadataDB").Cells(z, 3).Value
                        If ActiveWorkbook.Worksheets("HoursDB").Cells(i, 8).Value = Jobs(y) Then
                            For j = 0 To UBound(Names)
                                If StrComp(Names(j), ActiveWorkbook.Worksheets("HoursDB").Cells(i, 1).Value, vbTextCompare) = 0 Then
                                    If BeginDate <> "" Then
                                        If DateCompare2(ActiveWorkbook.Worksheets("HoursDB").Cells(i, 5).Value, BeginDate) Then
                                            Exit For
                                        End If
                                    End If
                                    If EndDate <> "" Then
                                        If DateCompare2(EndDate, ActiveWorkbook.Worksheets("HoursDB").Cells(i, 5).Value) Then
                                            Exit For
                                        End If
                                    End If
                                    Wages(j) = Wages(j) + DetermineWage(Names(j), ActiveWorkbook.Worksheets("HoursDB").Cells(i, 5).Value) * ActiveWorkbook.Worksheets("HoursDB").Cells(i, 3).Value * OvertimeModifier(ActiveWorkbook.Worksheets("HoursDB").Cells(i, 4).Value)
                                    Hours(j) = Hours(j) + ActiveWorkbook.Worksheets("HoursDB").Cells(i, 3).Value
                                End If
                            Next
                        End If
                    Next
                    Exit For
                End If
            Next
        Next
    Next
    Call BuildNames
    For x = 0 To UBound(Names)
        For y = 4 To ActiveWorkbook.Worksheets("Bonuses").UsedRange.Rows.Count
            If ActiveWorkbook.Worksheets("Bonuses").Cells(y, 1).Value = Names(x) Then
                ActiveWorkbook.Worksheets("Bonuses").Cells(y, 3).Value = Wages(x)
                ActiveWorkbook.Worksheets("Bonuses").Cells(y, 4).Value = Hours(x)
                TotalCost = TotalCost + Wages(x)
            End If
        Next
    Next
    For x = 4 To ActiveWorkbook.Worksheets("Bonuses").UsedRange.Rows.Count
        If ActiveWorkbook.Worksheets("Bonuses").Cells(x, 1).Value = "Barber, David J" Then
            ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value = Round(TotalBonus * 0.052)
        ElseIf ActiveWorkbook.Worksheets("Bonuses").Cells(x, 1).Value = "Carter, Timothy J" Then
            ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value = Round(TotalBonus * 0.078)
        ElseIf ActiveWorkbook.Worksheets("Bonuses").Cells(x, 1).Value = "Friedline, Kyle" Then
            ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value = Round(TotalBonus * 0.008)
        ElseIf ActiveWorkbook.Worksheets("Bonuses").Cells(x, 1).Value = "Houston, Audra" Then
            ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value = Round(TotalBonus * 0.008)
        ElseIf ActiveWorkbook.Worksheets("Bonuses").Cells(x, 1).Value = "Hull, Curtis R" Then
            ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value = Round(TotalBonus * 0.004)
        Else
            For y = 0 To UBound(Names)
                If Names(y) = ActiveWorkbook.Worksheets("Bonuses").Cells(x, 1).Value Then
                    ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value = Round((Wages(y) / TotalCost) * 0.6 * TotalBonus)
                    If ActiveWorkbook.Worksheets("Bonuses").Cells(x, 4).Value > 0 Then
                        ActiveWorkbook.Worksheets("Bonuses").Cells(x, 5).Value = Round(ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value / ActiveWorkbook.Worksheets("Bonuses").Cells(x, 4).Value, 2)
                    End If
                End If
            Next
        End If
    Next
End Sub

Public Sub TimeRangeBonus(Start As String, Ending As String)
    Dim Names() As String: Names = BuildNamesReturned()
    Dim Wages() As Double: ReDim Wages(UBound(Names))
    Dim Hours() As Double: ReDim Hours(UBound(Names))
    Dim x As Long
    Dim y As Integer
    Dim TotalBonus As Double: TotalBonus = ActiveWorkbook.Worksheets("Bonuses").Cells(3, 2).Value
    Dim SpecialNames() As String: SpecialNames = Split("Barber, David J|Carter, Timothy J|Friedline, Kyle|Houston, Audra|Hull, Curtis R", "|")
    For x = BinDateSearch(Start, "HoursDB", 5, False) To BinDateSearch(Ending, "HoursDB", 5, True)
        For y = 0 To UBound(Names)
            If StrComp(ActiveWorkbook.Worksheets("HoursDB").Cells(x, 1).Value, Names(y), vbTextCompare) = 0 Then
                Hours(y) = Hours(y) + ActiveWorkbook.Worksheets("HoursDB").Cells(x, 3).Value
                Wages(y) = Wages(y) + DetermineWage(Names(y), ActiveWorkbook.Worksheets("HoursDB").Cells(x, 5).Value) * ActiveWorkbook.Worksheets("HoursDB").Cells(x, 3).Value * OvertimeModifier(ActiveWorkbook.Worksheets("HoursDB").Cells(x, 4).Value)
                Exit For
            End If
        Next
    Next
    Call BuildNames
    For x = 0 To UBound(Names)
        For y = 4 To ActiveWorkbook.Worksheets("Bonuses").UsedRange.Rows.Count
            If ActiveWorkbook.Worksheets("Bonuses").Cells(y, 1).Value = Names(x) Then
                ActiveWorkbook.Worksheets("Bonuses").Cells(y, 3).Value = Wages(x)
                ActiveWorkbook.Worksheets("Bonuses").Cells(y, 4).Value = Hours(x)
                If IsNotIn(Names(x), SpecialNames, UBound(SpecialNames)) Then
                    TotalCost = TotalCost + Wages(x)
                End If
            End If
        Next
    Next
    For x = 4 To ActiveWorkbook.Worksheets("Bonuses").UsedRange.Rows.Count
        If ActiveWorkbook.Worksheets("Bonuses").Cells(x, 1).Value = "Barber, David J" Then
            ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value = Round(TotalBonus * 0.052)
            ActiveWorkbook.Worksheets("Bonuses").Cells(x, 5).Value = Round(ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value / ActiveWorkbook.Worksheets("Bonuses").Cells(x, 4).Value, 2)
        ElseIf ActiveWorkbook.Worksheets("Bonuses").Cells(x, 1).Value = "Carter, Timothy J" Then
            ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value = Round(TotalBonus * 0.078)
            ActiveWorkbook.Worksheets("Bonuses").Cells(x, 5).Value = Round(ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value / 585, 2)
        ElseIf ActiveWorkbook.Worksheets("Bonuses").Cells(x, 1).Value = "Friedline, Kyle" Then
            ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value = Round(TotalBonus * 0.008)
            ActiveWorkbook.Worksheets("Bonuses").Cells(x, 5).Value = Round(ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value / ActiveWorkbook.Worksheets("Bonuses").Cells(x, 4).Value, 2)
        ElseIf ActiveWorkbook.Worksheets("Bonuses").Cells(x, 1).Value = "Houston, Audra" Then
            ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value = Round(TotalBonus * 0.008)
            ActiveWorkbook.Worksheets("Bonuses").Cells(x, 5).Value = Round(ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value / ActiveWorkbook.Worksheets("Bonuses").Cells(x, 4).Value, 2)
        ElseIf ActiveWorkbook.Worksheets("Bonuses").Cells(x, 1).Value = "Hull, Curtis R" Then
            ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value = Round(TotalBonus * 0.004)
            ActiveWorkbook.Worksheets("Bonuses").Cells(x, 5).Value = Round(ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value / ActiveWorkbook.Worksheets("Bonuses").Cells(x, 4).Value, 2)
        Else
            For y = 0 To UBound(Names)
                If Names(y) = ActiveWorkbook.Worksheets("Bonuses").Cells(x, 1).Value Then
                    ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value = Round((Wages(y) / TotalCost) * 0.6 * TotalBonus)
                    If ActiveWorkbook.Worksheets("Bonuses").Cells(x, 4).Value > 0 Then
                        ActiveWorkbook.Worksheets("Bonuses").Cells(x, 5).Value = Round(ActiveWorkbook.Worksheets("Bonuses").Cells(x, 2).Value / ActiveWorkbook.Worksheets("Bonuses").Cells(x, 4).Value, 2)
                    End If
                End If
            Next
        End If
    Next
End Sub

Public Sub Wrapper()
    Call TimeRangeBonus("7-1-17", "9-30-17")
End Sub

Public Function BuildNamesReturned() As String()
    Dim x As Integer
    Dim Returned() As String: ReDim Returned(50)
    Dim ReturnedPointer As Integer: ReturnedPointer = 0
    For x = 2 To ActiveWorkbook.Worksheets("Employees").UsedRange.Rows.Count
        If ActiveWorkbook.Worksheets("Employees").Cells(x, 3).Value = 1 Then
            If ReturnedPointer = UBound(Returned) Then
                ReDim Preserve Returned(ReturnedPointer * 2)
            End If
            Returned(ReturnedPointer) = ActiveWorkbook.Worksheets("Employees").Cells(x, 5).Value
            ReturnedPointer = ReturnedPointer + 1
        End If
    Next
    ReDim Preserve Returned(ReturnedPointer - 1)
    BuildNamesReturned = Returned
End Function

Sub BuildNames()
    Dim CurentIndex As Integer: CurentIndex = 4
    Dim Found As Boolean: Found = True
    For x = 2 To ActiveWorkbook.Worksheets("Employees").UsedRange.Rows.Count
        If ActiveWorkbook.Worksheets("Employees").Cells(x, 3).Value = 1 Then
            If StrComp(ActiveWorkbook.Worksheets("Employees").Cells(x, 5), "Carter, Timothy J", vbTextCompare) = 1 And Found Then
                ActiveWorkbook.Worksheets("Bonuses").Cells(CurentIndex, 1).Value = "Carter, Timothy J"
                CurentIndex = CurentIndex + 1
                Found = False
            End If
            ActiveWorkbook.Worksheets("Bonuses").Cells(CurentIndex, 1).Value = ActiveWorkbook.Worksheets("Employees").Cells(x, 5).Value
            CurentIndex = CurentIndex + 1
        End If
    Next
End Sub

Sub ValidateJobs()
    Dim x As Integer
    Dim y As Integer
    Dim Found As Boolean: Found = False
    Dim jobname() As String
    For x = 2 To ActiveWorkbook.Worksheets("Bonuses").UsedRange.Columns.Count
        If ActiveWorkbook.Worksheets("Bonuses").Cells(1, x).Value = "" Then
            Exit For
        End If
        For y = 2 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count
            If ActiveWorkbook.Worksheets("MetadataDB").Cells(y, 1).Value = "" Then
                Exit For
            End If
            jobname = Split(ActiveWorkbook.Worksheets("MetadataDB").Cells(y, 1).Value, ":")
            If StrComp(jobname(UBound(jobname)), ActiveWorkbook.Worksheets("Bonuses").Cells(1, x).Value, vbTextCompare) = 0 Then
                If Found Then
                    MsgBox ActiveWorkbook.Worksheets("Bonuses").Cells(1, x).Value
                Else
                    Found = True
                End If
            End If
        Next
        Found = False
    Next
End Sub

