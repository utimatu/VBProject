Public Function GetWeekStartDate(WeekEndDate As String) As String
    Dim InitDate As Date: InitDate = Format(CDate(Replace(WeekEndDate, "-", "/")), "m,d,yy")
    GetWeekStartDate = CStr(InitDate - 6)
End Function

Public Function DateInside(Tested As String, Start As String, Ending As String) As Boolean
    Tested = Replace(Tested, "/", "-")
    Start = Replace(Start, "/", "-")
    Ending = Replace(Ending, "/", "-")
    If Not LooksLikeDate(Tested) Then
        DateInside = False
        Exit Function
    End If
    If Not LooksLikeDate(Start) Then
        DateInside = False
        Exit Function
    End If
    If Not LooksLikeDate(Ending) Then
        DateInside = False
        Exit Function
    End If
    Tested = FormatShortToLong(Tested)
    Start = FormatShortToLong(Start)
    Ending = FormatShortToLong(Ending)
    Dim Tested2() As String: Tested2 = Split(Tested, "-")
    Dim Start2() As String: Start2 = Split(Start, "-")
    Dim Ending2() As String: Ending2 = Split(Ending, "-")
    If DateCompare(Start2, Tested2) And DateCompare(Tested2, Ending2) Then
        DateInside = True
    Else
        DateInside = False
    End If
End Function

Private Function DateCompare(Early() As String, Late() As String)
    If CInt(Late(2)) > CInt(Early(2)) Then
        DateCompare = True
        Exit Function
    ElseIf CInt(Late(2)) = CInt(Early(2)) Then
        If CInt(Late(0)) > CInt(Early(0)) Then
            DateCompare = True
            Exit Function
        ElseIf CInt(Late(0)) = CInt(Early(0)) Then
            If CInt(Late(1)) >= CInt(Early(1)) Then
                DateCompare = True
                Exit Function
            Else
                DateCompare = False
                Exit Function
            End If
        Else
            DateCompare = False
            Exit Function
        End If
    Else
        DateCompare = False
        Exit Function
    End If
End Function

Public Function BinDateSearch(Da As String, Sheetname As String, ColumnIndex As Integer, Ending As Boolean) As Long
    Dim BeginIndex As Long: BeginIndex = 2
    Dim EndIndex As Long: EndIndex = ActiveWorkbook.Worksheets(Sheetname).UsedRange.Rows.Count
    Dim CompareIndex As Long: CompareIndex = BeginIndex + (EndIndex - BeginIndex) / 2
    Do While EndIndex - BeginIndex > 1
        If Ending Then
            If DateCompare2(ActiveWorkbook.Worksheets(Sheetname).Cells(CompareIndex, ColumnIndex).Value, Da) Then
                BeginIndex = CompareIndex
            Else
                EndIndex = CompareIndex
            End If
        Else
            If DateCompare2(Da, ActiveWorkbook.Worksheets(Sheetname).Cells(CompareIndex, ColumnIndex).Value) Then
                EndIndex = CompareIndex
            Else
                BeginIndex = CompareIndex
            End If
        End If
        CompareIndex = BeginIndex + (EndIndex - BeginIndex) / 2
    Loop
    If Ending Then
        If DateCompare2(ActiveWorkbook.Worksheets(Sheetname).Cells(EndIndex, ColumnIndex).Value, Da) Then
            BinDateSearch = EndIndex
        Else
            BinDateSearch = BeginIndex
        End If
    Else
        If DateCompare2(Da, ActiveWorkbook.Worksheets(Sheetname).Cells(BeginIndex, ColumnIndex).Value) Then
            BinDateSearch = BeginIndex
        Else
            BinDateSearch = EndIndex
        End If
    End If
End Function


Public Function DateCompare2(Early As String, Late As String)
    Early = FormatShortToLong(Early)
    Late = FormatShortToLong(Late)
    Dim Early2() As String: Early2 = Split(Early, "-")
    Dim Late2() As String: Late2 = Split(Late, "-")
    DateCompare2 = DateCompare(Early2, Late2)
End Function

Public Function LooksLikeDate(Tested As String) As Boolean
    On Error GoTo ErrCatcher
        Dim Tested2() As String: Tested2 = Split(Tested, "-")
        If CInt(Tested2(0)) < 13 And CInt(Tested2(0)) > 0 Then
            If CInt(Tested2(1)) < 32 And CInt(Tested2(1)) > 0 Then
                If (CInt(Tested2(2)) < 100 And CInt(Tested2(2)) > -1) Or (CInt(Tested2(2)) > 1950 And CInt(Tested2(2)) < 2100) Then
                    LooksLikeDate = True
                    Exit Function
                Else
                    LooksLikeDate = False
                    Exit Function
                End If
            Else
                LooksLikeDate = False
                Exit Function
            End If
        Else
            LooksLikeDate = False
            Exit Function
        End If
ErrCatcher:
    LooksLikeDate = False
End Function

Public Function GetMonthFromDate(D As String) As Integer
    D = Replace(D, "/", "-")
    Dim Da() As String: Da = Split(D, "-")
    GetMonthFromDate = CInt(Da(0))
End Function

Public Function DatesCross(S1 As String, E1 As String, S2 As String, E2 As String) As Boolean
    If DateCompare2(E1, S2) Or DateCompare2(E2, S1) Then
        DatesCross = False
    Else
        DatesCross = True
    End If
End Function

Public Function FormatShortToLong(str As String) As String
    Dim arr() As String
    str = Replace(str, "/", "-")
    arr = Split(str, "-")
    Dim Returned As String
    If CInt(arr(2)) < 100 Then
        Returned = arr(0) & "-" & arr(1) & "-20" & arr(2)
    Else
        Returned = str
    End If
    FormatShortToLong = Returned
End Function

