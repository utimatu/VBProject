Private Sub Worksheet_SelectionChange(ByVal Target As Range)
 Dim ShowHide As String: ShowHide = ""
 Dim Month As String: Month = 0
 Dim Year As String: Year = 0
If Target.Address = "$A$1" Then
    Call UnhideAllSheets
ElseIf Target.Address = "$A$2" Then
    Call CopyLastSheet
ElseIf Target.Address = "$C$1" Then
    Call BuildMetadata
End If
If Target.Cells(1).Value Like "Show*" Then
    ShowHide = "Show"
ElseIf Target.Cells(1).Value Like "Hide*" Then
    ShowHide = "Hide"
ElseIf Target.Cells(1).Value Like "Freeze*" Then
    ShowHide = "Freeze"
End If
If Target.Cells(1).Value Like "* 1-*" Then
    Month = "1"
ElseIf Target.Cells(1).Value Like "* 2-*" Then
    Month = "2"
ElseIf Target.Cells(1).Value Like "* 3-*" Then
    Month = "3"
ElseIf Target.Cells(1).Value Like "* 4-*" Then
    Month = "4"
ElseIf Target.Cells(1).Value Like "* 5-*" Then
    Month = "5"
ElseIf Target.Cells(1).Value Like "* 6-*" Then
    Month = "6"
ElseIf Target.Cells(1).Value Like "* 7-*" Then
    Month = "7"
ElseIf Target.Cells(1).Value Like "* 8-*" Then
    Month = "8"
ElseIf Target.Cells(1).Value Like "* 9-*" Then
    Month = "9"
ElseIf Target.Cells(1).Value Like "* 10-*" Then
    Month = "10"
ElseIf Target.Cells(1).Value Like "* 11-*" Then
    Month = "11"
ElseIf Target.Cells(1).Value Like "* 12-*" Then
    Month = "12"
End If
Year = Right(Target.Cells(1).Value, 2)

If ShowHide <> "" Then
    HideShow Year, Month, ShowHide
End If
End Sub

Function HideShow(Column As String, Row As String, hide As String)
    Dim strPattern As String: strPattern = "^[" & Mid(Row, 1, 1) & "]["
    If Len(Row) = 2 Then
        strPattern = strPattern & Mid(Row, 2, 1) & "]["
    End If
    strPattern = strPattern & "-][\d][\d]?[-][" & Mid(Column, 1, 1) & "]"
    If Len(Column) = 2 Then
        strPattern = strPattern & "[" & Mid(Column, 2, 1) & "]"
    End If
    Dim ws_count As Integer: ws_count = ActiveWorkbook.Worksheets.Count
    Dim i As Integer
    Dim regEx As New RegExp
    With regEx
        .Global = False
        .MultiLine = False
        .IgnoreCase = False
        .Pattern = strPattern
    End With
    For i = 1 To ws_count
        If regEx.Test(ActiveWorkbook.Worksheets(i).Name) Then
            If hide = "Hide" Then
                ActiveWorkbook.Worksheets(i).Visible = xlSheetHidden
            ElseIf hide = "Show" Then
                ActiveWorkbook.Worksheets(i).Visible = xlSheetVisible
            ElseIf hide = "Freeze" Then
                ActiveWorkbook.Worksheets(i).UsedRange.Value = ActiveWorkbook.Worksheets(i).UsedRange.Value
            End If
        End If
    Next i
End Function


