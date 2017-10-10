Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim ShowHide As String: ShowHide = ""
    If Target.Address = "$A$1" Then
        Call BuildMetadata3
        Call RebuildComboBox
        MsgBox "Rebuild Complete"
    ElseIf Not Application.Intersect(Range(Target.Address), Range("D3:G3")) Is Nothing Then
        ActiveWorkbook.Worksheets("JobReporting").Cells(1, 19).Value = Target.Value
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range
    Dim jobname As String
    Set KeyCells = Range("D3:G3")
    Dim oval As String: oval = ActiveWorkbook.Worksheets("JobReporting").Cells(1, 19).Value
    If Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then
        If Target.Address = "$D$3" Then
            jobname = GetJobName(ArgA:=oval)
        ElseIf Target.Address = "$E$3" Then
            jobname = GetJobName(ArgB:=oval)
        ElseIf Target.Address = "$F$3" Then
            jobname = GetJobName(ArgC:=oval)
        Else
            jobname = GetJobName(ArgD:=oval)
        End If
        Call BuildMetadata
        Call RebuildComboBox
        If jobname <> "" Then
            Call SetJob(jobname)
        End If
        Call myCombo_Change
    End If
End Sub



