
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Row <> 1 Then
        If Target.Column <> 1 Then
            Call BuildReport(ActiveWorkbook.Worksheets("MainBudget").Cells(1, Target.Column).Value, ActiveWorkbook.Worksheets("MainBudget").Cells(Target.Row, 1).Value)
        Else
            Call BuildReport(Account:=ActiveWorkbook.Worksheets("MainBudget").Cells(Target.Row, 1).Value)
        End If
    Else
        Call BuildReport(ActiveWorkbook.Worksheets("MainBudget").Cells(1, Target.Column).Value)
    End If
End Sub

