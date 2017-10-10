Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  Dim i As Integer
  If Target.Address = "$A$1" Then
    For i = 1 To ActiveWorkbook.Worksheets.Count
      If ActiveWorkbook.Worksheets(i).Name = "Bank" Then
        Call AddColumnToBank(i, "4-16-17", "4-9-17")
      End If
    Next i
  End If
End Sub




