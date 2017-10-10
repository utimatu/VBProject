Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim args As String: args = ""
Dim RetVal
On Error GoTo ErrCatcher:
    If Target.Cells(1).Value = "Pay All" Or Target.Cells(1).Value = "Bank>40" Or Target.Cells(1).Value = "Hold" Then
        args = """" & ActiveSheet.Cells(Target.Row, 1).Value & """" & " "
        If ActiveSheet.Cells(Target.Row, 3).Value = "" Then
            args = args & "0 "
        Else
            args = args & ActiveSheet.Cells(Target.Row, 3).Value & " "
        End If
    
        If ActiveSheet.Cells(Target.Row, 4).Value = "" Then
            args = args & "0 "
        Else
            args = args & ActiveSheet.Cells(Target.Row, 4).Value & " "
        End If
    
        If ActiveSheet.Cells(Target.Row, 5).Value = "" Then
            args = args & "0 "
        Else
            args = args & ActiveSheet.Cells(Target.Row, 5).Value & " "
        End If
    
        If ActiveSheet.Cells(Target.Row, 6).Value = "" Then
            args = args & "0 "
        Else
            args = args & ActiveSheet.Cells(Target.Row, 6).Value & " "
        End If
    
        If ActiveSheet.Cells(Target.Row, 7).Value = "" Then
            args = args & "0 "
        Else
            args = args & ActiveSheet.Cells(Target.Row, 7).Value & " "
        End If
    
        If ActiveSheet.Cells(Target.Row, 8).Value = "" Then
            args = args & "0 "
        Else
            args = args & ActiveSheet.Cells(Target.Row, 8).Value & " "
        End If
    
        If ActiveSheet.Cells(Target.Row, 9).Value = "" Then
            args = args & "0 "
        Else
            args = args & ActiveSheet.Cells(Target.Row, 9).Value & " "
        End If
        args = args & ActiveSheet.Cells(2, 3).Value & " "
        Dim i As Integer
        For i = 1 To ActiveWorkbook.Worksheets.Count
          If ActiveWorkbook.Worksheets(i).Name = "Employees" Then
            Dim x As Integer
            For x = 1 To ActiveWorkbook.Worksheets(i).UsedRange.Rows.Count + 1
              If ActiveWorkbook.Worksheets(i).Cells(x, 1).Value = ActiveSheet.Cells(Target.Row, 1) Then
                Dim y As Integer: y = 7
                Dim Wage As Double: Wage = ActiveWorkbook.Worksheets(i).Cells(x, y).Value
                Do While True
                  If ActiveWorkbook.Worksheets(i).Cells(x, y + 1).Value = "" Then
                    Exit Do
                  ElseIf DateCompare2(DateToString(ActiveWorkbook.Worksheets(i).Cells(x, y + 1).Value), ActiveSheet.Name) Then
                    y = y + 2
                    Wage = ActiveWorkbook.Worksheets(i).Cells(x, y).Value
                  Else
                    Exit Do
                  End If
                Loop
                args = args & Wage & " " & """"
                args = args & ActiveWorkbook.Worksheets(i).Cells(x, 4).Value & """"
                Exit For
              End If
            Next x
          End If
        Next i
        Call Shell("C:\Python27\python.exe S:\Work\Payroll\low_slip.py " & args, vbNormalFocus)
    End If
ErrCatcher:
    Exit Sub
End Sub







