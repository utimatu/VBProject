Public Function PrintToFile(s As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile("S:\Work\Payroll\Output.txt")
    oFile.WriteLine (s)
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
End Function

Public Sub WriteToFile(s As String, f As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile(f)
    oFile.WriteLine (s)
    Set fso = Nothing
    Set oFile = Nothing
End Sub

Public Sub Wrapper()
    Dim Report As Workbook: Set Report = NewWorkbook("Report")
    Report.Worksheets("Sheet1").Cells(1, 1).Value = "Test"
    
End Sub
Public Function BinarySearch(search As String, s() As String) As Integer
    Dim Upper As Integer: Upper = UBound(s)
    Dim Lower As Integer: Lower = 0
    Dim Mid As Integer
    If StrComp(search, s(Upper), vbTextCompare) = 0 Then
        BinarySearch = Upper
        Exit Function
    ElseIf StrComp(search, s(Lower), vbTextCompare) = 0 Then
        BinarySearch = Lower
        Exit Function
    End If
    Do While (Upper - Lower) > 1
        Mid = (Upper + Lower) / 2
        If StrComp(search, s(Mid), vbTextCompare) = 1 Then
            Lower = Mid
        ElseIf StrComp(search, s(Mid), vbTextCompare) = -1 Then
            Upper = Mid
        Else
            BinarySearch = Mid
            Exit Function
        End If
    Loop
    BinarySearch = -1
End Function


Public Function GetCode() As String
    Dim x As Integer
    Dim Code As String
    For x = 1 To ActiveWorkbook.VBProject.VBComponents.Count
        If ActiveWorkbook.VBProject.VBComponents(x).Name = "Module1" Then
            Code = ActiveWorkbook.VBProject.VBComponents(x).CodeModule.Lines(1, ActiveWorkbook.VBProject.VBComponents(x).CodeModule.CountOfLines)
        End If
    Next
    MsgBox (ActiveWorkbook.Worksheets("6-25-17").CodeName)
    
    MsgBox Code
    GetCode = Code
    Exit Function
End Function

Public Sub ExportCode()
    Dim x As Integer
    Dim y As Integer
    Dim Code As String
    Dim Found As Boolean: Found = False
    For x = 1 To ActiveWorkbook.VBProject.VBComponents.Count
        If ActiveWorkbook.VBProject.VBComponents(x).CodeModule.CountOfLines > 0 Then
            For y = 1 To ActiveWorkbook.Worksheets.Count
                If ActiveWorkbook.Worksheets(y).CodeName = ActiveWorkbook.VBProject.VBComponents(x).Name Then
                    Found = True
                    If Not IsNumeric(Mid(ActiveWorkbook.Worksheets(y).Name, 1, 1)) Then
                        Call WriteToFile(ActiveWorkbook.VBProject.VBComponents(x).CodeModule.Lines(1, ActiveWorkbook.VBProject.VBComponents(x).CodeModule.CountOfLines + 1), "C:\Users\Romano Masonry\Dropbox\Projects\VBProject\" & ActiveWorkbook.Worksheets(y).Name & ".vb")
                    End If
                End If
            Next
            If Not Found Then
                Call WriteToFile(ActiveWorkbook.VBProject.VBComponents(x).CodeModule.Lines(1, ActiveWorkbook.VBProject.VBComponents(x).CodeModule.CountOfLines + 1), "C:\Users\Romano Masonry\Dropbox\Projects\VBProject\" & ActiveWorkbook.VBProject.VBComponents(x).Name & ".vb")
            End If
            Found = False
        End If
    Next
End Sub


Public Sub Temp()
    Dim x As Integer
    For x = 1 To ActiveWorkbook.VBProject.VBComponents.Count
        If ActiveWorkbook.VBProject.VBComponents(x).Name = "Sheet621" Then
            Call ActiveWorkbook.VBProject.VBComponents.Remove(ActiveWorkbook.VBProject.VBComponents(x))
        End If
    Next
End Sub

