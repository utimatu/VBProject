Public Function NewWorkbook(jobname As String) As Workbook
    Dim CurrentWorkbook As Workbook: Set CurrentWorkbook = ActiveWorkbook
    Dim NewW As Workbook
    Workbooks.Add
    ActiveWorkbook.SaveAs "S:\Work\Jobs\Reports\" & jobname & " Materials and Hours Report.xlsx"
    Set NewW = ActiveWorkbook
    Dim ws As Worksheet
    Set ws = NewW.Worksheets.Add
    ws.Name = "Report"
    ws.Move Before:=Sheets(1)
    CurrentWorkbook.Activate
    Set NewWorkbook = NewW
End Function

Public Sub ExportJob(jobname As String)
    Dim ExportBook As Workbook: ExportBook = NewWorkbook(jobname)
    Call CreateTitle(ExportBook, jobname)
    Dim Bills() As Integer: ReDim Bills(GetBillCount(jobname))
    Dim x As Integer
    Dim y As Integer: y = 0
    For x = 1 To ActiveWorkbook.Worksheets("BillDB").UsedRange.Rows.Count
        If ActiveWorkbook.Worksheets("BillDB").Cells(x, 6).Value = jobname Then
            Bills(y) = x
            y = y + 1
        End If
    Next
    Dim SubHeaderTier As Integer: SubHeaderTier = 0
    Dim ParentTree(5) As String
    Dim ParentTreeIndex As Integer
    Dim Accounts() As String: Accounts = GetAccounts(Bills)
    Dim WriteLine As Integer: WriteLine = 2
    For x = 0 To UBound(Accounts)
        ParentTreeIndex = GetParentLength(ParentTree, Accounts(x))
        ParentTree = GetParentTree(ParentTree, Accounts(x))
        Call CreateSubHeader(ExportBook, Accounts(x), ParentTreeIndex, WriteLine)
        WriteLine = WriteLine + 1
        For y = 0 To UBound(Bills)
            If ActiveWorkbook.Worksheets("BillDB").Cells(Bills(y), 2).Value = Accounts(x) Then
                ExportBook.Worksheets("Report").Cells(WriteLine, 2).Value = ActiveWorkbook.Worksheets("BillDB").Cells(Bills(y), 1).Value
                ExportBook.Worksheets("Report").Cells(WriteLine, 3).Value = ActiveWorkbook.Worksheets("BillDB").Cells(Bills(y), 4).Value
                ExportBook.Worksheets("Report").Cells(WriteLine, 4).Value = ActiveWorkbook.Worksheets("BillDB").Cells(Bills(y), 5).Value
                ExportBook.Worksheets("Report").Cells(WriteLine, 5).Value = ActiveWorkbook.Worksheets("BillDB").Cells(Bills(y), 3).Value
                WriteLine = WriteLine + 1
            End If
        Next
        Call EndSubHeader(ExportBook, Accounts(x), ParentTreeIndex, WriteLine)
        WriteLine = WriteLine + 1
    Next
End Sub



Public Sub CreateSubHeader(ExportBook As Workbook, SubHeaderName As String, Indent As Integer, LineNumber As Integer)

End Sub

Public Function GetParentLength(ParentTree, Account) As Integer
    IsParentOf
End Function

Public Function GetParentTree(ParentTree, Account) As String()

End Function

Public Sub CreateTitle(ExportBook As Workbook, jobname As String)
    ExportBook.Worksheets("Report").Cells(1, 2).Value = "Job Costing Information of " & MakeJobName(jobname) & " for " & GetContractorName(jobname)
End Sub

Public Function MakeJobName(jobname As String) As String
    Dim List() As String: List = Split(jobname, ":")
    MakeJobName = List(UBound(List))
End Function

Public Function GetContractorName(jobname As String) As String
    Dim List() As String: List = Split(jobname, ":")
    GetContractorName = List(0)
End Function


Public Function GetBillCount(jobname As String) As Integer
    Dim BillCount As Integer
    Dim x As Integer
    For x = 1 To ActiveWorkbook.Worksheets("BillDB").UsedRange.Rows.Count
        If ActiveWorkbook.Worksheets("BillDB").Cells(x, 6).Value = jobname Then
            BillCount = BillCount + 1
        End If
    Next
End Function

Public Function GetAccounts(Bills() As String) As String()
    Dim Accounts(50) As String
    Dim AccountsIndex As Integer: AccountsIndex = 0
    Dim AccountsSize As Integer: AccountsSize = 50
    Dim x As Integer
    For x = 0 To UBound(Bills)
        If IsNotIn(ActiveWorkbook.Worksheets("BillDB").Cells(Bills(x), 2).Value, Accounts, AccountsIndex) Then
            If AccountsSize = AccountsIndex Then
                ReDim Preserve Accounts(AccountsSize * 2)
                AccountsSize = AccountsSize * 2
            End If
            Accounts(AccountsIndex) = ActiveWorkbook.Worksheets("BillDB").Cells(Bills(x), 2).Value
            AccountsIndex = AccountsIndex + 1
        End If
    Next
    Dim Returned(AccountsIndex) As String
    For x = 0 To AccountsIndex
        Returned(x) = Accounts(x)
    Next
    Returned = SortStringList(Returned)
    GetAccounts = Returned
End Function

Public Function IsNotIn(Value As String, List() As String, Length As Integer) As Boolean
    Dim x As Integer
    For x = 0 To Length
        If List(x) = Value Then
            IsNotIn = False
            Exit Function
        End If
    Next
    IsNotIn = True
End Function

Public Function SortStringList(List() As String) As String()
    Dim TempString As String
    Dim x As Integer
    Dim y As Integer
    For x = 0 To UBound(List) - 1
        For y = 0 To UBound(List) - 1
            If StrComp(List(y), List(y + 1), vbTextCompare) = 1 Then
                TempString = List(y)
                List(y) = List(y + 1)
                List(y + 1) = TempString
            End If
        Next
    Next
    SortStringList = List
End Function

