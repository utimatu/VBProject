Public Function Enumerate(Begin As String, Ending As String, Column1 As Integer, Column2 As Integer) As String()
    Dim JobList() As String: ReDim JobList(50)
    Dim JobListLength As Integer: JobListLength = 50
    Dim JobListIndex As Integer: JobListIndex = 0
    Dim x As Long
    Dim BeginIndex As Long: BeginIndex = BinDateSearch(Begin, "HoursDB", 5, False)
    Dim EndIndex As Long: EndIndex = BinDateSearch(Ending, "HoursDB", 5, True)
    For x = BeginIndex - 1 To EndIndex + 1
        If DateInside(ActiveWorkbook.Worksheets("HoursDB").Cells(x, 5).Value, Begin, Ending) Then
            If IsNotIn(ActiveWorkbook.Worksheets("HoursDB").Cells(x, Column1).Value, JobList, JobListIndex - 1) Then
                If JobListIndex + 1 = JobListLength Then
                    ReDim Preserve JobList(JobListLength * 2)
                    JobListLength = JobListLength * 2
                End If
                JobList(JobListIndex) = ActiveWorkbook.Worksheets("HoursDB").Cells(x, Column1).Value
                JobListIndex = JobListIndex + 1
            End If
        End If
    Next
    BeginIndex = BinDateSearch(Begin, "BillDB", 5, False)
    EndIndex = BinDateSearch(Ending, "BillDB", 5, True)
    
    For x = BeginIndex - 1 To EndIndex + 1
        If DateInside(ActiveWorkbook.Worksheets("BillDB").Cells(x, 5).Value, Begin, Ending) Then
            If IsNotIn(ActiveWorkbook.Worksheets("BillDB").Cells(x, Column2).Value, JobList, JobListIndex) Then
                If JobListIndex + 1 = JobListLength Then
                    ReDim Preserve JobList(JobListLength * 2)
                    JobListLength = JobListLength * 2
                End If
                JobList(JobListIndex) = ActiveWorkbook.Worksheets("BillDB").Cells(x, Column2).Value
                JobListIndex = JobListIndex + 1
            End If
        End If
    Next
    Dim Returned() As String: ReDim Returned(JobListIndex)
    For x = 0 To JobListIndex
        Returned(x) = JobList(x)
    Next
    Returned = SortStringList(Returned)
    Enumerate = Returned
End Function

Public Function GetBudgetFor(Job As String, Account As String, Begin As String, Ending As String, Metadata As Worksheet, OrderedData As Worksheet) As Double
    Dim MetadataRow As Integer: MetadataRow = GetNextRow(Metadata)
    Dim OrderedDataIndex As Integer: OrderedDataIndex = GetNextRow(OrderedData)
    Dim StartIndex As Long
    Dim EndingIndex As Long
    Dim TotalCost As Double: TotalCost = 0
    Dim x As Long
    Metadata.Cells(MetadataRow, 1).Value = Job
    Metadata.Cells(MetadataRow, 2).Value = Account
    Metadata.Cells(MetadataRow, 3).Value = OrderedDataIndex
    Metadata.Cells(MetadataRow, 4).Value = OrderedDataIndex
    If Not IsLabor(Account) Then
        For x = 1 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count
            If ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value = Job Then
                StartIndex = ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 4).Value
                EndingIndex = ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 5).Value
            End If
        Next
        If StartIndex = 0 Then
            GetBudgetFor = 0
            Exit Function
        End If
        For x = StartIndex To EndingIndex
            If ActiveWorkbook.Worksheets("BillDB").Cells(x, 6).Value = Job Then
                If ActiveWorkbook.Worksheets("BillDB").Cells(x, 2).Value = Account Then
                    OrderedData.Cells(OrderedDataIndex, 1).Value = ActiveWorkbook.Worksheets("BillDB").Cells(x, 6).Value
                    OrderedData.Cells(OrderedDataIndex, 2).Value = ActiveWorkbook.Worksheets("BillDB").Cells(x, 2).Value
                    OrderedData.Cells(OrderedDataIndex, 3).Value = ActiveWorkbook.Worksheets("BillDB").Cells(x, 4).Value
                    OrderedData.Cells(OrderedDataIndex, 4).Value = ActiveWorkbook.Worksheets("BillDB").Cells(x, 1).Value
                    OrderedData.Cells(OrderedDataIndex, 5).Value = ActiveWorkbook.Worksheets("BillDB").Cells(x, 3).Value
                    OrderedData.Cells(OrderedDataIndex, 6).Value = ActiveWorkbook.Worksheets("BillDB").Cells(x, 5).Value
                    TotalCost = TotalCost + ActiveWorkbook.Worksheets("BillDB").Cells(x, 3).Value
                    Metadata.Cells(MetadataRow, 4).Value = OrderedDataIndex
                    OrderedDataIndex = OrderedDataIndex + 1
                End If
            End If
        Next
    Else
        For x = 1 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count
            If ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value = Job Then
                StartIndex = ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 2).Value
                EndingIndex = ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 3).Value
            End If
        Next
        If StartIndex = 0 Then
            GetBudgetFor = 0
            Exit Function
        End If
        For x = StartIndex To EndingIndex
            If ActiveWorkbook.Worksheets("HoursDB").Cells(x, 8).Value = Job Then
                If ActiveWorkbook.Worksheets("HoursDB").Cells(x, 1).Value = Account Then
                    OrderedData.Cells(OrderedDataIndex, 1).Value = ActiveWorkbook.Worksheets("HoursDB").Cells(x, 8).Value
                    OrderedData.Cells(OrderedDataIndex, 2).Value = ActiveWorkbook.Worksheets("HoursDB").Cells(x, 1).Value
                    OrderedData.Cells(OrderedDataIndex, 3).Value = ActiveWorkbook.Worksheets("HoursDB").Cells(x, 3).Value
                    OrderedData.Cells(OrderedDataIndex, 4).Value = ActiveWorkbook.Worksheets("HoursDB").Cells(x, 4).Value
                    OrderedData.Cells(OrderedDataIndex, 5).Value = DetermineWage(Account, ActiveWorkbook.Worksheets("HoursDB").Cells(x, 5).Value) * OvertimeModifier(ActiveWorkbook.Worksheets("HoursDB").Cells(x, 4).Value) * ActiveWorkbook.Worksheets("HoursDB").Cells(x, 3).Value
                    OrderedData.Cells(OrderedDataIndex, 6).Value = ActiveWorkbook.Worksheets("HoursDB").Cells(x, 5).Value
                    TotalCost = TotalCost + OrderedData.Cells(OrderedDataIndex, 5).Value
                    Metadata.Cells(MetadataRow, 4).Value = OrderedDataIndex
                    OrderedDataIndex = OrderedDataIndex + 1
                End If
            End If
        Next
    End If
    GetBudgetFor = TotalCost
End Function

Public Function OvertimeModifier(Worktype As String) As Double
    If IsOvertime(Worktype) Then
        OvertimeModifier = 1.5
    Else
        OvertimeModifier = 1
    End If
End Function

Public Function GetNextRow(sheet As Worksheet) As Integer
    Dim x As Long: x = 1
    Do While sheet.Cells(x, 1).Value <> ""
        x = x + 1
    Loop
    GetNextRow = x
End Function

Public Function Newsheet(Sheetname As String) As Worksheet
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = Sheetname
    End With
    Set Newsheet = ActiveWorkbook.Worksheets(Sheetname)
End Function

Public Sub BuildBudgetMain(Begin As String, Ending As String)
    Dim Jobs() As String: Jobs = Enumerate(Begin, Ending, 8, 6)
    Dim Accounts() As String: Accounts = Enumerate(Begin, Ending, 1, 2)
    Dim Metadata As Worksheet: Set Metadata = Newsheet("MainBudgetMetadata")
    Dim OrderedData As Worksheet: Set OrderedData = Newsheet("MainBudgetOrderedData")
    Dim Report As Worksheet: Set Report = Newsheet("MainBudget")
    Dim x As Integer
    Dim y As Integer
    For x = 0 To UBound(Accounts)
        Report.Cells(x + 1, 1).Value = Accounts(x)
    Next
    For y = 0 To UBound(Jobs)
        Report.Cells(1, y + 1).Value = Jobs(y)
    Next
    For x = 0 To UBound(Accounts)
        For y = 0 To UBound(Jobs)
            If Accounts(x) <> "" And Jobs(y) <> "" Then
                Call VBA.DoEvents
                Report.Cells(x + 1, y + 1).Value = GetBudgetFor(Jobs(y), Accounts(x), Begin, Ending, Metadata, OrderedData)
            End If
        Next
    Next
    Report.UsedRange.Columns.AutoFit
End Sub

Public Sub Wrapper()
    Call BuildBudgetMain("1/1/17", "9/21/17")
End Sub

Public Function IsLabor(Account As String) As Boolean
    If IsNumeric(Mid(Account, 1, 1)) Then
        IsLabor = False
    Else
        IsLabor = True
    End If
End Function

Public Sub BuildReport(Optional Job As String, Optional Account As String)
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim i As Long
    Dim total As Double: total = 0
    Dim LaborCost As Double: LaborCost = 0
    Dim BillCost As Double: BillCost = 0
    Dim ReportPointer As Long
    Dim Report As Worksheet
    Dim Metadata As Worksheet: Set Metadata = ActiveWorkbook.Worksheets("MainBudgetMetadata")
    Dim Data As Worksheet: Set Data = ActiveWorkbook.Worksheets("MainBudgetOrderedData")
    If Job <> "" Then
        If Account <> "" Then
            Call Newsheet("Report")
            Set Report = ActiveWorkbook.Worksheets("Report")
            If IsLabor(Account) Then
                Report.Cells(1, 2).Value = "Work Type"
                Report.Cells(1, 3).Value = "Hours"
            Else
                Report.Cells(1, 2).Value = "Vendor"
                Report.Cells(1, 3).Value = "Item"
            End If
            Report.Cells(1, 4).Value = "Cost"
            Report.Cells(1, 5).Value = "Date"
            ReportPointer = 2
            For x = 1 To Metadata.UsedRange.Rows.Count
                If Metadata.Cells(x, 1).Value = Job Then
                    If Metadata.Cells(x, 2).Value = Account Then
                        For y = Metadata.Cells(x, 3).Value To Metadata.Cells(x, 4).Value
                            Report.Cells(ReportPointer, 2).Value = Data.Cells(y, 4).Value
                            Report.Cells(ReportPointer, 3).Value = Data.Cells(y, 3).Value
                            Report.Cells(ReportPointer, 4).Value = Format(Data.Cells(y, 5).Value, "Currency")
                            total = total + Report.Cells(ReportPointer, 4).Value
                            Report.Cells(ReportPointer, 5).Value = Data.Cells(y, 6).Value
                            ReportPointer = ReportPointer + 1
                        Next
                        Report.Cells(ReportPointer, 1).Value = "Total for " & Account & " At " & Job
                        Report.Cells(ReportPointer, 4).Value = Format(total, "Currency")
                    End If
                End If
            Next
        Else
            Call Newsheet("Report")
            Set Report = ActiveWorkbook.Worksheets("Report")
            For x = 2 To ActiveWorkbook.Worksheets("MainBudget").UsedRange.Columns.Count
                If ActiveWorkbook.Worksheets("MainBudget").Cells(1, x).Value = Job Then
                    y = x
                    Exit For
                End If
            Next
            ReportPointer = 5
            For x = 2 To ActiveWorkbook.Worksheets("MainBudget").UsedRange.Rows.Count
                If ActiveWorkbook.Worksheets("MainBudget").Cells(x, y).Value <> 0 Then
                    total = 0
                    Report.Cells(ReportPointer, 1).Value = ActiveWorkbook.Worksheets("MainBudget").Cells(x, 1).Value & ":"
                    ReportPointer = ReportPointer + 1
                    If IsLabor(ActiveWorkbook.Worksheets("MainBudget").Cells(x, 1).Value) Then
                        Report.Cells(ReportPointer, 2).Value = "Work Type"
                        Report.Cells(ReportPointer, 3).Value = "Hours"
                    Else
                        Report.Cells(ReportPointer, 2).Value = "Vendor"
                        Report.Cells(ReportPointer, 3).Value = "Item"
                    End If
                    Report.Cells(ReportPointer, 4).Value = "Cost"
                    Report.Cells(ReportPointer, 5).Value = "Date"
                    ReportPointer = ReportPointer + 1
                    For z = 1 To Metadata.UsedRange.Rows.Count
                        If Metadata.Cells(z, 1).Value = Job Then
                            If Metadata.Cells(z, 2).Value = ActiveWorkbook.Worksheets("MainBudget").Cells(x, 1).Value Then
                                For i = Metadata.Cells(z, 3).Value To Metadata.Cells(z, 4).Value
                                    Report.Cells(ReportPointer, 2).Value = Data.Cells(i, 4).Value
                                    Report.Cells(ReportPointer, 3).Value = Data.Cells(i, 3).Value
                                    Report.Cells(ReportPointer, 4).Value = Format(Data.Cells(i, 5).Value, "Currency")
                                    total = total + Report.Cells(ReportPointer, 4).Value
                                    If IsLabor(ActiveWorkbook.Worksheets("MainBudget").Cells(x, 1).Value) Then
                                        LaborCost = LaborCost + Report.Cells(ReportPointer, 4).Value
                                    Else
                                        BillCost = BillCost + Report.Cells(ReportPointer, 4).Value
                                    End If
                                    Report.Cells(ReportPointer, 5).Value = Data.Cells(i, 6).Value
                                    ReportPointer = ReportPointer + 1
                                Next
                            End If
                        End If
                    Next
                    Report.Cells(ReportPointer, 1).Value = "Total for " & ActiveWorkbook.Worksheets("MainBudget").Cells(x, 1).Value & " At " & Job
                    Report.Cells(ReportPointer, 4).Value = Format(total, "Currency")
                    ReportPointer = ReportPointer + 2
                End If
            Next
            Report.Cells(1, 2).Value = "Total Cost:" & Format(BillCost + (LaborCost * 1.345), "Currency")
            Report.Cells(2, 4).Value = "Date Range:"
            Report.Cells(3, 4).Value = "Start:"
            Report.Cells(4, 4).Value = "End:"
            Report.Cells(2, 2).Value = "Total Cost of Materials/Bills: " & Format(BillCost, "Currency")
            Report.Cells(3, 2).Value = "Total Cost of Labor: " & Format(LaborCost, "Currency")
            Report.Cells(4, 2).Value = "With Burden: " & Format(LaborCost * 1.345, "Currency")
        End If
    Else
        Call Newsheet("Report")
        Set Report = ActiveWorkbook.Worksheets("Report")
        For x = 2 To ActiveWorkbook.Worksheets("MainBudget").UsedRange.Rows.Count
            If ActiveWorkbook.Worksheets("MainBudget").Cells(x, 1).Value = Account Then
                y = x
                Exit For
            End If
        Next
        If IsLabor(Account) Then
            ReportPointer = 4
        Else
            ReportPointer = 3
        End If
        For x = 2 To ActiveWorkbook.Worksheets("MainBudget").UsedRange.Columns.Count
            If ActiveWorkbook.Worksheets("MainBudget").Cells(y, x).Value <> 0 Then
                total = 0
                Report.Cells(ReportPointer, 1).Value = ActiveWorkbook.Worksheets("MainBudget").Cells(1, x).Value & ":"
                ReportPointer = ReportPointer + 1
                If IsLabor(Account) Then
                    Report.Cells(ReportPointer, 2).Value = "Work Type"
                    Report.Cells(ReportPointer, 3).Value = "Hours"
                Else
                    Report.Cells(ReportPointer, 2).Value = "Vendor"
                    Report.Cells(ReportPointer, 3).Value = "Item"
                End If
                Report.Cells(ReportPointer, 4).Value = "Cost"
                Report.Cells(ReportPointer, 5).Value = "Date"
                ReportPointer = ReportPointer + 1
                For z = 1 To Metadata.UsedRange.Rows.Count
                    If Metadata.Cells(z, 1).Value = ActiveWorkbook.Worksheets("MainBudget").Cells(1, x).Value Then
                        If Metadata.Cells(z, 2).Value = Account Then
                            For i = Metadata.Cells(z, 3).Value To Metadata.Cells(z, 4).Value
                                Report.Cells(ReportPointer, 2).Value = Data.Cells(i, 4).Value
                                Report.Cells(ReportPointer, 3).Value = Data.Cells(i, 3).Value
                                Report.Cells(ReportPointer, 4).Value = Format(Data.Cells(i, 5).Value, "Currency")
                                total = total + Report.Cells(ReportPointer, 4).Value
                                If IsLabor(ActiveWorkbook.Worksheets("MainBudget").Cells(y, 1).Value) Then
                                    LaborCost = LaborCost + Report.Cells(ReportPointer, 4).Value
                                Else
                                    BillCost = BillCost + Report.Cells(ReportPointer, 4).Value
                                End If
                                Report.Cells(ReportPointer, 5).Value = Data.Cells(i, 6).Value
                                ReportPointer = ReportPointer + 1
                            Next
                        End If
                    End If
                Next
                Report.Cells(ReportPointer, 1).Value = "Total for " & Account & " At " & ActiveWorkbook.Worksheets("MainBudget").Cells(1, x).Value
                Report.Cells(ReportPointer, 4).Value = Format(total, "Currency")
                ReportPointer = ReportPointer + 2
            End If
        Next
        Report.Cells(2, 4).Value = "Date Range:"
        If IsLabor(Account) Then
            Report.Cells(2, 2).Value = "Total Cost of Labor: " & LaborCost
            Report.Cells(3, 2).Value = "With Burden: " & LaborCost * 1.345
        Else
            Report.Cells(2, 2).Value = "Total Cost of Materials/Bills: " & BillCost
        End If
    End If
    Report.UsedRange.Columns.AutoFit
    Report.Columns(1).ColumnWidth = 10
End Sub


