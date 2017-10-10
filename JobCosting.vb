Sub CalculageWages()
    Dim x As Integer: x = 5
    Do While x < Application.ActiveWorkbook.Worksheets("JobReporting").UsedRange.Rows.Count + 1
        Application.ActiveWorkbook.Worksheets("JobReporting").Cells(x, 8).Value = FindWage(Application.ActiveWorkbook.Worksheets("JobReporting").Cells(x, 1).Value, Application.ActiveWorkbook.Worksheets("Hours List").Cells(x, 2).Value)
        x = x + 1
    Loop
End Sub

Public Function AddJobNameColumn()
    Dim x As Long
    Dim jobname As String
    For x = 2 To ActiveWorkbook.Worksheets("HoursDB").UsedRange.Rows.Count + 1
        VBA.DoEvents
        jobname = ActiveWorkbook.Worksheets("HoursDB").Cells(x, 2).Value
        If ActiveWorkbook.Worksheets("HoursDB").Cells(x, 6).Value <> "" Then
            jobname = ActiveWorkbook.Worksheets("HoursDB").Cells(x, 6).Value & ":" & jobname
            If ActiveWorkbook.Worksheets("HoursDB").Cells(x, 7).Value <> "" Then
                jobname = ActiveWorkbook.Worksheets("HoursDB").Cells(x, 7).Value & ":" & jobname
            End If
        End If
        ActiveWorkbook.Worksheets("HoursDB").Cells(x, 8).Value = jobname
    Next
End Function

Public Function FindWage(Name, NDate) As Double
    Dim x As Integer: x = 2
    Do While x < Application.ActiveWorkbook.Worksheets("Wage List").UsedRange.Rows.Count + 1
        If Application.ActiveWorkbook.Worksheets("Wage List").Cells(x, 1).Value = Name Then
            Dim y As Integer: y = 4
            Do While y < Application.ActiveWorkbook.Worksheets("Wage List").UsedRange.Columns.Count + 1
                If Application.ActiveWorkbook.Worksheets("Wage List").Cells(x, y).Value = "" Then
                    FindWage = Application.ActiveWorkbook.Worksheets("Wage List").Cells(x, y - 1).Value
                    Exit Function
                ElseIf DateCompare2(NDate, Application.ActiveWorkbook.Worksheets("Wage List").Cells(x, y).Value) Then
                    FindWage = Application.ActiveWorkbook.Worksheets("Wage List").Cells(x, y - 1).Value
                    Exit Function
                End If
                y = y + 2
            Loop
        End If
        x = x + 1
    Loop
End Function



Public Function DateToString(D As String) As String
    Dim Returned As String
    Dim D2() As String: D2 = Split(D, "/")
    Returned = D2(0) & "-" & D2(1) & "-" & Mid(D2(2), 3, 2)
    DateToString = Returned
    Exit Function
End Function


Sub InsertNewMetadataNode(s As String, ByRef Metadata() As Variant, ByRef index As Integer, appearance As Long)
    Dim top As Integer: top = index
    Dim bottom As Integer: bottom = 0
    Dim Returned As Integer
    Dim tempLocation As Integer: tempLocation = bottom + (top - bottom) \ 2
    Dim PlacedIndex As Integer: PlacedIndex = -1
    If top = 0 Then
        Metadata(0, 0) = s
        Metadata(0, 1) = appearance
        Metadata(0, 2) = appearance
        index = index + 1
        Exit Sub
    ElseIf top = 1 Then
        If StrComp(s, Metadata(0, 0)) = 1 Then
            Metadata(1, 0) = s
            Metadata(1, 1) = appearance
            Metadata(1, 2) = appearance
            index = index + 1
            Exit Sub
        ElseIf StrComp(s, Metadata(0, 0)) = 0 Then
            If appearance > Metadata(0, 2) Then
                Metadata(0, 2) = appearance
            End If
            Exit Sub
        Else
            Metadata(1, 0) = Metadata(0, 0)
            Metadata(1, 1) = Metadata(0, 1)
            Metadata(1, 2) = Metadata(0, 2)
            Metadata(0, 0) = s
            Metadata(0, 1) = appearance
            Metadata(0, 2) = appearance
            index = index + 1
            Exit Sub
        End If
    End If
    Do While True
        Returned = StrComp(s, Metadata(tempLocation, 0), vbTextCompare)
        If Returned = 1 Then
            bottom = tempLocation
        ElseIf Returned = 0 Then
            StringInList = tempLocation
            If appearance > Metadata(tempLocation, 2) Then
                Metadata(tempLocation, 2) = appearance
            End If
            Exit Sub
        Else
            top = tempLocation
        End If
        If top - bottom = 1 Then
            If index = UBound(Metadata, 1) Then
                Dim TempData() As Variant
                ReDim TempData(UBound(Metadata, 1), UBound(Metadata, 2))
                Dim y As Integer: y = 0
                Dim z As Integer: z = 0
                For y = 0 To UBound(Metadata, 1)
                    For z = 0 To UBound(Metadata, 2)
                        TempData(y, z) = Metadata(y, z)
                    Next
                Next
                ReDim Metadata(UBound(Metadata, 1) * 2, 3)
                For y = 0 To UBound(TempData, 1)
                    For z = 0 To UBound(TempData, 2)
                        Metadata(y, z) = TempData(y, z)
                    Next
                Next
            End If
            If s = Metadata(bottom, 0) Then
                Metadata(bottom, 2) = appearance
                Exit Sub
            ElseIf s = Metadata(top, 0) Then
                Metadata(top, 2) = appearance
                Exit Sub
            ElseIf (StrComp(s, Metadata(bottom, 0), vbTextCompare) = -1 Or IsEmpty(Metadata(bottom, 0))) Then
                PlacedIndex = bottom
            ElseIf StrComp(s, Metadata(top, 0), vbTextCompare) = -1 Or IsEmpty(Metadata(top, 0)) Then
                PlacedIndex = top
            ElseIf StrComp(s, Metadata(top, 0), vbTextCompare) = 1 Then
                PlacedIndex = top + 1
            End If
            If PlacedIndex <> -1 Then
                For x = index To PlacedIndex Step -1
                    Metadata(x + 1, 0) = Metadata(x, 0)
                    Metadata(x + 1, 1) = Metadata(x, 1)
                    Metadata(x + 1, 2) = Metadata(x, 2)
                Next
                Metadata(PlacedIndex, 0) = s
                Metadata(PlacedIndex, 1) = appearance
                Metadata(PlacedIndex, 2) = appearance
                index = index + 1
                Exit Sub
            End If
        End If
        tempLocation = bottom + (top - bottom) \ 2
    Loop
End Sub

Sub BuildMetadata()
    Dim Metadata() As Variant
    ReDim Preserve Metadata(100, 3)
    Dim CurrentLength As Integer: CurrentLength = 100
    Dim CurrentPointer As Integer: CurrentPointer = 0
    Dim x As Long: x = 1
    Dim index As Integer: index = 0
    Dim JN As String
    For x = 2 To ActiveWorkbook.Worksheets("HoursDB").UsedRange.Rows.Count
        JN = ActiveWorkbook.Worksheets("HoursDB").Cells(x, 2).Value
        If ActiveWorkbook.Worksheets("HoursDB").Cells(x, 6).Value <> "" Then
            JN = ActiveWorkbook.Worksheets("HoursDB").Cells(x, 6).Value & ":" & JN
            If ActiveWorkbook.Worksheets("HoursDB").Cells(x, 7).Value <> "" Then
                JN = ActiveWorkbook.Worksheets("HoursDB").Cells(x, 7).Value & ":" & JN
            End If
        End If
        Call InsertNewMetadataNode(JN, Metadata, index, x)
    Next
    For x = 0 To UBound(Metadata, 1)
        ActiveWorkbook.Worksheets("MetadataDB").Cells(x + 1, 1).Value = Metadata(x, 0)
        ActiveWorkbook.Worksheets("MetadataDB").Cells(x + 1, 2).Value = Metadata(x, 1)
        ActiveWorkbook.Worksheets("MetadataDB").Cells(x + 1, 3).Value = Metadata(x, 2)
    Next
End Sub

Sub BuildMetadata3()
    Dim x As Long
    Dim Temp As Integer
    Dim Jobs() As String: Jobs = Enumerate("1-1-1990", CStr(Date), 8, 6)
    Dim JobData() As Long: ReDim JobData(UBound(Jobs), 4)
    Dim HoursDB As Worksheet: Set HoursDB = ActiveWorkbook.Worksheets("HoursDB")
    Dim BillDB As Worksheet: Set BillDB = ActiveWorkbook.Worksheets("BillDB")
    For x = 2 To HoursDB.UsedRange.Rows.Count
        If HoursDB.Cells(x, 1).Value = "" Then
            Exit For
        End If
        Temp = BinarySearch(HoursDB.Cells(x, 8).Value, Jobs)
        JobData(Temp, 1) = x
        If JobData(Temp, 0) = 0 Then
            JobData(Temp, 0) = x
        End If
    Next
    For x = 2 To BillDB.UsedRange.Rows.Count
        If BillDB.Cells(x, 1).Value = "" Then
            Exit For
        End If
        Temp = BinarySearch(BillDB.Cells(x, 6).Value, Jobs)
        JobData(Temp, 3) = x
        If JobData(Temp, 2) = 0 Then
            JobData(Temp, 2) = x
        End If
    Next
    For x = 0 To UBound(Jobs)
        ActiveWorkbook.Worksheets("MetadataDB").Cells(x + 1, 1).Value = Jobs(x)
        ActiveWorkbook.Worksheets("MetadataDB").Cells(x + 1, 2).Value = JobData(x, 0)
        ActiveWorkbook.Worksheets("MetadataDB").Cells(x + 1, 3).Value = JobData(x, 1)
        ActiveWorkbook.Worksheets("MetadataDB").Cells(x + 1, 4).Value = JobData(x, 2)
        ActiveWorkbook.Worksheets("MetadataDB").Cells(x + 1, 5).Value = JobData(x, 3)
    Next
End Sub

Sub ReadMetadata(ByRef Metadata() As Variant)
    Dim x As Integer
    Dim MetadataPointer As Integer: MetadataPointer = 0
    For x = 0 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count
        If ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1) <> "" Then
            If MetadataPointer = Metadata.Length Then
                ReDim Preserve Metadata(Metadata.Length * 2)
            End If
            Metadata(MetadataPointer, 0) = ActiveWorkbook.Worksheets("MetadataDB").Cells(x + 1, 1).Value
            Metadata(MetadataPointer, 1) = ActiveWorkbook.Worksheets("MetadataDB").Cells(x + 1, 2).Value
            Metadata(MetadataPointer, 2) = ActiveWorkbook.Worksheets("MetadataDB").Cells(x + 1, 3).Value
            MetadataPointer = MetadataPointer + 1
        End If
    Next
End Sub

Public Function IsOvertime(HoursType As String) As Boolean
    If InStr(HoursType, "Overtime") Then
        IsOvertime = True
    Else
        IsOvertime = False
    End If
End Function

Sub HoursQuery(jobname As String)

    Dim cnn As ADODB.Connection, rstNew As ADODB.Recordset
    Dim strSQL As String
    Set cnn = New ADODB.Connection
    With cnn
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & ActiveWorkbook.FullName & ";" & _
        "Extended Properties=Excel 8.0;"
        .Open
    End With
    Dim Address As String: Address = "HoursDB!" & ActiveWorkbook.Worksheets("HoursDB").Range("Table_Query_from_Romano_Masonry_201610").Address
    strSQL = "Select [name], [Hours], [time_activity_date], Case When [name3] Like '%Overtime%' Then '1' Else '' End As Overtime FROM [Table_Query_from_Romano_Masonry_201610] WHERE [name2]= '" & jobname & "'"
    Set rstNew = cnn.Execute(strSQL)
    ActiveWorkbook.Worksheets("JobReporting").Range("A5").CopyFromRecordset rstNew
    rstNew.Close
    Set rstNew = Nothing
    cnn.Close
    Set cnn = Nothing
End Sub

Sub GetHours(jobname As String)
    Dim JobData() As Variant
    ReDim Preserve JobData(3)
    Dim JobReportingPointer As Integer: JobReportingPointer = 5
    Dim x As Long
    Dim JN As String
    JobData(0) = jobname
    For x = 1 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count + 1
        If ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value = jobname Then
            JobData(1) = ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 2)
            JobData(2) = ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 3)
        End If
    Next
    Dim StartDate As String: StartDate = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 4).Value
    Dim EndDate As String: EndDate = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 5).Value
    If StartDate = "" Then
        StartDate = ActiveWorkbook.Worksheets("HoursDB").Cells(JobData(1), 5).Value
    End If
    If EndDate = "" Then
        EndDate = ActiveWorkbook.Worksheets("HoursDB").Cells(JobData(2), 5).Value
    End If
    For x = JobData(1) - 1 To JobData(2) + 1
        If ActiveWorkbook.Worksheets("HoursDB").Cells(x, 8).Value = jobname And DateInside(ActiveWorkbook.Worksheets("HoursDB").Cells(x, 5).Value, StartDate, EndDate) Then
            ActiveWorkbook.Worksheets("JobReporting").Cells(JobReportingPointer, 1).Value = ActiveWorkbook.Worksheets("HoursDB").Cells(x, 1).Value
            ActiveWorkbook.Worksheets("JobReporting").Cells(JobReportingPointer, 2).Value = ActiveWorkbook.Worksheets("HoursDB").Cells(x, 3).Value / 24
            ActiveWorkbook.Worksheets("JobReporting").Cells(JobReportingPointer, 3).Value = ActiveWorkbook.Worksheets("HoursDB").Cells(x, 5).Value
            If IsOvertime(ActiveWorkbook.Worksheets("HoursDB").Cells(x, 4).Value) Then
                ActiveWorkbook.Worksheets("JobReporting").Cells(JobReportingPointer, 4).Value = 1
            Else
                ActiveWorkbook.Worksheets("JobReporting").Cells(JobReportingPointer, 4).Value = ""
            End If
            JobReportingPointer = JobReportingPointer + 1
        End If
    Next
End Sub

Sub PopulateComboBox()
    Dim Jobs() As String
    Dim x As Integer
    ReDim Jobs(ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count)
    For x = 1 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count + 1
        Jobs(x - 1) = ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value
    Next
    Set curCombo = ActiveSheet.Shapes.AddFormControl(xlDropDown, Left:=Cells(3, 1).Left, top:=Cells(3, 1).top, Width:=100, Height:=20)
    With curCombo
        .ControlFormat.DropDownLines = ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count
        For x = 0 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count
            .ControlFormat.AddItem Jobs(x)
        Next
        .Name = "JobSelection"
        .OnAction = "myCombo_Change"
    End With
End Sub

Sub RebuildComboBox()
    Dim ComboBox As Shape
    Dim StartDate As String
    Dim x As Integer
    If ActiveWorkbook.Worksheets("JobReporting").Cells(3, 4).Value <> "" Then
        StartDate = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 4).Value
    Else
        StartDate = ActiveWorkbook.Worksheets("HoursDB").Cells(2, 5).Value
    End If
    Dim EndDate As String
    If ActiveWorkbook.Worksheets("JobReporting").Cells(3, 5).Value <> "" Then
        EndDate = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 5).Value
    Else
        EndDate = CStr(Date)
    End If
    Dim AlphaStart As String
    If ActiveWorkbook.Worksheets("JobReporting").Cells(3, 6).Value <> "" Then
        AlphaStart = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 6).Value
    Else
        AlphaStart = ActiveWorkbook.Worksheets("MetadataDB").Cells(1, 1).Value
    End If
    Dim AlphaEnd As String
    If ActiveWorkbook.Worksheets("JobReporting").Cells(3, 7).Value <> "" Then
        AlphaEnd = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 7).Value
    Else
        For x = 2 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count + 2
            If ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value = "" Then
                AlphaEnd = ActiveWorkbook.Worksheets("MetadataDB").Cells(x - 1, 1).Value
                Exit For
            End If
        Next
    End If
    Dim IterDateE As String
    Dim IterDateS As String
    Dim y As Long
    Dim z As Integer: z = 0
    For Each ComboBox In ActiveSheet.Shapes
        If ComboBox.Name = "JobSelection" Then
            ComboBox.ControlFormat.RemoveAllItems
            For x = 1 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count - 1
                y = ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 2).Value
                If y = 0 Then
                    Exit For
                End If
                IterDateS = ActiveWorkbook.Worksheets("HoursDB").Cells(y, 5).Value
                y = ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 3).Value
                IterDateE = ActiveWorkbook.Worksheets("HoursDB").Cells(y, 5).Value
                If DatesCross(StartDate, EndDate, IterDateS, IterDateE) Then
                    If StrComp(ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value, AlphaStart) >= 0 Then
                        If StrComp(AlphaEnd, ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value) >= 0 Then
                            z = z + 1
                        End If
                    End If
                End If
            Next
            ComboBox.ControlFormat.DropDownLines = 1024
            ComboBox.Delete
        End If
    Next
    Set curCombo = ActiveSheet.Shapes.AddFormControl(xlDropDown, Left:=Cells(3, 1).Left, top:=Cells(3, 1).top, Width:=100, Height:=20)
    With curCombo
        .ControlFormat.DropDownLines = Application.Max(0, z - 1)
        For x = 1 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count
            y = ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 2).Value
            If y = 0 Then
                Exit For
            End If
            IterDateS = ActiveWorkbook.Worksheets("HoursDB").Cells(y, 5).Value
            y = ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 3).Value
            IterDateE = ActiveWorkbook.Worksheets("HoursDB").Cells(y, 5).Value
            If DatesCross(StartDate, EndDate, IterDateS, IterDateE) Then
                If StrComp(ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value, AlphaStart) >= 0 Then
                    If StrComp(AlphaEnd, ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value) >= 0 Then
                        .ControlFormat.AddItem ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value
                    End If
                End If
            End If
        Next
        .Name = "JobSelection"
        .OnAction = "myCombo_Change"
    End With
End Sub

Public Function GetJobName(Optional ArgA As String = "-1", Optional ArgB As String = "-1", Optional ArgC As String = "-1", Optional ArgD As String = "-1") As String
    Dim ComboBox As Shape
    Dim x As Long
    Dim IterDateS As String
    Dim IterDateE As String
    Dim StartDate As String
    Dim AlphaStart As String
    Dim AlphaEnd As String
    Dim jobname As String
    If ArgA <> "-1" Then
        If ArgA <> "" Then
            StartDate = ArgA
        Else
            StartDate = ActiveWorkbook.Worksheets("HoursDB").Cells(2, 5).Value
        End If
    ElseIf ActiveWorkbook.Worksheets("JobReporting").Cells(3, 4).Value <> "" Then
        StartDate = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 4).Value
    Else
        StartDate = ActiveWorkbook.Worksheets("HoursDB").Cells(2, 5).Value
    End If
    Dim EndDate As String
    If ArgB <> "-1" Then
        If ArgB <> "" Then
            EndDate = ArgB
        Else
            EndDate = CStr(Date)
        End If
    ElseIf ActiveWorkbook.Worksheets("JobReporting").Cells(3, 5).Value <> "" Then
        EndDate = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 5).Value
    Else
        EndDate = CStr(Date)
    End If
    If ArgC <> "-1" Then
        If ArgC <> "" Then
            AlphaStart = ArgC
        Else
            AlphaStart = ActiveWorkbook.Worksheets("MetadataDB").Cells(1, 1).Value
        End If
    ElseIf ActiveWorkbook.Worksheets("JobReporting").Cells(3, 6).Value = "" Then
        AlphaStart = ActiveWorkbook.Worksheets("MetadataDB").Cells(1, 1).Value
    Else
        AlphaStart = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 6).Value
    End If
    If ArgD <> "-1" Then
        If ArgD <> "" Then
            AlphaEnd = ArgD
        Else
            For x = 1 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count + 2
                If ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value = "" Then
                    AlphaEnd = ActiveWorkbook.Worksheets("MetadataDB").Cells(x - 1, 1).Value
                    Exit For
                End If
            Next
        End If
    ElseIf ActiveWorkbook.Worksheets("JobReporting").Cells(3, 7).Value = "" Then
        For x = 1 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count + 2
            If ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value = "" Then
                AlphaEnd = ActiveWorkbook.Worksheets("MetadataDB").Cells(x - 1, 1).Value
                Exit For
            End If
        Next
    Else
        AlphaEnd = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 7).Value
    End If
    For Each ComboBox In ActiveWorkbook.Worksheets("JobReporting").Shapes
        If ComboBox.Name = "JobSelection" Then
            Dim JobNumber As Integer: JobNumber = ComboBox.OLEFormat.Object.Value
            If JobNumber = 0 Then
                GetJobName = ""
                Exit Function
            End If
            For x = 1 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count
                If ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value = "" Then
                    Exit For
                End If
                IterDateS = ActiveWorkbook.Worksheets("HoursDB").Cells(ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 2).Value, 5).Value
                IterDateE = ActiveWorkbook.Worksheets("HoursDB").Cells(ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 3).Value, 5).Value
                If DatesCross(StartDate, EndDate, IterDateS, IterDateE) Then
                    If StrComp(ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value, AlphaStart) >= 0 Then
                        If StrComp(AlphaEnd, ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value) >= 0 Then
                            JobNumber = JobNumber - 1
                        End If
                    End If
                End If
                If JobNumber = 0 Then
                    GetJobName = ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value
                    Exit Function
                End If
            Next
        End If
    Next
End Function

Public Function GetJobIndex(jobname As String) As Integer
    Dim StartDate As String: StartDate = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 4).Value
    Dim EndDate As String: EndDate = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 5).Value
    Dim AlphaStart As String: AlphaStart = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 6).Value
    Dim AlphaEnd As String: AlphaEnd = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 7).Value
    Dim IterDateS As String
    Dim IterDateE As String
    Dim x As Long
    Dim JobsList() As String
    ReDim Preserve JobsList(1000)
    Dim JobsListPointer As Integer: JobsListPointer = 0
    If AlphaStart = "" Then
        AlphaStart = ActiveWorkbook.Worksheets("MetadataDB").Cells(1, 1).Value
    End If
    If AlphaEnd = "" Then
        x = 1
        Do While ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value <> ""
            x = x + 1
        Loop
        AlphaEnd = ActiveWorkbook.Worksheets("MetadataDB").Cells(x - 1, 1).Value
    End If
    If StartDate = "" Then
        StartDate = ActiveWorkbook.Worksheets("HoursDB").Cells(2, 5).Value
    End If
    If EndDate = "" Then
        x = 2
        Do While ActiveWorkbook.Worksheets("HoursDB").Cells(x, 5).Value <> ""
            x = x + 1
        Loop
        EndDate = ActiveWorkbook.Worksheets("HoursDB").Cells(x - 1, 5).Value
    End If
    For x = 1 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count
        If ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value = "" Then
            Exit For
        End If
        IterDateS = ActiveWorkbook.Worksheets("HoursDB").Cells(ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 2).Value, 5).Value
        IterDateE = ActiveWorkbook.Worksheets("HoursDB").Cells(ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 3).Value, 5).Value
        If DatesCross(StartDate, EndDate, IterDateS, IterDateE) Then
            If StrComp(ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value, AlphaStart) >= 0 Then
                If StrComp(AlphaEnd, ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value) >= 0 Then
                    If UBound(JobsList) = JobsListPointer Then
                        ReDim Preserve JobsList(UBound(JobsList) * 2)
                    End If
                    JobsList(JobsListPointer) = ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value
                    JobsListPointer = JobsListPointer + 1
                End If
            End If
        End If
    Next
    For x = 0 To JobsListPointer
        If JobsList(x) = jobname Then
            GetJobIndex = x + 1
            Exit Function
        End If
    Next
End Function

Sub SetJob(jobname As String)
    Dim ComboBox As Shape
    Dim JobNumber As Integer
    For Each ComboBox In ActiveWorkbook.Worksheets("JobReporting").Shapes
        If ComboBox.Name = "JobSelection" Then
            JobNumber = GetJobIndex(jobname)
            ComboBox.OLEFormat.Object.Value = JobNumber
        End If
    Next
End Sub

Sub myCombo_Change()
    Dim ComboBox As Shape
    Dim x As Long
    Dim IterDateS As String
    Dim IterDateE As String
    Dim AlphaStart As String
    Dim AlphaEnd As String
    Dim StartDate As String
    Dim jobname As String
    If ActiveWorkbook.Worksheets("JobReporting").Cells(3, 4).Value <> "" Then
        StartDate = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 4).Value
    Else
        StartDate = ActiveWorkbook.Worksheets("HoursDB").Cells(2, 5).Value
    End If
    Dim EndDate As String
    If ActiveWorkbook.Worksheets("JobReporting").Cells(3, 5).Value <> "" Then
        EndDate = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 5).Value
    Else
        EndDate = CStr(Date)
    End If
    If ActiveWorkbook.Worksheets("JobReporting").Cells(3, 6).Value <> "" Then
        AlphaStart = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 6).Value
    Else
        AlphaStart = ActiveWorkbook.Worksheets("MetadataDB").Cells(1, 1).Value
    End If
    If ActiveWorkbook.Worksheets("JobReporting").Cells(3, 7).Value <> "" Then
        AlphaEnd = ActiveWorkbook.Worksheets("JobReporting").Cells(3, 7).Value
    Else
        For x = 1 To ActiveWorkbook.Worksheets("MetadataDB").UsedRange.Rows.Count + 2
            If ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value = "" Then
                AlphaEnd = ActiveWorkbook.Worksheets("MetadataDB").Cells(x - 1, 1).Value
                Exit For
            End If
        Next
    End If
    For Each ComboBox In ActiveWorkbook.Worksheets("JobReporting").Shapes
        If ComboBox.Name = "JobSelection" Then
            Dim JobNumber As Integer: JobNumber = ComboBox.OLEFormat.Object.Value
            If JobNumber = 0 Then
                jobname = ""
                Exit For
            End If
            x = 1
            Do While ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value <> ""
                IterDateS = ActiveWorkbook.Worksheets("HoursDB").Cells(ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 2).Value, 5).Value
                IterDateE = ActiveWorkbook.Worksheets("HoursDB").Cells(ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 3).Value, 5).Value
                If DatesCross(StartDate, EndDate, IterDateS, IterDateE) Then
                    If StrComp(ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value, AlphaStart) >= 0 Then
                        If StrComp(AlphaEnd, ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1).Value) >= 0 Then
                            JobNumber = JobNumber - 1
                        End If
                    End If
                End If
                If JobNumber = 0 Then
                    jobname = ActiveWorkbook.Worksheets("MetadataDB").Cells(x, 1)
                    Exit For
                End If
                x = x + 1
            Loop
        End If
    Next
    Dim JobReportingPointer As Long: JobReportingPointer = 5
    Dim HoursDBPointer As Long
    Dim MaxDBPointer As Long
    Application.ScreenUpdating = False
    ActiveWorkbook.Worksheets("JobReporting").Range("A5:D6000").Value = ""
    If jobname <> "" Then
        Call GetHours(jobname)
    End If
    Application.CalculateFull
    Application.ScreenUpdating = True
End Sub

Public Function DetermineWage(Employee As String, Da As Date) As String
    Dim x As Integer
    If Employee = "" Then
        DetermineWage = ""
        Exit Function
    End If
    For x = 2 To ActiveWorkbook.Worksheets("Employees").UsedRange.Rows.Count
        If ActiveWorkbook.Worksheets("Employees").Cells(x, 1).Value = Employee Or ActiveWorkbook.Worksheets("Employees").Cells(x, 5).Value = Employee Then
            Dim ColumnNumber As Integer: ColumnNumber = 7
            Do While True
                If ActiveWorkbook.Worksheets("Employees").Cells(x, ColumnNumber + 1).Value <> "" Then
                    If DateCompare2(ActiveWorkbook.Worksheets("Employees").Cells(x, ColumnNumber + 1), Date) Then
                        ColumnNumber = ColumnNumber + 2
                    Else
                        DetermineWage = ActiveWorkbook.Worksheets("Employees").Cells(x, ColumnNumber).Value
                        Exit Function
                    End If
                Else
                    DetermineWage = CStr(ActiveWorkbook.Worksheets("Employees").Cells(x, ColumnNumber).Value)
                    Exit Function
                End If
            Loop
        End If
    Next
    x = 5
    Do While ActiveWorkbook.Worksheets("JobReporting").Cells(x, 18).Value <> ""
        If ActiveWorkbook.Worksheets("JobReporting").Cells(x, 18).Value = Employee Then
            DetermineWage = CStr(ActiveWorkbook.Worksheets("JobReporting").Cells(x, 19).Value)
            Exit Function
        End If
        x = x + 1
    Loop
    DetermineWage = CVErr(xlErrValue)
End Function

