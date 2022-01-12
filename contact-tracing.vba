Option Explicit
Public Sub CheckStudentAllocations()

    ' Connect to Syllabus+.
    On Error GoTo CouldConnectToSyllabus
    Dim ProgID As String
    ProgID = InputBox("Enter Prog ID", "Prog ID Required", "2122TURING")
    Dim App As SplusServer.Application
    Set App = CreateObject(Trim(ProgID) + ".application")
    Dim Coll As SplusServer.College
    Set Coll = App.ActiveCollege
    
    ' Process the initial list of students.
    On Error GoTo ErrorProcessingStudents
    Dim row As Integer
    Dim studentHK As String
    Dim myStudentSet As SplusServer.StudentSet
    Dim myStudentSets As SplusServer.StudentSets
    Set myStudentSets = Coll.CreateStudentSets
    row = 2
    studentHK = Sheet1.Range("A" + Trim(Str(row))).Text
    
    While Len(studentHK) > 0
        Select Case Coll.StudentSets.Find(HostKey:=studentHK).Count
            Case 0:
                Sheet1.Range("B" + Trim(Str(row))).Value = "Student Set Not Found"
            Case 1:
                Set myStudentSet = Coll.StudentSets.Find(HostKey:=studentHK).Item(1)
                Sheet1.Range("B" + Trim(Str(row))).Value = myStudentSet.Name
                Call myStudentSets.Add(myStudentSet)
            Case Else:
                Sheet1.Range("B" + Trim(Str(row))).Value = "Multiple Student Sets Found"
        End Select
        row = row + 1
        studentHK = Sheet1.Range("A" + Trim(Str(row))).Text
    Wend
    
    ' Get the date range parameters we're searching for.
    On Error GoTo ErrorProcessingDatesProvided
    Dim startDateTime As Date
    Dim endDateTime As Date
    startDateTime = Sheet1.Range("E2").Value
    endDateTime = Sheet1.Range("F2").Value
    Dim periodToInvestigate As SplusServer.PeriodInYearPattern
    Set periodToInvestigate = Coll.CreatePeriodInYearPattern
    Call periodToInvestigate.SetByDateTimeRange(startDateTime, endDateTime, True)
    
    ' We now have a collection of the student sets we're investigating.
    ' We need to find their activities in the given timeframe.
    On Error GoTo ErrorRetrievingActivities
    Dim myActivity As SplusServer.Activity
    Dim myActivities As SplusServer.Activities
    Set myActivities = Coll.CreateActivities
    For Each myStudentSet In myStudentSets
        For Each myActivity In myStudentSet.ActivitiesAllocatedTo
            If myActivity.SchedulingStatus = cpSchedulingStatusTypeScheduled Then
                If myActivity.ScheduledStartPeriods.Intersects(periodToInvestigate) = True Then
                    Call myActivities.Add(myActivity)
                End If
            End If
        Next myActivity
    Next myStudentSet
    
    ' Clear Sheet 2 from previous runs.
    On Error GoTo ErrorLabellingRowsAndColumns
    Sheet2.Range("A1:ZZ9999").Clear
    
    ' List out the students in Sheet 2.
    For row = 2 To myStudentSets.Count + 1
        Sheet2.Range("A" + Trim(Str(row))).Value = myStudentSets.Item(row - 1).Name
    Next row
    
    ' List out the activities in Sheet 2.
    Dim col As Integer
    For col = 2 To myActivities.Count + 1
        Sheet2.Cells(1, col).Value = myActivities.Item(col - 1).Name
        Sheet2.Cells(1, col).Orientation = 90
    Next col
    
    ' Now we need to work out whether each student is at each activity.
    On Error GoTo ErrorLoggingResults
    For row = 2 To myStudentSets.Count + 1
        For col = 2 To myActivities.Count + 1
            Set myStudentSet = myStudentSets.Item(row - 1)
            Set myActivity = myActivities.Item(col - 1)
            If myStudentSet.ActivitiesAllocatedTo.Find(HostKey:=myActivity.HostKey).Count > 0 Then
                Sheet2.Cells(row, col).Value = "X"
            End If
        Next col
    Next row
    
    Exit Sub
    
    ' Reporting on errors.
    
CouldConnectToSyllabus:
    MsgBox ("An error occurred when connecting to Syllabus+")
    Exit Sub

ErrorProcessingStudents:
    MsgBox ("An error occurred when processing the list of students")
    Exit Sub
    
ErrorProcessingDatesProvided:
    MsgBox ("An error occurred when processing the start and end dates provided")
    Exit Sub
    
ErrorRetrievingActivities:
    MsgBox ("An error occurred when retrieving the list of activities for these students")
    Exit Sub
    
ErrorLabellingRowsAndColumns:
    MsgBox ("An error occurred when applying column and row headers to the results")
    Exit Sub
    
ErrorLoggingResults:
    MsgBox ("An error occurred when logging results")
    Exit Sub

End Sub
