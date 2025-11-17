Attribute VB_Name = "modStandardWorkAutomation"
Option Explicit

' Public entry point that normalizes every section on the "STDW Form" tab
' after the Basic (STDWork_tbl) and Specific (SpecificWork_tbl) tables are edited.
Public Sub UpdateStandardWorkForm()
    Dim wsSource As Worksheet
    Dim wsForm As Worksheet

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wsSource = ThisWorkbook.Worksheets("Standard Work")
    Set wsForm = ThisWorkbook.Worksheets("STDW Form")

    Dim basicTable As ListObject
    Dim specificTable As ListObject

    Set basicTable = wsSource.ListObjects("STDWork_tbl")
    Set specificTable = wsSource.ListObjects("SpecificWork_tbl")

    Dim startTasks As Variant
    Dim duringTasks As Variant
    Dim endTasks As Variant
    Dim weeklyTasks As Variant
    Dim specificTasks As Variant

    startTasks = CollectDailyTasks(basicTable, "Start of Shift")
    duringTasks = CollectDailyTasks(basicTable, "During Shift")
    endTasks = CollectDailyTasks(basicTable, "End of Shift")
    weeklyTasks = CollectWeeklyTasks(basicTable)
    specificTasks = CollectSpecificTasks(specificTable)

    WriteDailySection wsForm, "Start Of Shift Tasks", "During Shift Tasks", startTasks, 1
    WriteDailySection wsForm, "During Shift Tasks", "End of Shift Tasks", duringTasks, 1
    WriteDailySection wsForm, "End of Shift Tasks", "Weekly Tasks", endTasks, 13

    WriteWeeklyAndSpecific wsForm, weeklyTasks, specificTasks

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "Unable to refresh the Standard Work form: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Function CollectDailyTasks(ByVal lo As ListObject, ByVal whenText As String) As Variant
    Dim results() As Variant
    Dim count As Long
    Dim normalizedWhen As String

    normalizedWhen = NormalizeText(whenText)

    Dim lr As ListRow
    For Each lr In lo.ListRows
        Dim rowCells As Range
        Dim taskName As String
        Dim whenValue As String

        Set rowCells = lr.Range
        taskName = Trim$(CStr(rowCells.Cells(1, 3).Value))
        If Len(taskName) = 0 Then GoTo ContinueLoop

        whenValue = NormalizeText(rowCells.Cells(1, 4).Value)
        If whenValue = normalizedWhen Then
            count = count + 1
            ReDim Preserve results(1 To count, 1 To 3)
            results(count, 1) = taskName
            results(count, 2) = Trim$(CStr(rowCells.Cells(1, 2).Value))
            results(count, 3) = Trim$(CStr(rowCells.Cells(1, 5).Value))
        End If
ContinueLoop:
    Next lr

    If count > 0 Then
        CollectDailyTasks = results
    Else
        CollectDailyTasks = Empty
    End If
End Function

Private Function CollectWeeklyTasks(ByVal lo As ListObject) As Variant
    Dim results() As Variant
    Dim count As Long

    Dim lr As ListRow
    For Each lr In lo.ListRows
        Dim rowCells As Range
        Dim frequencyText As String
        Dim taskName As String

        Set rowCells = lr.Range
        taskName = Trim$(CStr(rowCells.Cells(1, 3).Value))
        If Len(taskName) = 0 Then GoTo ContinueLoop

        frequencyText = NormalizeText(rowCells.Cells(1, 5).Value)
        If InStr(1, frequencyText, "week", vbTextCompare) > 0 Then
            count = count + 1
            ReDim Preserve results(1 To count, 1 To 1)
            results(count, 1) = taskName
        End If
ContinueLoop:
    Next lr

    If count > 0 Then
        CollectWeeklyTasks = results
    Else
        CollectWeeklyTasks = Empty
    End If
End Function

Private Function CollectSpecificTasks(ByVal lo As ListObject) As Variant
    Dim results() As Variant
    Dim count As Long

    Dim lr As ListRow
    For Each lr In lo.ListRows
        Dim rowCells As Range
        Dim taskName As String

        Set rowCells = lr.Range
        taskName = Trim$(CStr(rowCells.Cells(1, 3).Value))
        If Len(taskName) = 0 Then GoTo ContinueLoop

        count = count + 1
        ReDim Preserve results(1 To count, 1 To 1)
        results(count, 1) = taskName
ContinueLoop:
    Next lr

    If count > 0 Then
        CollectSpecificTasks = results
    Else
        CollectSpecificTasks = Empty
    End If
End Function

Private Sub WriteDailySection(ByVal ws As Worksheet, _
                              ByVal headerText As String, _
                              ByVal nextHeaderText As String, _
                              ByVal tasks As Variant, _
                              ByVal nextHeaderColumn As Long)
    Dim headerRow As Long
    Dim nextHeaderRow As Long
    Dim insertRow As Long
    Dim rowCount As Long

    headerRow = FindRow(ws, headerText, 1)
    nextHeaderRow = FindRow(ws, nextHeaderText, nextHeaderColumn)

    ClearRowsBetween ws, headerRow, nextHeaderRow

    rowCount = MatrixRowCount(tasks)
    If rowCount = 0 Then rowCount = 1

    insertRow = headerRow + 1
    ws.Rows(insertRow).Resize(rowCount).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    If MatrixRowCount(tasks) > 0 Then
        ws.Range("A" & insertRow).Resize(rowCount, 3).Value = tasks
    Else
        ws.Range("A" & insertRow).Resize(rowCount, 3).ClearContents
    End If
End Sub

Private Sub WriteWeeklyAndSpecific(ByVal ws As Worksheet, _
                                   ByVal weeklyTasks As Variant, _
                                   ByVal specificTasks As Variant)
    Dim weeklyHeaderRow As Long
    Dim teamHeaderRow As Long
    Dim notesHeaderRow As Long

    weeklyHeaderRow = FindRow(ws, "Weekly Tasks", 13)
    teamHeaderRow = FindRow(ws, "Team Member Specific Tasks", 13)

    WriteVerticalSection ws, weeklyHeaderRow, teamHeaderRow, weeklyTasks, "N"

    teamHeaderRow = FindRow(ws, "Team Member Specific Tasks", 13)
    notesHeaderRow = FindRow(ws, "Notes, Issues / Roadblocks, Concerns, or Suggestions", 13)

    WriteVerticalSection ws, teamHeaderRow, notesHeaderRow, specificTasks, "N"

    UpdateNotesSection ws
End Sub

Private Sub WriteVerticalSection(ByVal ws As Worksheet, _
                                 ByVal headerRow As Long, _
                                 ByVal nextHeaderRow As Long, _
                                 ByVal items As Variant, _
                                 ByVal targetColumn As String)
    Dim rowCount As Long
    Dim insertRow As Long

    ClearRowsBetween ws, headerRow, nextHeaderRow

    rowCount = MatrixRowCount(items)
    If rowCount = 0 Then rowCount = 1

    insertRow = headerRow + 1
    ws.Rows(insertRow).Resize(rowCount).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    If MatrixRowCount(items) > 0 Then
        ws.Range(targetColumn & insertRow).Resize(rowCount, 1).Value = items
    Else
        ws.Range(targetColumn & insertRow).Resize(rowCount, 1).ClearContents
    End If
End Sub

Private Sub UpdateNotesSection(ByVal ws As Worksheet)
    Dim notesHeaderRow As Long
    Dim certificateRow As Long

    notesHeaderRow = FindRow(ws, "Notes, Issues / Roadblocks, Concerns, or Suggestions", 13)
    certificateRow = FindRow(ws, "I certify that all required Daily and Weekly Standard Work checks for the week shown have been completed, and any exceptions are recorded in the Notes / Issues section or escalated to Leadership.", 13)

    ClearRowsBetween ws, notesHeaderRow, certificateRow
    ws.Rows(notesHeaderRow + 1).Resize(7).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub

Private Sub ClearRowsBetween(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal nextHeaderRow As Long)
    Dim firstRow As Long
    Dim lastRow As Long

    firstRow = headerRow + 1
    lastRow = nextHeaderRow - 1

    If lastRow >= firstRow Then
        ws.Rows(firstRow & ":" & lastRow).Delete
    End If
End Sub

Private Function FindRow(ByVal ws As Worksheet, ByVal searchText As String, ByVal columnIndex As Long) As Long
    Dim rng As Range
    Dim searchColumn As Range
    Dim firstAddress As String

    Set searchColumn = ws.Columns(columnIndex)
    With searchColumn
        Set rng = .Find(What:=searchText, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=False)
        If rng Is Nothing Then
            Err.Raise vbObjectError + 513, "modStandardWorkAutomation", _
                      "Unable to locate """ & searchText & """ in column " & columnIndex & "."
        End If
    End With

    FindRow = rng.Row
End Function

Private Function MatrixRowCount(ByVal matrix As Variant) As Long
    If IsEmpty(matrix) Then
        MatrixRowCount = 0
    Else
        MatrixRowCount = UBound(matrix, 1) - LBound(matrix, 1) + 1
    End If
End Function

Private Function NormalizeText(ByVal value As Variant) As String
    NormalizeText = LCase$(Trim$(CStr(value)))
End Function
