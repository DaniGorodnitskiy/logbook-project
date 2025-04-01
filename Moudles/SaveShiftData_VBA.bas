Option Explicit

'=================================================================================
' SaveShiftData
'---------------------------------------------------------------------------------
' 1) Reads:
'    - Date (G5)
'    - Shift (E5)
'    - Operator (D5)
'
' 2) Saves eight blocks to the DB sheet:
'    Block1: D10..D18 => columns D..L   (9 cells)
'    Block2: F10..F18 => columns M..U   (9 cells)
'    Block3: D21..D33 => columns V..AH  (13 cells)
'    Block4: E21..E33 => columns AI..AU (13 cells)
'    Block5: F21..F33 => columns AV..BH (13 cells)
'    Block6: G21..G33 => columns BI..BU (13 cells)
'    Block7: B6..B14  => columns BV..CD (9 cells)
'    Block8: C6..C14  => columns CE..CM (9 cells)
'=================================================================================
Public Sub SaveShiftData()

    Dim wsMain As Worksheet, wsDB As Worksheet
    Set wsMain = ThisWorkbook.Sheets("Main")
    Set wsDB = ThisWorkbook.Sheets("DB")
    
    ' Validate date in G5
    If Not IsDate(wsMain.Range("G5").Value) Then
        MsgBox "Date in G5 is invalid.", vbExclamation
        Exit Sub
    End If
    Dim shiftDate As Date
    shiftDate = CDate(wsMain.Range("G5").Value)
    
    ' Validate shift in E5
    Dim shiftName As String
    shiftName = Trim(wsMain.Range("E5").Value)
    If shiftName = "" Then
        MsgBox "No shift selected in E5.", vbExclamation
        Exit Sub
    End If
    
    ' Operator in D5
    Dim operatorName As String
    operatorName = wsMain.Range("D5").Value
    
    ' Find or create row in DB
    Dim lastRow As Long, rowFound As Long
    rowFound = 0
    lastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        If wsDB.Cells(i, 1).Value = shiftDate And wsDB.Cells(i, 2).Value = shiftName Then
            rowFound = i
            Exit For
        End If
    Next i
    
    If rowFound = 0 Then
        rowFound = lastRow + 1
        wsDB.Cells(rowFound, 1).Value = shiftDate ' col A => date
        wsDB.Cells(rowFound, 2).Value = shiftName ' col B => shift
    End If
    
    ' col C => operator
    wsDB.Cells(rowFound, 3).Value = operatorName
    
    Dim r As Long, c As Long
    
    '--- Block1: D10..D18 => columns D..L ---
    c = 4 ' D=4
    For r = 10 To 18
        wsDB.Cells(rowFound, c).Value = wsMain.Range("D" & r).Value
        c = c + 1
    Next r
    
    '--- Block2: F10..F18 => columns M..U ---
    c = 13 ' M=13
    For r = 10 To 18
        wsDB.Cells(rowFound, c).Value = wsMain.Range("F" & r).Value
        c = c + 1
    Next r
    
    '--- Block3: D21..D33 => columns V..AH ---
    c = 22 ' V=22
    For r = 21 To 33
        wsDB.Cells(rowFound, c).Value = wsMain.Range("D" & r).Value
        c = c + 1
    Next r
    
    '--- Block4: E21..E33 => columns AI..AU ---
    c = 35 ' AI=35
    For r = 21 To 33
        wsDB.Cells(rowFound, c).Value = wsMain.Range("E" & r).Value
        c = c + 1
    Next r
    
    '--- Block5: F21..F33 => columns AV..BH ---
    c = 48 ' AV=48
    For r = 21 To 33
        wsDB.Cells(rowFound, c).Value = wsMain.Range("F" & r).Value
        c = c + 1
    Next r
    
    '--- Block6: G21..G33 => columns BI..BU ---
    c = 61 ' BI=61
    For r = 21 To 33
        wsDB.Cells(rowFound, c).Value = wsMain.Range("G" & r).Value
        c = c + 1
    Next r
    
    '--- Block7: B6..B14 => columns BV..CD ---
    c = 74 ' BV=74
    For r = 6 To 14
        wsDB.Cells(rowFound, c).Value = wsMain.Range("B" & r).Value
        c = c + 1
    Next r
    
    '--- Block8: C6..C14 => columns CE..CM ---
    c = 83 ' CE=83
    For r = 6 To 14
        wsDB.Cells(rowFound, c).Value = wsMain.Range("C" & r).Value
        c = c + 1
    Next r
    
    MsgBox "Data saved successfully!", vbInformation
End Sub
