Option Explicit

'=================================================================================
' LoadShiftData
'---------------------------------------------------------------------------------
' 1) Reads date (G5) + shift (E5)
' 2) Clears old data for blocks
' 3) If record found => overwrites Main sheet from DB
'    If no record => sets checkboxes B6..B14, C6..C14 to FALSE
' 4) Re-inserts formulas in E10..E18, G10..G18 referencing B6..B14, C6..C14
' 5) Colors "Done"/"Not Done" in E10..E18, G10..G18
'=================================================================================
Public Sub LoadShiftData()
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
    
    ' 1) Clear old data
    ' wsMain.Range("D10:D18").ClearContents
    ' wsMain.Range("F10:F18").ClearContents
    wsMain.Range("D21:D33").ClearContents
    wsMain.Range("E21:E33").ClearContents
    'wsMain.Range("F21:F33").ClearContents
    wsMain.Range("G21:G33").ClearContents
    
    wsMain.Range("B6:B14").ClearContents
    wsMain.Range("C6:C14").ClearContents
    
    ' 2) find row in DB
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
    
    If rowFound > 0 Then
        ' col C => operator => D5
        wsMain.Range("D5").Value = wsDB.Cells(rowFound, 3).Value
        
        Dim r As Long, c As Long
        
        '--- Block1 => D10..D18 => D..L (4..12)
        c = 4
        For r = 10 To 18
            wsMain.Range("D" & r).Value = wsDB.Cells(rowFound, c).Value
            c = c + 1
        Next r
        
        '--- Block2 => F10..F18 => M..U (13..21)
        c = 13
        For r = 10 To 18
            wsMain.Range("F" & r).Value = wsDB.Cells(rowFound, c).Value
            c = c + 1
        Next r
        
        '--- Block3 => D21..D33 => V..AH (22..34)
        c = 22
        For r = 21 To 33
            wsMain.Range("D" & r).Value = wsDB.Cells(rowFound, c).Value
            c = c + 1
        Next r
        
        '--- Block4 => E21..E33 => AI..AU (35..47)
        c = 35
        For r = 21 To 33
            wsMain.Range("E" & r).Value = wsDB.Cells(rowFound, c).Value
            c = c + 1
        Next r
        
        '--- Block5 => F21..F33 => AV..BH (48..60)
        c = 48
        For r = 21 To 33
            wsMain.Range("F" & r).Value = wsDB.Cells(rowFound, c).Value
            c = c + 1
        Next r
        
        '--- Block6 => G21..G33 => BI..BU (61..73)
        c = 61
        For r = 21 To 33
            wsMain.Range("G" & r).Value = wsDB.Cells(rowFound, c).Value
            c = c + 1
        Next r
        
        '--- Block7 => B6..B14 => BV..CD (74..82)
        c = 74
        For r = 6 To 14
            wsMain.Range("B" & r).Value = wsDB.Cells(rowFound, c).Value
            c = c + 1
        Next r
        
        '--- Block8 => C6..C14 => CE..CM (83..91)
        Dim c2 As Long
        c2 = 83
        For r = 6 To 14
            wsMain.Range("C" & r).Value = wsDB.Cells(rowFound, c2).Value
            c2 = c2 + 1
        Next r
        
    Else
        ' No row => checkboxes = FALSE
        Dim rr As Long
        For rr = 6 To 14
            wsMain.Range("B" & rr).Value = False
            wsMain.Range("C" & rr).Value = False
        Next rr
    End If
    
    ' 3) Re-insert formulas in E10..E18 => referencing B6..B14
    '                G10..G18 => referencing C6..C14
    Dim rowFormula As Long
    For rowFormula = 10 To 18
        Dim off As Long
        off = rowFormula - 4
        wsMain.Range("E" & rowFormula).Formula = "=IF($B$" & off & ",""Done"",""Not Done"")"
        wsMain.Range("G" & rowFormula).Formula = "=IF($C$" & off & ",""Done"",""Not Done"")"
    Next rowFormula
    
    ' 4) Color "Done"/"Not Done" in E10..E18, G10..G18
    Dim rowColor As Long
    For rowColor = 10 To 18
        Dim valE As String
        valE = wsMain.Range("E" & rowColor).Value
        Select Case valE
            Case "Done"
                wsMain.Range("E" & rowColor).Font.Color = RGB(0, 128, 0)
            Case "Not Done"
                wsMain.Range("E" & rowColor).Font.Color = RGB(255, 0, 0)
            Case Else
                wsMain.Range("E" & rowColor).Font.Color = RGB(0, 0, 0)
        End Select
        
        Dim valG As String
        valG = wsMain.Range("G" & rowColor).Value
        Select Case valG
            Case "Done"
                wsMain.Range("G" & rowColor).Font.Color = RGB(0, 128, 0)
            Case "Not Done"
                wsMain.Range("G" & rowColor).Font.Color = RGB(255, 0, 0)
            Case Else
                wsMain.Range("G" & rowColor).Font.Color = RGB(0, 0, 0)
        End Select
    Next rowColor
End Sub
