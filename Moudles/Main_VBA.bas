Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    Application.EnableEvents = False
    
    ' SHIFT is in E5, DATE is in G5, OPERATOR is in D5 (not used directly here).
    ' 1) If SHIFT (E5) changed => copy tasks from "Tasks" into D10..D18 and F10..F18
    If Not Intersect(Target, Me.Range("E5")) Is Nothing Then
        
        ' Clear old tasks in D10..D18 and F10..F18
        Me.Range("D10:D18").ClearContents
        Me.Range("F10:F18").ClearContents
        
        Dim shiftName As String
        shiftName = LCase(Trim(Me.Range("E5").Value)) ' e.g. "morning", "friday", etc.
        
        Dim wsTasks As Worksheet
        Set wsTasks = ThisWorkbook.Sheets("Tasks") ' Must exist, spelled exactly "Tasks"
        
        Dim rowStart1 As Long, rowEnd1 As Long
        Dim rowStart2 As Long, rowEnd2 As Long
        
        Select Case shiftName
            Case "morning"
                ' MORNING => (A2..A10 -> D10..D18) and (A11..A19 -> F10..F18)
                rowStart1 = 2
                rowEnd1 = 10
                rowStart2 = 11
                rowEnd2 = 19
                
            Case "evening"
                ' EVENING => (A21..A29 -> D10..D18) and (A30..A38 -> F10..F18)
                rowStart1 = 21
                rowEnd1 = 29
                rowStart2 = 30
                rowEnd2 = 38
                
            Case "night"
                ' NIGHT => (A40..A48 -> D10..D18) and (A49..A57 -> F10..F18)
                rowStart1 = 40
                rowEnd1 = 48
                rowStart2 = 49
                rowEnd2 = 57
                
            Case "friday"
                ' FRIDAY => (A59..A67 -> D10..D18) and (A68..A76 -> F10..F18)
                rowStart1 = 59
                rowEnd1 = 67
                rowStart2 = 68
                rowEnd2 = 76
                
            Case "saturday"
                ' SATURDAY => (A78..A86 -> D10..D18) and (A87..A95 -> F10..F18)
                rowStart1 = 78
                rowEnd1 = 86
                rowStart2 = 87
                rowEnd2 = 95
                
            Case Else
                ' SHIFT not recognized => do nothing
                GoTo SafeExit
        End Select
        
        ' Copy first range => D10..D18
        Dim i As Long, destRow As Long
        destRow = 10
        For i = rowStart1 To rowEnd1
            Me.Range("D" & destRow).Value = wsTasks.Range("A" & i).Value
            destRow = destRow + 1
        Next i
        
        ' Copy second range => F10..F18
        destRow = 10
        Dim j As Long
        For j = rowStart2 To rowEnd2
            Me.Range("F" & destRow).Value = wsTasks.Range("A" & j).Value
            destRow = destRow + 1
        Next j
        
    End If
    
    ' 2) If DATE (G5) or SHIFT (E5) changed => call LoadShiftData
    If Not Intersect(Target, Me.Range("G5")) Is Nothing _
       Or Not Intersect(Target, Me.Range("E5")) Is Nothing Then
       
        LoadShiftData
    End If

SafeExit:
    Application.EnableEvents = True
End Sub


