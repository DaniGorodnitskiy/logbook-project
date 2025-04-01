Option Explicit

Public Sub CreateYesterdayStaticCopy()

    ' 1) Variable declarations
    Dim masterWB As Workbook
    Dim newWB As Workbook
    Dim dateStr As String
    Dim archivePath As String
    Dim newFilePath As String
    Dim linkArray As Variant
    Dim i As Long
    
    ' 2) Reference the current workbook (the "master" that holds this macro)
    Set masterWB = ThisWorkbook
    
    ' 3) Build the date string for "yesterday"
    '    If you prefer "today", just use Date without '- 1'
    dateStr = Format(Date - 1, "yyyy-mm-dd")  ' e.g. "2025-01-24"
    
    ' 4) Specify the folder path where you want to save the copy
    '    Make sure it ends with a backslash "\"
    archivePath = "C:\Users\DanielG\Tevel Metro\Noa Kirel - NOC\Daniel\logbook\GitHub\Archive"
    
    ' 5) Construct the full path for the new file
    '    If you want a macro-enabled copy, use ".xlsm"
    '    Otherwise, for a non-macro file, use ".xlsx"
    newFilePath = archivePath & "LogBookReport_" & dateStr & ".xlsm"
    
    ' Check if the archive folder exists; if not, create it
    If Dir(archivePath, vbDirectory) = "" Then
        MkDir archivePath
    End If
    
    ' 6) Save a copy of the current (master) workbook using the new file name
    masterWB.SaveCopyAs newFilePath
    
    ' 7) Turn off ScreenUpdating and Alerts for a cleaner user experience
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 8) Open the newly created copy to break external links
    Set newWB = Workbooks.Open(newFilePath)
    
    ' 9) Break any external links in the new workbook
    On Error Resume Next
    linkArray = newWB.LinkSources(Type:=xlExcelLinks)
    If Not IsEmpty(linkArray) Then
        For i = LBound(linkArray) To UBound(linkArray)
            newWB.BreakLink Name:=linkArray(i), Type:=xlLinkTypeExcelLinks
        Next i
    End If
    On Error GoTo 0
    
    ' 10) Close and save the new workbook (now with no external links)
    newWB.Close SaveChanges:=True
    
    ' 11) Restore screen updating and alerts
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' 12) Notify the user that the process is done
    MsgBox "A static copy (no external links) was created at:" & vbCrLf & newFilePath, vbInformation, "Success!"

End Sub



