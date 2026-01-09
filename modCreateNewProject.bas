Option Explicit

' Dev_Mode should be set to True only when working on and testing this macro
' Flip to 'False' for production
Public Const DEV_MODE As Boolean = False

Private Const DEV_PROJECT_NUMBER As String = "DEV-TEST-0001"
Private Const DEV_FOLDER_NAME As String = "SCE_DEV"

Public Sub CreateNewProject()
    Dim sourceProjectNumber As String
    Dim newProjectNumber As String
    Dim baseFolder As String
    Dim desktopPath As String
    Dim newFilePath As String
    Dim newFileName As String
    Dim wbSource As Workbook
    Dim wbNew As Workbook
    
    Set wbSource = ThisWorkbook
    
    sourceProjectNumber = GetProjectNumberFromWorkbook(wbSource)
    
    ' -------------------------
    ' Resolve Desktop Path
    ' -------------------------
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    
    ' -------------------------
    ' DEV MODE vs PRODUCTION
    ' -------------------------
    If DEV_MODE Then
        Debug.Print "WARNING: DEV_MODE is enabled"
        newProjectNumber = DEV_PROJECT_NUMBER
        baseFolder = desktopPath & "\" & DEV_FOLDER_NAME
    Else
        If Not ConfirmCreateNewProject() Then Exit Sub
        
        newProjectNumber = PromptForProjectNumber()
        If newProjectNumber = "" Then Exit Sub
        
        baseFolder = desktopPath & "\" & newProjectNumber
    End If
    
    ' -------------------------
    ' Ensure Folder Exists
    ' -------------------------
    EnsureFolderExists baseFolder
    
    ' -------------------------
    ' Build File Name
    ' -------------------------
    newFileName = newProjectNumber & "_Cost-Estimate.xlsm"
    newFilePath = baseFolder & "\" & newFileName
    
    ' -------------------------
    ' Handle Existing File
    ' -------------------------
    If FileExists(newFilePath) Then
        If DEV_MODE Then
            Kill newFilePath
        Else
            If Not ConfirmOverwrite(newFilePath) Then Exit Sub
            Kill newFilePath
        End If
    End If
    
    ' -------------------------
    ' Create Copy & Open
    ' -------------------------
    wbSource.SaveCopyAs newFilePath
    Set wbNew = Workbooks.Open(newFilePath)
    
    ' Call InitializeNewProject
    InitializeNewProject wbNew, sourceProjectNumber
    
    Call UpdateEstimateMetaData(ThisWorkbook)
    Call LogEstimateChange("Project Created", "New project file generated from this workbook")

    
    MsgBox "New project file created:" & vbCrLf & newFilePath, vbInformation
End Sub

'------------- Helper Functions --------------

' Confrimation prompt (not shown when "Dev Mode" is enabled)

Private Function ConfirmCreateNewProject() As Boolean
    ConfirmCreateNewProject = _
        MsgBox( _
            "This will create a new, empty Standard Cost Estimate file." & vbCrLf & _
            "The current file will NOT be modified." & vbCrLf & vbCrLf & _
            "Do you want to continue?", _
            vbQuestion + vbYesNo, _
            "Create New Project" _
        ) = vbYes
End Function

' Project number prompt
Private Function PromptForProjectNumber() As String
    Dim inputVal As String
    
    inputVal = InputBox("Enter the Project Number:", "New Project")
    inputVal = Trim(inputVal)
    
    PromptForProjectNumber = SanitizeFileName(inputVal)
End Function


Private Function ConfirmOverwrite(filePath As String) As Boolean
    ConfirmOverwrite = _
        MsgBox( _
            "The following file already exists:" & vbCrLf & _
            filePath & vbCrLf & vbCrLf & _
            "Do you want to overwrite it?", _
            vbExclamation + vbYesNo, _
            "File Exists" _
        ) = vbYes
End Function


Private Sub EnsureFolderExists(folderPath As String)
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
End Sub

Private Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function

Private Function SanitizeFileName(ByVal fileName As String) As String
    Dim badChars As Variant
    Dim i As Long
    
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    
    For i = LBound(badChars) To UBound(badChars)
        fileName = Replace(fileName, badChars(i), "_")
    Next i
    
    SanitizeFileName = fileName
End Function

' -----------------------Central initializer sub-----------------------------

Public Sub InitializeNewProject(ByVal wb As Workbook, Optional sourceProjectNumber As String = "")
    On Error GoTo CleanFail
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    ResetProjectInfo wb
    ResetProjectTowns wb
    ResetProjectRoutes wb
    DeleteItemBreakoutTabs wb
    ResetItemList wb
    ResetEstimateMetaData wb, sourceProjectNumber

CleanExit:
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    Exit Sub

CleanFail:
    MsgBox "Error during project initialization:" & vbCrLf & Err.Description, vbCritical
    Resume CleanExit
End Sub

' Reset ProjectInfo Named Cells

Private Sub ResetProjectInfo(ByVal wb As Workbook)
    SetNamedValue wb, "ProjNumDOT", "0000-0000"
    SetNamedValue wb, "ProjName", "(Project Name)"
    SetNamedValue wb, "CTDOTPM", "(Name)"
    SetNamedValue wb, "CTDOTDU", ""
    SetNamedValue wb, "CTDOTDistrict", ""
    
    SetNamedValue wb, "ProjNumCDM", "(CDM Project #)"
    SetNamedValue wb, "CDMPM", ""
    SetNamedValue wb, "CDMPTL", ""
    SetNamedValue wb, "CDMTaskLead", ""
    SetNamedValue wb, "CDMQAQC", ""
    
    SetNamedValue wb, "ProjType", ""
End Sub

' Helper for named ranges
Private Sub SetNamedValue(ByVal wb As Workbook, ByVal nameText As String, ByVal value As Variant)
    On Error Resume Next
    wb.Names(nameText).RefersToRange.value = value
    On Error GoTo 0
End Sub

' Reset ProjectTowns table (on ProjectInfo)
Private Sub ResetProjectTowns(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim townCol As Long
    
    Set ws = wb.Worksheets("ProjectInfo")
    Set lo = ws.ListObjects("ProjectTowns")
    
    townCol = lo.ListColumns("Town").Index
    
    ' Clear values only (preserve rows & formatting)
    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.Columns(townCol).ClearContents
    End If
    
    ' Seed first town
    lo.DataBodyRange.Cells(1, townCol).value = "Town #1"

End Sub
Private Sub ResetProjectRoutes(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim colRoute As Long
    Dim colStart As Long
    Dim colEnd As Long
    Dim rng As Range
    
    Set ws = wb.Worksheets("ProjectInfo")
    Set lo = ws.ListObjects("ProjectRoutes")
    
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set rng = lo.DataBodyRange
    
    ' Identify columns
    colRoute = lo.ListColumns("Route").Index
    colStart = lo.ListColumns("Start MP").Index
    colEnd = lo.ListColumns("End MP").Index
    
    ' Clear only user-entered columns
    rng.Columns(colRoute).ClearContents
    rng.Columns(colStart).ClearContents
    rng.Columns(colEnd).ClearContents
    
    ' Seed routes (defensive)
    If rng.Rows.count >= 1 Then
        rng.Cells(1, colRoute).value = "Route #1"
    End If
    
    If rng.Rows.count >= 2 Then
        rng.Cells(2, colRoute).value = "Route #2"
    End If
End Sub

' Helper sub for clearing out Item Breakout tabs
' Deletes all sheets EXCEPT the protected list, within the provided workbook

Private Sub DeleteItemBreakoutTabs(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim protectedSheets As Object
    Dim sheetName As String
    Dim i As Long

    ' Build protected sheet lookup
    Set protectedSheets = CreateObject("Scripting.Dictionary")
    protectedSheets.CompareMode = vbTextCompare

    protectedSheets("Dash") = True
    protectedSheets("_MetaData") = True
    protectedSheets("_ItemBreakoutTemplate") = True
    protectedSheets("_MasterItemBidList") = True
    protectedSheets("_UnitPrices") = True
    protectedSheets("ProjectInfo") = True
    protectedSheets("SummaryCDM") = True
    protectedSheets("SummaryDOT") = True
    protectedSheets("ItemList") = True

    ' Optional: treat _ErrorReport as deletable if present
    ' It should not be protected, but only attempt to delete if it exists in new workbook

    Application.DisplayAlerts = False

    ' Loop backwards when deleting sheets
    For i = wb.Worksheets.count To 1 Step -1
        Set ws = wb.Worksheets(i)
        sheetName = ws.name

        ' Delete any sheet that is NOT protected
        If Not protectedSheets.Exists(sheetName) Then
            On Error Resume Next     ' <-- safely ignore deletion errors
            If DEV_MODE Then Debug.Print "Deleting breakout tab in new workbook: " & sheetName
             ws.Visible = xlSheetVisible
            ws.Delete
            On Error GoTo 0
        End If
    Next i

    Application.DisplayAlerts = True
End Sub


' Logic for clearing out items from the ItemList tab
' Preserves:
'   - Category headers
'   - Subheaders
'   - Hidden template rows
'   - Category totals
'   - Spacer rows

Private Sub ResetItemList(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim itemVal As String
    Dim wasProtected As Boolean
    
    Set ws = wb.Worksheets("ItemList")
    
    Const headerEndRow As Long = 6 ' Row for stopping the bottom up loop (to prevent deletion of a main header row)
    
    ' -------------------------
    ' Handle protection
    ' -------------------------
    wasProtected = ws.ProtectContents
    If wasProtected Then
        ws.Unprotect    ' add Password:=... here if needed later
    End If
    
    On Error GoTo CleanFail
    
    ' Determine last used row based on Item No. column (B)
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    
    ' Loop bottom-up when deleting rows
    For r = lastRow To headerEndRow Step -1
        
        ' Skip hidden rows (template rows)
        If ws.Rows(r).Hidden Then GoTo NextRow
        
        ' Skip merged rows (category headers / totals)
        If ws.Cells(r, "B").MergeCells Then GoTo NextRow
        
        itemVal = Trim(ws.Cells(r, "B").value)
        
        ' Skip blanks and subheader row
        If itemVal = "" Then GoTo NextRow
        If itemVal = "Item No." Then GoTo NextRow
        
        ' Real item row ? delete
        If DEV_MODE Then
            Debug.Print "Deleting ItemList row " & r & _
                        " (Item No: " & itemVal & ")"
        End If
        
        ws.Rows(r).Delete
        
NextRow:
    Next r

CleanExit:
    ' Re-protect if needed
    If wasProtected Then
        ws.Protect     ' add Password:=... here if needed later
    End If
    Exit Sub

CleanFail:
    ' Ensure protection is restored even on error
    If wasProtected Then
        ws.Protect
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


' Clears the estimate change log in _MetaData and logs initial project creation
Public Sub ResetEstimateMetaData(Optional targetWB As Workbook = Nothing, Optional sourceProjectNumber As String = "")

    Dim wb As Workbook
    Dim wsMeta As Worksheet
    Dim logTable As ListObject
    Dim creationNote As String
    
    If targetWB Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetWB
    End If
    
    Set wsMeta = wb.Worksheets("_MetaData")
    Set logTable = wsMeta.ListObjects("TblChangeLog")
    
    ' Clear log
    If logTable.ListRows.count > 0 Then
        logTable.DataBodyRange.Delete
    End If
    
    ' Update metadata
    Call UpdateEstimateMetaData(wb)
    
    If sourceProjectNumber <> "" Then
        creationNote = "New project file created from Project #" & sourceProjectNumber
    Else
        creationNote = "New project file created from template"
    End If
    
    ' Log project creation
    Call LogEstimateChange("Project Created", creationNote, wb)

End Sub





