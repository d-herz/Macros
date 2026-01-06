Option Explicit

' Dev_Mode should be set to True only when working on and testing this macro
' Flip to 'False' for production
Public Const DEV_MODE As Boolean = True

Private Const DEV_PROJECT_NUMBER As String = "DEV-TEST-0001"
Private Const DEV_FOLDER_NAME As String = "SCE_DEV"

Public Sub CreateNewProject()
    Dim projectNumber As String
    Dim baseFolder As String
    Dim desktopPath As String
    Dim newFilePath As String
    Dim newFileName As String
    Dim wbSource As Workbook
    Dim wbNew As Workbook
    
    Set wbSource = ThisWorkbook
    
    ' -------------------------
    ' Resolve Desktop Path
    ' -------------------------
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    
    ' -------------------------
    ' DEV MODE vs PRODUCTION
    ' -------------------------
    If DEV_MODE Then
        projectNumber = DEV_PROJECT_NUMBER
        baseFolder = desktopPath & "\" & DEV_FOLDER_NAME
    Else
        If Not ConfirmCreateNewProject() Then Exit Sub
        
        projectNumber = PromptForProjectNumber()
        If projectNumber = "" Then Exit Sub
        
        baseFolder = desktopPath & "\" & projectNumber
    End If
    
    ' -------------------------
    ' Ensure Folder Exists
    ' -------------------------
    EnsureFolderExists baseFolder
    
    ' -------------------------
    ' Build File Name
    ' -------------------------
    newFileName = projectNumber & "_Standard_Cost_Estimate.xlsm"
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
    
    ' Placeholder for next phase: XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    InitializeNewProject wbNew
    
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

' Central initializer sub

Public Sub InitializeNewProject(ByVal wb As Workbook)
    On Error GoTo CleanFail
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    ResetProjectInfo wb
    ResetProjectTowns wb
    ResetProjectRoutes wb

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






