' ====================================================
' ThisWorkbook Module
' Tracks ItemBreakout changes, prevents deletion of protected sheets,
' sets zoom/page break preview, and updates LastUpdated metadata
' ====================================================

' --- Dictionary to store old values when a sheet is activated ---
' Must be declared at top-level so all subs can access it
Private SheetOldValues As Object

Private SummaryOldValues As Object

' ---------------- Protected sheet warning ----------------
Private Sub Workbook_SheetBeforeDelete(ByVal Sh As Object)
    Dim protectedSheets As Variant
    protectedSheets = Array("_ItemBreakoutTemplate", "_UnitPrices", "_MasterBidItemList") ' Add any sheets to protect

    Dim i As Long
    For i = LBound(protectedSheets) To UBound(protectedSheets)
        If Sh.name = protectedSheets(i) Then
            MsgBox "The sheet '" & Sh.name & "' is protected and cannot be deleted.", vbExclamation, "Protected Sheet"
            Application.EnableEvents = False
            Sh.Visible = xlSheetVisible
            Sh.Activate
            Application.EnableEvents = True
            Exit Sub
        End If
    Next i
End Sub

' ---------------- Workbook_Open ----------------
Private Sub Workbook_Open()
    Dim ws As Worksheet
    
    ' ----------- Initialize dictionary (for tracking and logging value changes) ------------
    If SheetOldValues Is Nothing Then
        Set SheetOldValues = CreateObject("Scripting.Dictionary")
    End If
    
    ' ----------- Existing zoom/page break preview ------------
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ActiveWindow.View = xlPageBreakPreview
        ActiveWindow.Zoom = 100
    Next ws
    
    Call CancelMetaAutoHide
    Call AutoHideMetaSheets
    
    ' Set Dash as the active sheet at the end
    Sheets("Dash").Activate
End Sub



' ---------------- Sheet Activate/Deactivate ----------------
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' Only track ItemBreakout sheets (numeric names)
    If IsItemBreakoutSheet(Sh.name) Then
        StoreOldValues Sh
        ' Summary CDM Sheet
    ElseIf Sh.name = "SummaryCDM" Then
        Call StoreSummaryOldValues(Sh)
    End If
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    ' Only track ItemBreakout sheets (numeric names)
    If IsItemBreakoutSheet(Sh.name) Then
        CompareAndLogChanges Sh
         'Summary CDM Sheet
    ElseIf Sh.name = "SummaryCDM" Then
        Call CompareSummaryChanges(Sh)
    End If
End Sub

' ================= Helper Functions =================
Private Function IsItemBreakoutSheet(sheetName As String) As Boolean
    ' Returns True if sheet name is numeric or numeric + "A" at the end
    Dim baseName As String
    Dim lastChar As String
    
    If Len(sheetName) = 0 Then
        IsItemBreakoutSheet = False
        Exit Function
    End If
    
    lastChar = Right(sheetName, 1)
    
    If lastChar = "A" Or lastChar = "a" Then
        baseName = Left(sheetName, Len(sheetName) - 1)
    Else
        baseName = sheetName
    End If
    
    IsItemBreakoutSheet = IsNumeric(baseName)
End Function

Private Sub StoreOldValues(ws As Worksheet)
    Dim subtotalCell As Range, unassignedCell As Range
    Dim oldValues As Collection
    Set oldValues = New Collection
    
    ' --- ProjectWide Subtotal ---
    Set subtotalCell = ws.Columns("K").Find(What:="ProjectWide Subtotal", LookAt:=xlWhole)
    If Not subtotalCell Is Nothing Then
        oldValues.Add subtotalCell.offset(0, 1).Value  ' Column L
    Else
        oldValues.Add Empty
    End If
    
    ' --- Unassigned ---
    Set unassignedCell = ws.Columns("K").Find(What:="Unassigned", LookAt:=xlWhole)
    If Not unassignedCell Is Nothing Then
        oldValues.Add unassignedCell.offset(0, -1).Value  ' Column J
    Else
        oldValues.Add Empty
    End If
    
    ' --- Store in dictionary ---
    If SheetOldValues Is Nothing Then Set SheetOldValues = CreateObject("Scripting.Dictionary")
    Set SheetOldValues(ws.name) = oldValues   
End Sub


Private Sub CompareAndLogChanges(ws As Worksheet)
    Dim subtotalCell As Range, unassignedCell As Range
    Dim oldValues As Collection
    Dim newSubtotal As Variant, newUnassigned As Variant
    
    If SheetOldValues Is Nothing Then Exit Sub
    If Not SheetOldValues.Exists(ws.name) Then Exit Sub
    Set oldValues = SheetOldValues(ws.name)
    
    ' --- ProjectWide Subtotal ---
    Set subtotalCell = ws.Columns("K").Find(What:="ProjectWide Subtotal", LookAt:=xlWhole)
    If Not subtotalCell Is Nothing Then
        newSubtotal = subtotalCell.offset(0, 1).Value  ' Column L
        If oldValues(1) <> Empty And newSubtotal <> oldValues(1) Then
            Application.EnableEvents = False
            Call LogEstimateChange("Manual Edit", _
                "Item: #" & ws.name & ", ProjectWide Subtotal changed from " & oldValues(1) & " to " & newSubtotal)
            Call UpdateEstimateMetaData
            Application.EnableEvents = True
        End If
    End If
    
    ' --- Unassigned ---
    Set unassignedCell = ws.Columns("K").Find(What:="Unassigned", LookAt:=xlWhole)
    If Not unassignedCell Is Nothing Then
        newUnassigned = unassignedCell.offset(0, -1).Value  ' Column J
        If oldValues(2) <> Empty And newUnassigned <> oldValues(2) Then
            Application.EnableEvents = False
            Call LogEstimateChange("Manual Edit", _
                "Item: #" & ws.name & ", Unassigned percentage changed from " & (oldValues(2) * 100) & "% to " & (newUnassigned * 100) & "%")
            Call UpdateEstimateMetaData
            Application.EnableEvents = True
        End If
    End If
End Sub


Private Sub StoreSummaryOldValues(ws As Worksheet)
    Dim nm As name
    ' Initialize dictionary
    If SummaryOldValues Is Nothing Then Set SummaryOldValues = CreateObject("Scripting.Dictionary")
    
    ' Loop through all named ranges to track
    Dim namesToTrack As Variant
    namesToTrack = Array("MinorItemAllowance", "OblYears", "InflFactor", "Contingencies", "Incdntl", "EstBy", "ChkdBy", "DevPhase")
    
    Dim n As Variant
    For Each n In namesToTrack
        On Error Resume Next
        SummaryOldValues(n) = ws.Range(n).Value
        On Error GoTo 0
    Next n
End Sub

Private Sub CompareSummaryChanges(ws As Worksheet)
    Dim n As Variant
    Dim oldValue As Variant, newValue As Variant
    Dim displayName As String
    Dim SummaryLabels As Object
    
    ' Exit if dictionary not initialized
    If SummaryOldValues Is Nothing Then Exit Sub
    
    ' Initialize labels dictionary
    Set SummaryLabels = CreateObject("Scripting.Dictionary")
    SummaryLabels("MinorItemAllowance") = "'Minor Item Allowance'"
    SummaryLabels("OblYears") = "Obligation Years"
    SummaryLabels("InflFactor") = "Inflation Factor"
    SummaryLabels("Contingencies") = "Contingencies"
    SummaryLabels("Incdntl") = "Incidentals"
    SummaryLabels("EstBy") = "'Estimated By:'"
    SummaryLabels("ChkdBy") = "'Checked By:'"
    SummaryLabels("DevPhase") = "Phase of Development"
    
    ' List of Named Ranges to track
    Dim namesToTrack As Variant
    namesToTrack = Array("MinorItemAllowance", "OblYears", "InflFactor", "Contingencies", "Incdntl", "EstBy", "ChkdBy", "DevPhase")
    
    Application.EnableEvents = False
    
    For Each n In namesToTrack
        On Error Resume Next
        oldValue = SummaryOldValues(n)
        newValue = ws.Range(n).Value
        On Error GoTo 0
        
        ' Only log if value changed
        If oldValue <> newValue Then
            
            ' Determine display name
            If SummaryLabels.Exists(n) Then
                displayName = SummaryLabels(n)
            Else
                displayName = n
            End If
            
            Dim logMsg As String
            
            ' Format numeric/percentage fields
            Select Case n
                Case "MinorItemAllowance", "Contingencies", "Incdntl", "InflFactor"
                    logMsg = "Sheet: SummaryCDM, " & displayName & " changed from " & Format(oldValue * 100, "0") & "% to " & Format(newValue * 100, "0") & "%"
                Case "OblYears"
                    logMsg = "Sheet: SummaryCDM, " & displayName & " changed from " & oldValue & " to " & newValue & " years"
                Case Else
                    logMsg = "Sheet: SummaryCDM, " & displayName & " changed from '" & oldValue & "' to '" & newValue & "'"
            End Select
            
            ' Log the change
            Call LogEstimateChange("Manual Edit", logMsg)
            Call UpdateEstimateMetaData
        End If
    Next n
    
    Application.EnableEvents = True
End Sub







' ---------------- Workbook_BeforeSave ----------------
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error GoTo SafeExit
    ThisWorkbook.Names("LastUpdatedBy").RefersToRange.Value = UserName()
    ThisWorkbook.Names("LastUpdatedOn").RefersToRange.Value = Now
SafeExit:
End Sub


