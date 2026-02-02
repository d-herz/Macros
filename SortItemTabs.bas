' =============================================================================
' Module: SortItemTabs
' Author: DFH
' Created: January 2026
'
' Purpose:
' This module controls the physical ordering of item-related worksheets within
' the Standard Cost Estimate workbook. It ensures consistent, predictable sheet
' sequencing for both user navigation and downstream processes such as printing
' and combined PDF generation.
'
' Key Behaviors:
' - Numeric item breakout sheets are sorted in ascending numerical order
' - DES-related sheets (sheet names beginning with "DES") are grouped together
'   and positioned immediately after the ItemList sheet
' - Non-item and non-DES sheets are ignored
'
' Usage:
' - Can be executed directly via a button on the Dash and ItemList
' - Also invoked programmatically by PrintAll_CombinedPDF() to guarantee
'   correct sheet order prior to output
'
' Design Notes:
' - Sheet reordering is performed using worksheet.Move to physically reorder tabs
' - Workbook state (screen updating, calculation, events) is safely restored
'   even if an error occurs
' - Optional parameters allow UI messaging and active-sheet restoration to be
'   controlled by the calling context
'
' Dependencies:
' - Assumes the presence of an "ItemList" worksheet as the primary anchor
' - Uses SheetExists() helper function (defined in the UtilitiesModule)
' =============================================================================

Option Explicit

' Button Wrapper helper
Sub SortItemBreakoutTabs_Button()
    Call SortItemBreakoutTabs
End Sub

Sub SortItemBreakoutTabs(Optional showMsg As Boolean = True, Optional restoreSheet As Boolean = True)
    Dim ws As Worksheet
    Dim itemSheets() As String
    Dim sortKeys() As Long
    Dim count As Long, i As Long, j As Long
    Dim tempName As String, tempKey As Long
    Dim numericPart As String
    Dim originalSheet As Worksheet
    Dim desSheets As Collection
    Dim lastAnchor As String
    
    ' ------------------ Optimization ------------------
    On Error GoTo CleanUp
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set originalSheet = ActiveSheet
    
    ' ------------------ Identify numeric item breakout sheets ------------------
    ReDim itemSheets(1 To ThisWorkbook.Worksheets.count)
    ReDim sortKeys(1 To ThisWorkbook.Worksheets.count)
    count = 0
    
    Set desSheets = New Collection
    
    For Each ws In ThisWorkbook.Worksheets
        numericPart = ws.name
        
        ' Check for DES sheets (names starting with "DES")
        If UCase(Left(numericPart, 3)) = "DES" Then
            desSheets.Add ws.name
        Else
            ' Check for numeric sheets (including numeric + "A")
            If Right(numericPart, 1) = "A" Then numericPart = Left(numericPart, Len(numericPart) - 1)
            If IsNumeric(numericPart) Then
                count = count + 1
                itemSheets(count) = ws.name
                sortKeys(count) = CLng(numericPart)
            End If
        End If
    Next ws
    
    ' Exit if no numeric sheets found
    If count = 0 And desSheets.count = 0 Then GoTo CleanUp
    
    ' ------------------ Sort numeric item sheets ------------------
    If count > 1 Then
        For i = 1 To count - 1
            For j = i + 1 To count
                If sortKeys(j) < sortKeys(i) Then
                    tempKey = sortKeys(i): sortKeys(i) = sortKeys(j): sortKeys(j) = tempKey
                    tempName = itemSheets(i): itemSheets(i) = itemSheets(j): itemSheets(j) = tempName
                End If
            Next j
        Next i
    End If
    
    ' ------------------ Move DES sheets immediately after ItemList ------------------
    lastAnchor = "ItemList"
    For i = 1 To desSheets.count
        ThisWorkbook.Sheets(desSheets(i)).Move After:=ThisWorkbook.Sheets(lastAnchor)
        lastAnchor = desSheets(i)
    Next i
    
    ' ------------------ Move sorted numeric item breakout sheets ------------------
    For i = 1 To count
        ThisWorkbook.Sheets(itemSheets(i)).Move After:=ThisWorkbook.Sheets(lastAnchor)
        lastAnchor = itemSheets(i)
    Next i
    
    ' ------------------ Restore settings ------------------
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    If restoreSheet Then
        If Not originalSheet Is Nothing Then originalSheet.Activate
    End If
    
    If Err.Number <> 0 Then
        MsgBox "SortItemBreakoutTabs Failed: " & Err.Description, vbCritical
    ElseIf showMsg Then
        MsgBox "Item breakout tabs sorted successfully!", vbInformation
    End If
End Sub



