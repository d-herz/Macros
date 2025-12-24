Option Explicit
' This macro is for physically ordering the item breakout tabs

Sub SortItemBreakoutTabs(Optional showMsg As Boolean = True, Optional restoreSheet As Boolean = True)
    Dim ws As Worksheet
    Dim itemSheets() As String
    Dim sortKeys() As Long
    Dim count As Long, i As Long, j As Long
    Dim tempName As String
    Dim tempKey As Long
    Dim numericPart As String
    Dim originalSheet As Worksheet
    
    
    ' 1) Optimization: Turn off "noise" that slows down execution
    On Error GoTo CleanUp
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set originalSheet = ActiveSheet


    ' 2) Capture item breakout sheet names
    ' Pre-size the array to total worksheets to avoid constant ReDim Preserve
    ReDim itemSheets(1 To ThisWorkbook.Worksheets.count)
    ReDim sortKeys(1 To ThisWorkbook.Worksheets.count)
    count = 0
    
    For Each ws In ThisWorkbook.Worksheets
        numericPart = ws.name
        If Right(numericPart, 1) = "A" Then numericPart = Left(numericPart, Len(numericPart) - 1)
        
        If IsNumeric(numericPart) Then
            count = count + 1
            itemSheets(count) = ws.name
            sortKeys(count) = CLng(numericPart)
        End If
    Next ws
    
    ' Exit if no sheets found
    If count = 0 Then GoTo CleanUp

    ' 3) Sort item sheets (Bubble Sort is fine for small/medium sheet counts)
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
    
    ' 4) Reorder the sheets physically
    ' Moving sheets is expensive; doing it once per sheet is the way to go
    Dim anchorSheet As String
    anchorSheet = "ItemList"
    
    For i = 1 To count
        ThisWorkbook.Sheets(itemSheets(i)).Move After:=ThisWorkbook.Sheets(anchorSheet)
        anchorSheet = itemSheets(i) ' The current sheet becomes the new anchor
    Next i

    ThisWorkbook.Sheets("ItemList").Activate

CleanUp:
    ' 5) Restore settings
    
    If restoreSheet Then
        If Not originalSheet Is Nothing Then originalSheet.Activate
    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    If Err.Number <> 0 Then
        MsgBox "An error occurred: " & Err.Description, vbCritical
    End If
    
    If showMsg Then MsgBox "Item breakout tabs sorted successfully!", vbInformation
    
End Sub
