' This Macro is for physically sorting the item breakout tabs in order

Sub SortItemBreakoutTabs()

    Dim ws As Worksheet
    Dim itemSheets() As String
    Dim sortKeys() As Long
    Dim count As Long, i As Long, j As Long
    Dim tempName As String
    Dim tempKey As Long
    Dim numericPart As String

    '====================
    ' 1) Capture item breakout sheet names (numeric or numeric + "A")
    '====================
    count = 0
    For Each ws In ThisWorkbook.Worksheets
        numericPart = ws.Name
        ' Remove trailing "A" if exists
        If Right(numericPart, 1) = "A" Then numericPart = Left(numericPart, Len(numericPart) - 1)
        
        ' Check if remaining part is numeric
        If IsNumeric(numericPart) Then
            ReDim Preserve itemSheets(count)
            ReDim Preserve sortKeys(count)
            itemSheets(count) = ws.Name
            sortKeys(count) = CLng(numericPart)
            count = count + 1
        End If
    Next ws
    
    '====================
    ' 2) Sort item sheets by numeric value
    '====================
    If count > 1 Then
        For i = 0 To count - 2
            For j = i + 1 To count - 1
                If sortKeys(j) < sortKeys(i) Then
                    ' Swap keys
                    tempKey = sortKeys(i)
                    sortKeys(i) = sortKeys(j)
                    sortKeys(j) = tempKey
                    ' Swap sheet names
                    tempName = itemSheets(i)
                    itemSheets(i) = itemSheets(j)
                    itemSheets(j) = tempName
                End If
            Next j
        Next i
    End If
    
    '====================
    ' 3) Reorder the sheets in workbook
    '====================
    If count > 0 Then
        ' Move the first item sheet after "ItemList"
        ThisWorkbook.Sheets(itemSheets(0)).Move After:=ThisWorkbook.Sheets("ItemList")
        ' Move remaining item sheets after the previous sorted sheet
        For i = 1 To count - 1
            ThisWorkbook.Sheets(itemSheets(i)).Move After:=ThisWorkbook.Sheets(itemSheets(i - 1))
        Next i
    End If

    '====================
    ' 4) Return to ItemList tab
    '====================
    ThisWorkbook.Sheets("ItemList").Activate


    MsgBox "Item breakout tabs sorted successfully!", vbInformation

End Sub
