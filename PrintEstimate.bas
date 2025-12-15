' This macro is for printing the entire estimate to PDF (Summary Sheet, Item List, and Item Breakouts)
' It also physically sorts the Item Breakout tabs (so they print in order)
Sub PrintAll_CombinedPDF()

    Dim ws As Worksheet
    Dim itemSheets() As String
    Dim sortKeys() As Long
    Dim projectID As String
    Dim fileName As String
    Dim fullPath As String
    Dim currentDate As String
    Dim count As Long, i As Long, j As Long
    Dim tempName As String
    Dim tempKey As Long
    Dim outputSheets() As Variant
    Dim numericPart As String
    
    '====================
    ' 1) Capture item sheet names (including those with trailing "A")
    '====================
    count = 0
    For Each ws In ThisWorkbook.Worksheets
        numericPart = ws.Name
        ' Remove trailing "A" if exists
        If Right(numericPart, 1) = "A" Then numericPart = Left(numericPart, Len(numericPart) - 1)
        
        ' Check if the remaining part is numeric
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
        ThisWorkbook.Sheets(itemSheets(0)).Move After:=ThisWorkbook.Sheets("ItemList")
        For i = 1 To count - 1
            ThisWorkbook.Sheets(itemSheets(i)).Move After:=ThisWorkbook.Sheets(itemSheets(i - 1))
        Next i
    End If
    
    '====================
    ' 4) Build final print array
    '====================
    ReDim outputSheets(0 To count + 1)
    outputSheets(0) = "SummaryCDM"
    outputSheets(1) = "ItemList"
    For i = 0 To count - 1
        outputSheets(i + 2) = itemSheets(i)
    Next i
    
    '====================
    ' 5) File naming
    '====================
    On Error Resume Next
    projectID = ThisWorkbook.Sheets("ProjectInfo").Range("D6").Value
    On Error GoTo 0
    If Trim(projectID) = "" Then projectID = "0000-0000"
    
    currentDate = Format(Date, "mm-dd-yyyy")
    fileName = projectID & "_Cost-Estimate_" & currentDate & ".pdf"
    fullPath = ThisWorkbook.Path & "\" & fileName
    
    '====================
    ' 6) Export to PDF
    '====================
    Application.ScreenUpdating = False
    ThisWorkbook.Sheets(outputSheets).Select
    
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fullPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    Application.ScreenUpdating = True
    
    MsgBox "Combined PDF exported successfully!" & vbCrLf & vbCrLf & _
           "Saved To:" & vbCrLf & fullPath, vbInformation

End Sub