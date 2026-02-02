'================================================================================
' Module: modValidatorTool
' Author: DFH
' Created: January 2026
'
' Purpose: Scans the workbook for errors (#REF!, #NAME?, #VALUE!, etc.) in
'          key estimate sheets and generates a consolidated "_ErrorReport" sheet.
'
' Key Functionality:
'   - Deletes any previous "_ErrorReport" sheet safely.
'   - Creates a new "_ErrorReport" sheet with headers, timestamp, and username.
'   - Loops through relevant sheets (ProjectInfo, SummaryDOT, SummaryCDM, ItemList,
'     and item breakout sheets) to identify cells with formula errors.
'   - Populates the report with Sheet Name, Cell Address, Error Type, and clickable
'     hyperlink to navigate directly to the cell.
'   - Autofits columns and logs the validation run in the estimate metadata.
'
' Notes / Assumptions:
'   - Ignores sheets starting with "_" (MetaData) and the "UnitPrices" sheet.
'   - Item breakout sheets are identified by names starting with numeric characters.
'   - Uses FreezeUI / UnfreezeUI to prevent screen flicker and speed up execution.
'================================================================================

Sub ValidateWorkbook()
    Dim ws As Worksheet
    Dim errReport As Worksheet
    Dim cell As Range
    Dim rng As Range
    Dim reportRow As Long
    Dim sheetName As String
    Dim erSheetExists As Boolean
    Dim errorCount As Long
    
    FreezeUI
    On Error GoTo CleanExit
    
    ' -------------------------
    ' Delete previous _ErrorReport sheet safely
    ' -------------------------
    erSheetExists = False
    If SheetExists("_ErrorReport") Then
        Set errReport = ThisWorkbook.Sheets("_ErrorReport")
        erSheetExists = True
        ' Unhide if very hidden so it can be deleted
        If errReport.Visible = xlSheetVeryHidden Then errReport.Visible = xlSheetVisible
        ' Delete the sheet
        Application.DisplayAlerts = False
        errReport.Delete
        Application.DisplayAlerts = True
    End If
    
    ' -------------------------
    ' Create new _ErrorReport sheet
    ' -------------------------
    Set errReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Dash"))
    errReport.name = "_ErrorReport"
    
    ' Add date, time, and username at the top
    errReport.Range("A1").value = "Report Generated:"
    errReport.Range("B1").value = Now
    errReport.Range("A2").value = "User:"
    errReport.Range("B2").value = Environ("USERNAME")
    
    ' Set headers starting at row 4
    errReport.Range("A4:D4").value = Array("Sheet Name", "Cell Address", "Error Type", "Link")
    
    ' Format headers: bold, underline, center, thick bottom border
    With errReport.Range("A4:D4")
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .HorizontalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
    
    reportRow = 5 ' Start adding errors below headers
    
    ' -------------------------
    ' Loop through sheets
    ' -------------------------
    For Each ws In ThisWorkbook.Sheets
        sheetName = ws.name
        
        ' Skip MetaData and other ignored sheets
        If Left(sheetName, 1) <> "_" And _
           sheetName <> "UnitPrices" Then
           
            ' Include only relevant sheets: ProjectInfo, SummaryDOT, SummaryCDM, ItemList, or item breakout sheets
            If sheetName = "ProjectInfo" Or sheetName = "SummaryDOT" Or sheetName = "SummaryCDM" Or sheetName = "ItemList" _
               Or sheetName Like "[0-9]*" Then
               
                ' Check used range for errors
                On Error Resume Next
                Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
                On Error GoTo 0
                
                If Not rng Is Nothing Then
                    For Each cell In rng
                        ' Add details to report
                        errReport.Cells(reportRow, 1).value = ws.name
                        errReport.Cells(reportRow, 2).value = cell.Address(False, False)
                        errReport.Cells(reportRow, 3).value = cell.value
                        ' Add hyperlink to cell
                        errReport.Hyperlinks.Add Anchor:=errReport.Cells(reportRow, 4), _
                            Address:="", SubAddress:="'" & ws.name & "'!" & cell.Address, _
                            TextToDisplay:="Go To"
                        reportRow = reportRow + 1
                    Next cell
                End If
                
                Set rng = Nothing
            End If
        End If
    Next ws
    
    ' Autofit columns for readability
    errReport.Columns("A:D").AutoFit
    
    
    errorCount = reportRow - 5
    
    Call UpdateEstimateMetaData
    Call LogEstimateChange("Error Report", "Error Report Generated, " & errorCount & " error(s) found")
    
    MsgBox "Validation complete. Check the '_ErrorReport' sheet for details.", vbInformation
    
CleanExit:
    UnfreezeUI
    
End Sub


