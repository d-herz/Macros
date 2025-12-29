Sub ValidateWorkbook()
    Dim ws As Worksheet
    Dim errReport As Worksheet
    Dim cell As Range
    Dim rng As Range
    Dim lastRow As Long
    Dim reportRow As Long
    Dim errorTypes As Variant
    Dim sheetName As String
    
    ' Define the errors to check
    errorTypes = Array(xlErrDiv0, xlErrNA, xlErrName, xlErrNull, xlErrNum, xlErrRef, xlErrValue)
    
    ' Delete previous _ErrorReport sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("_ErrorReport").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new _ErrorReport sheet
    Set errReport = ThisWorkbook.Sheets.Add
    errReport.name = "_ErrorReport"
    
    ' Add date, time, and username at the top
    errReport.Range("A1").Value = "Report Generated:"
    errReport.Range("B1").Value = Now
    errReport.Range("A2").Value = "User:"
    errReport.Range("B2").Value = Environ("USERNAME")
    
    ' Set headers starting at row 4
    errReport.Range("A4:D4").Value = Array("Sheet Name", "Cell Address", "Error Type", "Link")
    
    ' Format headers: bold, underline, center, thick bottom border
    With errReport.Range("A4:D4")
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .HorizontalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
    
    reportRow = 5 ' Start adding errors below headers
    
    ' Loop through sheets
    For Each ws In ThisWorkbook.Sheets
        sheetName = ws.name
        
        ' Skip MetaData and other ignored sheets
        If Left(sheetName, 1) <> "_" And _
           sheetName <> "UnitPrices" And _
           sheetName <> "_MetaData" And _
           sheetName <> "_ItemBreakoutTemplate" And _
           sheetName <> "_MasterItemBidList" Then
           
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
                        errReport.Cells(reportRow, 1).Value = ws.name
                        errReport.Cells(reportRow, 2).Value = cell.Address(False, False)
                        errReport.Cells(reportRow, 3).Value = cell.Text
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
    
    MsgBox "Validation complete. Check the '_ErrorReport' sheet for details.", vbInformation
End Sub


