Option Explicit


Public DESOutOfDate As Boolean

'===========================================================
' 1) BUTTON MACROS
'===========================================================

Sub PrintItemListToPDF()
    ExportToPDF ThisWorkbook.Sheets("ItemList"), "Estimate-ItemList", "ProjNumDOT"
    
    Call UpdateEstimateMetaData
    Call LogEstimateChange("Print: ItemList", "Item List exported to PDF")
End Sub

Sub PrintSummaryCDMToPDF()
    ExportToPDF ThisWorkbook.Sheets("SummaryCDM"), "CDM-Estimate-Summary", "ProjNumDOT"
    
    Call UpdateEstimateMetaData
    Call LogEstimateChange("Print: SummaryCDM", "CDM Summary exported to PDF")
End Sub

Sub PrintSummaryDOTToPDF()
    ExportToPDF ThisWorkbook.Sheets("SummaryDOT"), "DOT-Estimate-Summary", "ProjNumDOT"
    
    Call UpdateEstimateMetaData
    Call LogEstimateChange("Print: SummaryDOT", "DOT Summary exported to PDF")
End Sub

Sub PrintThisItemBreakout()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim itemName As String
    
    ' Validate it's a breakout sheet
    If Not IsNumeric(Left(ws.name, 1)) Then
        MsgBox "This does not appear to be an Item Breakout sheet.", vbExclamation
        Exit Sub
    End If
    
    itemName = Replace(ws.Range("C9").Value, " ", "-")
    ExportToPDF ws, ws.name & "_" & itemName, "ProjNumDOT"
    
    Call UpdateEstimateMetaData
    Call LogEstimateChange( _
        "Print: Item Breakout", _
        "Item: #" & ws.name & " " & Replace(ws.Range("C9").Value, vbCrLf, " ") & " exported to PDF" _
    )
End Sub

Sub PrintAll_CombinedPDF()
    Dim ws As Worksheet
    Dim outputSheets As Collection
    Dim sheetArray() As String
    Dim i As Integer
    
    Set outputSheets = New Collection
    
    ' 1) Sort breakout tabs (silent)
    Call SortItemBreakoutTabs(False)
    
    ' 2) Define standard sheets
    outputSheets.Add "SummaryCDM"
    outputSheets.Add "ItemList"
    
    ' 3) Add breakout sheets
    For Each ws In ThisWorkbook.Worksheets
        If IsNumeric(Left(ws.name, 7)) Then
            outputSheets.Add ws.name
        End If
    Next ws
    
    ' 4) Convert Collection to Array
    ReDim sheetArray(0 To outputSheets.count - 1)
    For i = 1 To outputSheets.count
        sheetArray(i - 1) = outputSheets(i)
    Next i
    
    ' 5) Export
    ExportToPDF sheetArray, "Cost-Estimate", "ProjNumDOT"
    
    Call UpdateEstimateMetaData
    Call LogEstimateChange( _
        "Print: Full Estimate", _
        "Summary, Item List, and all Item Breakouts exported to PDF" _
    )
End Sub

Sub ExportDEStoPDF()
    Dim ws As Worksheet
    Dim desSheetNames() As String
    Dim count As Long
    
    count = 0
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.name, 4) = "DES_" Then
            count = count + 1
            ReDim Preserve desSheetNames(0 To count - 1)
            desSheetNames(count - 1) = ws.name
        End If
    Next ws
    
    If count = 0 Then
        Dim userResponse As VbMsgBoxResult
        
        userResponse = MsgBox("No Detailed Estimate Sheets (DES) were found." & vbCrLf & vbCrLf & _
        "Would you like to generate the DES sheets now?", _
        vbQuestion + vbYesNo, _
        "Generate DES Sheets?")
        
        If userResponse = vbYes Then
            Call GenerateDES
            
            ' Rerun the export after generating the DES
            Call ExportDEStoPDF
        End If
            
        Exit Sub
    End If
    
    ExportToPDF desSheetNames, "DES", "ProjNumDOT"
    
    Call UpdateEstimateMetaData
    Call LogEstimateChange("Print: DES", "Detailed Estimate Sheets exported to PDF")
    
    On Error Resume Next
   ' ThisWorkbook.Sheets("DES_1").Select
End Sub



'===========================================================
' 2) Helper functions
'===========================================================

Private Sub ExportToPDF(sheetTarget As Variant, suffix As String, idRange As String)
    ' sheetTarget: Can be a Worksheet object OR an Array of sheet names
    ' suffix: The descriptive part of the filename
    ' idRange: The cell on ProjectInfo to look for the Project ID
    
    Dim projectID As String
    Dim fileName As String
    Dim fullPath As String
    Dim currentDate As String
    Dim originalSheet As Worksheet
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    ' Capture where the user started
    Set originalSheet = ActiveSheet
    
    ' Get Project ID using helper
    projectID = GetProjectID(idRange)
    currentDate = Format(Date, "mm-dd-yyyy")
    
    ' Build filename
    fileName = projectID & "_" & suffix & "_" & currentDate & ".pdf"
    fullPath = ThisWorkbook.Path & "\" & fileName
    
    ' Handle Multiple Sheets vs Single Sheet
    If IsArray(sheetTarget) Then
        ThisWorkbook.Sheets(sheetTarget).Select
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fullPath, OpenAfterPublish:=True

    Else
        sheetTarget.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fullPath, OpenAfterPublish:=True
    End If
    
CleanExit:
    If Not originalSheet Is Nothing Then originalSheet.Activate
    Application.ScreenUpdating = True
    
    MsgBox "PDF Created Successfully:" & vbCrLf & fileName, vbInformation, "Success"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    If Not originalSheet Is Nothing Then originalSheet.Activate
    MsgBox "Error creating PDF: " & Err.Description, vbCritical, "Error"
End Sub

Private Function GetProjectID(cellAddress As String) As String
    Dim val As Variant
    On Error Resume Next
    val = ThisWorkbook.Sheets("ProjectInfo").Range(cellAddress).Value
    On Error GoTo 0
    
    If Trim(val) = "" Then
        GetProjectID = "0000-0000"
    Else
        GetProjectID = CStr(val)
    End If
End Function
