Option Explicit

'===========================================================
' 1) BUTTON MACROS
'===========================================================

Sub PrintItemListToPDF()
    ExportToPDF ThisWorkbook.Sheets("ItemList"), "Estimate-ItemList", "ProjNumDOT"
End Sub

Sub PrintSummaryCDMToPDF()
    ExportToPDF ThisWorkbook.Sheets("SummaryCDM"), "CDM-Estimate-Summary", "ProjNumDOT"
End Sub

Sub PrintSummaryDOTToPDF()
    ExportToPDF ThisWorkbook.Sheets("SummaryDOT"), "DOT-Estimate-Summary", "ProjNumDOT"
End Sub

Sub PrintThisItemBreakout()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim itemName As String
    
    ' Validate it's a breakout sheet (IsNumeric check)
    If Not IsNumeric(Left(ws.name, 1)) Then
        MsgBox "This does not appear to be an Item Breakout sheet.", vbExclamation
        Exit Sub
    End If
    
    itemName = Replace(ws.Range("C9").Value, " ", "-")
    ExportToPDF ws, ws.name & "_" & itemName, "ProjNumDOT"
End Sub

Sub PrintAll_CombinedPDF()
    Dim ws As Worksheet
    Dim outputSheets As Collection
    Dim sheetArray() As String
    Dim i As Integer
    
    Set outputSheets = New Collection
    
    ' 1) Sort (Passing False for no MessageBox)
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
    
    ' 4) Convert Collection to Array for the .Select method
    ReDim sheetArray(0 To outputSheets.count - 1)
    For i = 1 To outputSheets.count
        sheetArray(i - 1) = outputSheets(i)
    Next i
    
    ' 5) Export the array of sheets
    ExportToPDF sheetArray, "Cost-Estimate", "ProjNumDOT"
End Sub

Sub ExportDEStoPDF()
    Dim ws As Worksheet
    Dim desSheetNames() As String
    Dim count As Long
    
    count = 0
    ' 1. Collect all sheet names starting with "DES_"
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 4) = "DES_" Then
            count = count + 1
            ReDim Preserve desSheetNames(0 To count - 1)
            desSheetNames(count - 1) = ws.Name
        End If
    Next ws
    
    If count = 0 Then
        MsgBox "No Detailed Estimate Sheets (DES) found to export.", vbExclamation
        Exit Sub
    End If
    
    ' 2. We pass the array of sheet names, the suffix, and the Named Range for the ID
    ExportToPDF desSheetNames, "DES", "ProjNumDOT"
    
    ' 3. Clean up: Return focus to the first DES sheet
    On Error Resume Next
    ThisWorkbook.Sheets("DES_1").Select
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
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
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
        ThisWorkbook.Sheets("ItemList").Select ' Return to home base
    Else
        sheetTarget.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fullPath, OpenAfterPublish:=True
    End If
    
    Application.ScreenUpdating = True
    MsgBox "PDF Created Successfully:" & vbCrLf & fileName, vbInformation, "Success"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
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
