Option Explicit

Sub GenerateDES()
    ' -------------------------
    ' 1. SPEED SWITCHES ON
    ' -------------------------
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With

    On Error GoTo ErrorHandler

    ' --- Variable Declarations ---
    Dim wsItemList As Worksheet
    Dim wsProjectInfo As Worksheet
    Dim wsDES As Worksheet
    Dim lastItemRow As Long
    Dim itemRow As Long
    Dim itemCount As Long, routeCount As Long
    Dim i As Long, j As Long
    Dim currentCategoryInList As String
    Dim itemData() As Variant
    Dim missingBreakouts As String
    Dim routeList() As String
    Dim breakoutTabName As String
    Dim desIndex As Long, sectionRow As Long, maxColumns As Long
    
    Dim colPtr As Long
    Dim catStartCol As Long
    Dim prevCategoryName As String
    
    Dim ItemListRange As Range
    Dim ItemListData As Variant
    
    Const COL_ITEM_NUM As Long = 2   ' B
    Const COL_A_FLAG As Long = 3     ' C
    Const COL_DESC As Long = 4       ' D
    Const COL_UNIT As Long = 5       ' E

    ' -------------------------
    ' Set references
    ' -------------------------
    Set wsItemList = ThisWorkbook.Sheets("ItemList")
    Set wsProjectInfo = ThisWorkbook.Sheets("ProjectInfo")

    ' Delete existing DES sheets
    For i = ThisWorkbook.Sheets.count To 1 Step -1
        If Left(ThisWorkbook.Sheets(i).Name, 3) = "DES" Then
            ThisWorkbook.Sheets(i).Delete
        End If
    Next i

    ' Get routes
    With wsProjectInfo.ListObjects("ProjectRoutes")
        routeCount = 0
        For i = 1 To .ListRows.count
            If Trim(.DataBodyRange(i, 1).Value) <> "" Then routeCount = routeCount + 1
        Next i
        
        ReDim routeList(1 To IIf(routeCount = 0, 1, routeCount))
        Dim idx As Long: idx = 1
        For i = 1 To .ListRows.count
            If Trim(.DataBodyRange(i, 1).Value) <> "" Then
                routeList(idx) = .DataBodyRange(i, 1).Value
                idx = idx + 1
            End If
        Next i
    End With
    
    Dim sectionHeight As Long
    sectionHeight = 8 + routeCount
    Dim section2Offset As Long
    section2Offset = sectionHeight + 1

    ' Read ItemList
    lastItemRow = wsItemList.Cells(wsItemList.Rows.count, COL_ITEM_NUM).End(xlUp).Row
    
    If lastItemRow < 7 Then
        MsgBox "The ItemList sheet contains no items.", vbInformation
        GoTo FinalCleanUp
    End If
    
    Set ItemListRange = wsItemList.Range(wsItemList.Cells(7, COL_ITEM_NUM), wsItemList.Cells(lastItemRow, COL_UNIT))
    ItemListData = ItemListRange.Value
    
    itemCount = 0
    currentCategoryInList = ""
    
    For itemRow = LBound(ItemListData, 1) To UBound(ItemListData, 1)
        Dim cellB As Variant: cellB = ItemListData(itemRow, 1)
        Dim cellE As Variant: cellE = ItemListData(itemRow, 4)

        If Not IsNumeric(cellB) And Trim(CStr(cellB)) <> "" And Trim(CStr(cellE)) = "" Then
            currentCategoryInList = Trim(CStr(cellB))
        ElseIf IsNumeric(cellB) And currentCategoryInList <> "" Then
            If LCase(Trim(CStr(cellE))) <> "est." Then
                itemCount = itemCount + 1
                ReDim Preserve itemData(1 To 5, 1 To itemCount)
                itemData(1, itemCount) = CStr(cellB)
                itemData(2, itemCount) = Trim(CStr(ItemListData(itemRow, 2)))
                itemData(3, itemCount) = ItemListData(itemRow, 3)
                itemData(4, itemCount) = ItemListData(itemRow, 4)
                itemData(5, itemCount) = currentCategoryInList
            End If
        End If
    Next itemRow

    If itemCount = 0 Then
        MsgBox "No valid items were found in the ItemList.", vbInformation
        GoTo FinalCleanUp
    End If

    ' -------------------------
    ' Generate DES Initial Sheet
    ' -------------------------
    desIndex = 1
    sectionRow = 1
    colPtr = 2
    maxColumns = 26

    Set wsDES = ThisWorkbook.Sheets.Add
    wsDES.Name = "DES_" & desIndex
    wsDES.Cells.Font.Name = "Calibri"
    
    Call Apply_Print_Settings(wsDES)
    Call Add_Row_Headers(wsDES, routeList, routeCount, 0, maxColumns)

    ' -------------------------
    ' Items Loop
    ' -------------------------
    missingBreakouts = ""
    colPtr = 2
    sectionRow = 1
    catStartCol = 0
    prevCategoryName = ""
    
    Dim wsBreakout As Worksheet
    Dim BreakoutData As Variant
    Dim lastBreakoutRow As Long

    For i = 1 To itemCount
        If Trim(itemData(1, i)) <> "" Then
            breakoutTabName = CStr(itemData(1, i))
            If LCase(Trim(itemData(2, i))) = "a" Then breakoutTabName = breakoutTabName & "A"
            breakoutTabName = Replace(breakoutTabName, " ", "")

            If SheetExists(breakoutTabName) Then
                Set wsBreakout = ThisWorkbook.Sheets(breakoutTabName)
                lastBreakoutRow = wsBreakout.Cells(wsBreakout.Rows.count, "K").End(xlUp).Row
                
                If lastBreakoutRow >= 1 Then
                    BreakoutData = wsBreakout.Range("K1:L" & lastBreakoutRow).Value
                Else
                    ReDim BreakoutData(1 To 1, 1 To 2)
                End If

                Dim rowOffset As Long
                If sectionRow = 1 Then rowOffset = 0 Else rowOffset = section2Offset

                If i = 1 Or itemData(5, i) <> itemData(5, i - 1) Then
                    Call Finalize_Category_Header(wsDES, catStartCol, colPtr, prevCategoryName, rowOffset)
                    If catStartCol > 0 Then colPtr = colPtr + 1
                    prevCategoryName = itemData(5, i)
                    catStartCol = colPtr
                End If

                wsDES.Cells(3 + rowOffset, colPtr).NumberFormat = "@"
                wsDES.Cells(3 + rowOffset, colPtr).Value = CStr(itemData(1, i))

                If LCase(Trim(itemData(2, i))) = "a" Then
                    wsDES.Cells(2 + rowOffset, colPtr).Value = "A"
                Else
                    wsDES.Cells(2 + rowOffset, colPtr).ClearContents
                End If

                wsDES.Cells(4 + rowOffset, colPtr).Value = itemData(3, i)
                wsDES.Cells(4 + rowOffset, colPtr).WrapText = True

                With wsDES.Cells(5 + rowOffset, colPtr)
                    .Value = UCase(Trim(itemData(4, i)))
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Font.Bold = True
                End With

                For j = 1 To routeCount
                    wsDES.Cells(5 + j + rowOffset, colPtr).Value = GetQuantityFromArray(BreakoutData, routeList(j) & " Subtotal")
                    wsDES.Cells(5 + j + rowOffset, colPtr).HorizontalAlignment = xlCenter
                Next j

                wsDES.Cells(6 + routeCount + rowOffset, colPtr).Value = GetQuantityFromArray(BreakoutData, "ProjectWide Subtotal")
                wsDES.Cells(7 + routeCount + rowOffset, colPtr).Value = GetQuantityFromArray(BreakoutData, "Unassigned")
                wsDES.Cells(8 + routeCount + rowOffset, colPtr).Value = GetQuantityFromArray(BreakoutData, "Total")
                wsDES.Range(wsDES.Cells(6 + routeCount + rowOffset, colPtr), wsDES.Cells(8 + routeCount + rowOffset, colPtr)).HorizontalAlignment = xlCenter

                With wsDES.Range(wsDES.Cells(2 + rowOffset, colPtr), wsDES.Cells(4 + rowOffset, colPtr))
                    .Orientation = 90
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With

                With wsDES.Range(wsDES.Cells(2 + rowOffset, 2), wsDES.Cells(8 + routeCount + rowOffset, maxColumns))
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                End With

                wsDES.Range(wsDES.Cells(6 + routeCount + rowOffset, 2), wsDES.Cells(8 + routeCount + rowOffset, maxColumns)).Interior.Color = RGB(223, 227, 229)

                ' Row sizing for current item
                wsDES.Rows(3 + rowOffset).EntireRow.AutoFit
                wsDES.Rows(4 + rowOffset).EntireRow.AutoFit

                colPtr = colPtr + 1

                ' Section transition (for top and bottom rows of "items")
                If colPtr > maxColumns Then
                    If sectionRow = 1 Then
                        Call Finalize_Category_Header(wsDES, catStartCol, colPtr, prevCategoryName, 0)
                        colPtr = 2
                        sectionRow = 3
                        catStartCol = 0
                        prevCategoryName = ""
                        Call Add_Row_Headers(wsDES, routeList, routeCount, section2Offset, maxColumns)
                    Else
                        Call Finalize_Category_Header(wsDES, catStartCol, colPtr, prevCategoryName, section2Offset)
                        Call Finalize_DES_Sheet(wsDES, wsProjectInfo, maxColumns, section2Offset)
                        
                        desIndex = desIndex + 1
                        Set wsDES = ThisWorkbook.Sheets.Add(After:=wsDES)
                        wsDES.Name = "DES_" & desIndex
                        wsDES.Cells.Font.Name = "Calibri"
                        
                        Call Apply_Print_Settings(wsDES)
                        Call Add_Row_Headers(wsDES, routeList, routeCount, 0, maxColumns)
                        colPtr = 2
                        sectionRow = 1
                        catStartCol = 0
                        prevCategoryName = ""
                    End If
                End If
            Else
                ' --- FLAG MISSING BREAKOUT (Now with Item Name) ---
                missingBreakouts = missingBreakouts & vbCrLf & "• " & itemData(1, i) & ": " & itemData(3, i)
            End If
        End If
    Next i

    If Len(missingBreakouts) > 0 Then
        MsgBox "The following breakout tabs were not found:" & vbCrLf & missingBreakouts, vbExclamation
    End If

    ' Finalize last sheet
    Dim finalOffset As Long
    finalOffset = IIf(sectionRow = 1, 0, section2Offset)
    Call Finalize_Category_Header(wsDES, catStartCol, colPtr, prevCategoryName, finalOffset)
    Call Finalize_DES_Sheet(wsDES, wsProjectInfo, maxColumns, section2Offset)
    
    ' -------------------------
    ' FOCUS BACK TO FIRST SHEET
    ' -------------------------
    On Error Resume Next
    ThisWorkbook.Sheets("DES_1").Select
    On Error GoTo ErrorHandler

    ' MsgBox "Detailed Estimate Sheets generated successfully!", vbInformation

    ' Prompt user for PDF export
    Dim exportAnswer As VbMsgBoxResult
    exportAnswer = MsgBox("Detailed Estimate Sheets generated successfully!" & vbCrLf & vbCrLf & _
                          "Would you like to export them to PDF now?", vbYesNo + vbQuestion, "Export to PDF")
                          
    If exportAnswer = vbYes Then
        Call ExportDEStoPDF
    End If

FinalCleanUp:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error occurred (" & Err.Number & "): " & Err.Description, vbCritical
    Resume FinalCleanUp
End Sub

' --------------------------
' HELPER SUBS AND FUNCTIONS
'--------------------------

Sub Add_Row_Headers(wsDES As Worksheet, routeList() As String, routeCount As Long, offset As Long, maxColumns As Long)
    Dim j As Long
    Dim headers() As Variant
    ReDim headers(1 To 8 + routeCount, 1 To 1)
    
    headers(2, 1) = "A"
    headers(3, 1) = "Item Number"
    headers(4, 1) = "Item"
    headers(5, 1) = "Unit"

    For j = 1 To routeCount: headers(5 + j, 1) = routeList(j): Next j

    headers(6 + routeCount, 1) = "Subtotal"
    headers(7 + routeCount, 1) = "Unassigned"
    headers(8 + routeCount, 1) = "Total"
    
    wsDES.Cells(1 + offset, 1).Resize(UBound(headers, 1), 1).Value = headers
    wsDES.Columns("A").EntireColumn.AutoFit

    With wsDES.Range("A" & (2 + offset) & ":A" & (8 + routeCount + offset))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        wsDES.Range("A" & (6 + routeCount + offset) & ":A" & (8 + routeCount + offset)).Interior.Color = RGB(223, 227, 229)
        .EntireRow.AutoFit
    End With

    With wsDES.Range(wsDES.Cells(2 + offset, 2), wsDES.Cells(8 + routeCount + offset, maxColumns))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub

Sub Finalize_Category_Header(wsDES As Worksheet, ByVal catStartCol As Long, ByVal colPtr As Long, ByVal prevCategoryName As String, offset As Long)
    If catStartCol > 0 Then
        If colPtr - catStartCol > 0 Then wsDES.Range(wsDES.Cells(1 + offset, catStartCol), wsDES.Cells(1 + offset, colPtr - 1)).Merge
        With wsDES.Cells(1 + offset, catStartCol)
            .Value = prevCategoryName
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
        End With
    End If
End Sub

Sub Add_Footer_Info(wsDES As Worksheet, wsProjectInfo As Worksheet, maxColumns As Long, Optional startRow As Long = 27)
    Dim wsSummaryCDM As Worksheet, i As Long, dataRow As Long
    Set wsSummaryCDM = Nothing
    On Error Resume Next
    Set wsSummaryCDM = ThisWorkbook.Sheets("SummaryCDM")
    On Error GoTo 0
    If wsSummaryCDM Is Nothing Then Exit Sub

    wsDES.Rows(startRow).Insert Shift:=xlDown
    With wsDES.Range(wsDES.Cells(startRow, 1), wsDES.Cells(startRow, maxColumns))
        .Merge
        .Interior.Color = RGB(4, 117, 188)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    dataRow = startRow + 1
    Dim ColAText, ColBData
    ColAText = Array("State Project No.", "Project Name", "District", "Towns")
    ColBData = Array(wsSummaryCDM.Range("C7").Value, wsSummaryCDM.Range("B5").Value, wsProjectInfo.Range("D10").Value, wsSummaryCDM.Range("C9").Value)
    
    For i = 0 To 3
        wsDES.Cells(dataRow + i, "A").Value = ColAText(i)
        With wsDES.Range(wsDES.Cells(dataRow + i, 2), wsDES.Cells(dataRow + i, maxColumns))
            .Merge
            .Value = ColBData(i)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Size = IIf(i = 0, 12, 10)
            .Font.Bold = True
            .WrapText = True
        End With
        wsDES.Rows(dataRow + i).EntireRow.AutoFit
    Next i
    
    With wsDES.Range(wsDES.Cells(startRow + 1, 1), wsDES.Cells(dataRow + 3, 1))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Function GetQuantityFromArray(DataArray, label As String) As Double
    Dim k As Long
    GetQuantityFromArray = 0
    If Not IsArray(DataArray) Then Exit Function
    On Error GoTo HandleError
    For k = LBound(DataArray, 1) To UBound(DataArray, 1)
        If Trim(CStr(DataArray(k, 1))) = label Then
            GetQuantityFromArray = DataArray(k, 2)
            Exit Function
        End If
    Next k
HandleError:
End Function

Sub Apply_Outline_Border(targetRange As Range)
    With targetRange
        .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlThick
    End With
End Sub

Sub Apply_Final_Borders(wsDES As Worksheet, maxColumns As Long, lastItemDESRow As Long, section2Offset As Long)
    Dim finalBottomRow As Long
    finalBottomRow = lastItemDESRow + 2 + 4
    
    Call Apply_Outline_Border(wsDES.Range(wsDES.Cells(1, 1), wsDES.Cells(finalBottomRow, maxColumns)))
    Call Apply_Outline_Border(wsDES.Range(wsDES.Cells(1, 1), wsDES.Cells(section2Offset - 1, maxColumns)))
    
    If lastItemDESRow >= section2Offset Then
        Call Apply_Outline_Border(wsDES.Range(wsDES.Cells(section2Offset + 1, 1), wsDES.Cells(lastItemDESRow + 1, maxColumns)))
    End If
    Call Apply_Outline_Border(wsDES.Range(wsDES.Cells(lastItemDESRow + 2, 1), wsDES.Cells(finalBottomRow, maxColumns)))
End Sub

Sub Apply_Print_Settings(ws As Worksheet)
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .CenterHorizontally = True: .CenterVertically = True
    End With
End Sub

Sub Finalize_DES_Sheet(wsDES As Worksheet, wsProjectInfo As Worksheet, maxColumns As Long, section2Offset As Long)
    Dim lastItemDESRow As Long
    wsDES.Rows(3).EntireRow.AutoFit
    wsDES.Rows(4).EntireRow.AutoFit
    wsDES.Rows(3 + section2Offset).EntireRow.AutoFit
    wsDES.Rows(4 + section2Offset).EntireRow.AutoFit

    lastItemDESRow = wsDES.Cells(wsDES.Rows.count, "B").End(xlUp).Row
    If lastItemDESRow >= section2Offset - 1 Then
        Call Add_Footer_Info(wsDES, wsProjectInfo, maxColumns, lastItemDESRow + 2)
        Call Apply_Final_Borders(wsDES, maxColumns, lastItemDESRow, section2Offset)
    Else
        lastItemDESRow = wsDES.Cells(wsDES.Rows.count, "A").End(xlUp).Row
        If lastItemDESRow > 1 Then Call Apply_Outline_Border(wsDES.Range(wsDES.Cells(1, 1), wsDES.Cells(lastItemDESRow, maxColumns)))
    End If
End Sub

Sub ExportDEStoPDF()
    Dim ws As Worksheet
    Dim desSheetNames() As String
    Dim count As Long
    Dim projNum As String
    Dim fileName As String
    Dim folderPath As String
    Dim fullPath As String
    
    count = 0
    
    ' 1. Check if the workbook has been saved (otherwise Path is empty)
    folderPath = ThisWorkbook.Path
    If folderPath = "" Then
        MsgBox "Please save this Excel workbook first so the macro knows where to save the PDF.", vbExclamation
        Exit Sub
    End If
    
    ' 2. Collect all sheet names starting with "DES_"
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 4) = "DES_" Then
            count = count + 1
            ReDim Preserve desSheetNames(1 To count)
            desSheetNames(count) = ws.Name
        End If
    Next ws
    
    If count = 0 Then
        MsgBox "No Detailed Estimate Sheets (DES) found to export.", vbExclamation
        Exit Sub
    End If
    
    ' 3. Get Project Number and Construct Filename
    ' Logic: [ProjectNum]_DES_[MM-DD-YYYY].pdf
    projNum = Trim(ThisWorkbook.Sheets("ProjectInfo").Range("D6").Value)
    If projNum = "" Then projNum = "0000-0000" ' Fallback if empty
    
    fileName = projNum & "_DES_" & Format(Date, "mm-dd-yyyy") & ".pdf"
    fullPath = folderPath & "\" & fileName
    
    ' 4. Export the sheets
    On Error GoTo PDFError
    
    ' Select the group of sheets
    Sheets(desSheetNames).Select
    
    ' Export as single PDF
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                    fileName:=fullPath, _
                                    Quality:=xlQualityStandard, _
                                    IncludeDocProperties:=True, _
                                    IgnorePrintAreas:=False, _
                                    OpenAfterPublish:=True
    
    ' Return focus to DES_1 and deselect others
    Sheets("DES_1").Select
    
    MsgBox "PDF successfully created and saved to:" & vbCrLf & fullPath, vbInformation
    Exit Sub

PDFError:
    MsgBox "An error occurred while creating the PDF." & vbCrLf & _
           "Please check if a file with the same name is already open.", vbCritical
    On Error Resume Next
    Sheets("DES_1").Select
End Sub
