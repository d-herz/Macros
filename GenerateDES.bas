Option Explicit


Sub GenerateDES()
    ' -------------------------
    ' SPEED SWITCHES ON
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
        If Left(ThisWorkbook.Sheets(i).name, 3) = "DES" Then
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

    Set wsDES = ThisWorkbook.Sheets.Add(After:=wsItemList)
    wsDES.name = "DES_" & desIndex
    wsDES.Cells.Font.name = "Calibri"
    
    Call Apply_Print_Settings(wsDES)
    Call Add_Row_Headers(wsDES, routeList, routeCount, 0, maxColumns)
    Call Add_Row_Headers(wsDES, routeList, routeCount, section2Offset, maxColumns)

    ' -------------------------
    ' Items Loop
    ' -------------------------
   
    missingBreakouts = ""
    colPtr = 2
    sectionRow = 1
    catStartCol = 2
    prevCategoryName = ""
    
    Dim wsBreakout As Worksheet
    Dim BreakoutData As Variant
    Dim lastBreakoutRow As Long

    For i = 1 To itemCount
        ' --- SECTION/SHEET WRAP LOGIC ---
        If colPtr > maxColumns Then
            ' Finalize the category header for the row that just finished
            Dim wrapOffset As Long
            wrapOffset = IIf(sectionRow = 1, 0, section2Offset)
            Call Finalize_Category_Header(wsDES, catStartCol, colPtr, prevCategoryName, wrapOffset)
            
            If sectionRow = 1 Then
                ' Move from Section 1 to Section 2
                sectionRow = 2
                colPtr = 2
                catStartCol = 2
            Else
                ' Section 2 is full. ONLY create a new sheet if there are more items left to process.
                If i <= itemCount Then
                    ' Finish current sheet before moving to the next
                    Call Finalize_DES_Sheet(wsDES, wsProjectInfo, maxColumns, section2Offset, routeCount)
                    
                    desIndex = desIndex + 1
                    ' FIX: Added After:= parameter to keep sheets in order
                    Set wsDES = ThisWorkbook.Sheets.Add(After:=wsDES)
                    wsDES.name = "DES_" & desIndex
                    wsDES.Cells.Font.name = "Calibri"
                    
                    Call Apply_Print_Settings(wsDES)
                    Call Add_Row_Headers(wsDES, routeList, routeCount, 0, maxColumns)
                    Call Add_Row_Headers(wsDES, routeList, routeCount, section2Offset, maxColumns)
                    
                    sectionRow = 1
                    colPtr = 2
                    catStartCol = 2
                End If
            End If
        End If
        ' --------------------------------

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

                Dim currentOffset As Long
                currentOffset = IIf(sectionRow = 1, 0, section2Offset)

                ' Category header logic
                If i = 1 Or itemData(5, i) <> prevCategoryName Then
                    ' Finalize the previous category on this same row
                    If i > 1 Then
                        Call Finalize_Category_Header(wsDES, catStartCol, colPtr, prevCategoryName, currentOffset)
                    End If
                    
                    prevCategoryName = itemData(5, i)
                    catStartCol = colPtr
                End If

                ' Item placement (Standard logic continues...)
                wsDES.Cells(3 + currentOffset, colPtr).NumberFormat = "@"
                wsDES.Cells(3 + currentOffset, colPtr).Value = CStr(itemData(1, i))
                wsDES.Cells(2 + currentOffset, colPtr).Value = IIf(LCase(Trim(itemData(2, i))) = "a", "A", "")
                wsDES.Cells(4 + currentOffset, colPtr).Value = itemData(3, i)
                wsDES.Cells(4 + currentOffset, colPtr).WrapText = True

                With wsDES.Cells(5 + currentOffset, colPtr)
                    .Value = UCase(Trim(itemData(4, i)))
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Font.Bold = True
                End With

                For j = 1 To routeCount
                    wsDES.Cells(5 + j + currentOffset, colPtr).Value = GetQuantityFromArray(BreakoutData, routeList(j) & " Subtotal")
                    wsDES.Cells(5 + j + currentOffset, colPtr).HorizontalAlignment = xlCenter
                Next j

                wsDES.Cells(6 + routeCount + currentOffset, colPtr).Value = GetQuantityFromArray(BreakoutData, "ProjectWide Subtotal")
                wsDES.Cells(7 + routeCount + currentOffset, colPtr).Value = GetQuantityFromArray(BreakoutData, "Unassigned")
                wsDES.Cells(8 + routeCount + currentOffset, colPtr).Value = GetQuantityFromArray(BreakoutData, "Total")
                wsDES.Range(wsDES.Cells(6 + routeCount + currentOffset, colPtr), wsDES.Cells(8 + routeCount + currentOffset, colPtr)).HorizontalAlignment = xlCenter

                With wsDES.Range(wsDES.Cells(2 + currentOffset, colPtr), wsDES.Cells(4 + currentOffset, colPtr))
                    .Orientation = 90
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With

                colPtr = colPtr + 1
            Else
                missingBreakouts = missingBreakouts & vbCrLf & "• " & itemData(1, i) & ": " & itemData(3, i)
            End If
        End If
    Next i

    ' --- IMPORTANT: Finalize the very last category header after loop ends ---
    Dim finalOffset As Long
    finalOffset = IIf(sectionRow = 1, 0, section2Offset)
    Call Finalize_Category_Header(wsDES, catStartCol, colPtr, prevCategoryName, finalOffset)

    If Len(missingBreakouts) > 0 Then
        MsgBox "The following breakout tabs were not found:" & vbCrLf & missingBreakouts, vbExclamation
    End If

    ' Finalize last sheet

    Call Finalize_Category_Header(wsDES, catStartCol, colPtr, prevCategoryName, 0)
    
    ' Ensure the second section and footer are generated and formatted
    Call Finalize_DES_Sheet(wsDES, wsProjectInfo, maxColumns, section2Offset, routeCount)


    ' -------------------------
    ' FOCUS BACK TO FIRST SHEET
    ' -------------------------
    On Error Resume Next
    ThisWorkbook.Sheets("DES_1").Select
    On Error GoTo ErrorHandler

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
    ' Dimension array to hold all possible row labels
    ReDim headers(1 To 8 + routeCount, 1 To 1)
    
    headers(2, 1) = "A"
    headers(3, 1) = "Item Number"
    headers(4, 1) = "Item"
    headers(5, 1) = "Unit"

    For j = 1 To routeCount
        headers(5 + j, 1) = routeList(j)
    Next j

    ' Correct indices for the totals based on how many routes exist
    headers(6 + routeCount, 1) = "Subtotal"
    headers(7 + routeCount, 1) = "Unassigned"
    headers(8 + routeCount, 1) = "Total"
    
    ' Apply labels to Column A
    wsDES.Cells(1 + offset, 1).Resize(UBound(headers, 1), 1).Value = headers
    wsDES.Columns("A").EntireColumn.AutoFit

    ' Format column A
    With wsDES.Range("A" & (2 + offset) & ":A" & (8 + routeCount + offset))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Apply grey background to subtotal/total headers
    wsDES.Range("A" & (6 + routeCount + offset) & ":A" & (8 + routeCount + offset)).Interior.Color = RGB(223, 227, 229)

    ' Row Height Logic: Specifically target the "Item" row (4th row of section)
    ' If there is no item data in Column B, set height to 256
    If wsDES.Cells(4 + offset, 2).Value = "" Then
        wsDES.Rows(4 + offset).RowHeight = 256
    End If

    ' Grid borders for the data area
    With wsDES.Range(wsDES.Cells(2 + offset, 2), wsDES.Cells(8 + routeCount + offset, maxColumns))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    With wsDES.Range(wsDES.Cells(1 + offset, 1), wsDES.Cells(1 + offset, maxColumns)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    
End Sub

Sub Finalize_Category_Header(wsDES As Worksheet, ByVal catStartCol As Long, ByVal colPtr As Long, ByVal prevCategoryName As String, offset As Long)
    If catStartCol > 0 And colPtr > catStartCol Then
        Dim headerRange As Range
        Set headerRange = wsDES.Range(wsDES.Cells(1 + offset, catStartCol), wsDES.Cells(1 + offset, colPtr - 1))
        
        ' Merge first, then format the entire range
        headerRange.Merge
        
        With headerRange
            .Value = prevCategoryName
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            
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
    Dim section1Bottom As Long
    Dim section2Top As Long, section2Bottom As Long
    Dim footerTop As Long, footerBottom As Long

    ' Section 1 bottom = row before Section 2 headers
    section1Bottom = section2Offset - 1

    ' Section 2 top = first header row of section 2
    section2Top = section2Offset + 1

    ' Section 2 bottom = last item row in Section 2 including subtotal/unassigned/total
    section2Bottom = lastItemDESRow

    ' Footer = 2 rows after last item row (or adjust if needed)
    footerTop = lastItemDESRow + 1
    footerBottom = wsDES.Cells(wsDES.Rows.count, "B").End(xlUp).Row

    ' Apply outline borders
    ' Section 1
    Call Apply_Outline_Border(wsDES.Range(wsDES.Cells(1, 1), wsDES.Cells(section1Bottom, maxColumns)))

    ' Section 2
    Call Apply_Outline_Border(wsDES.Range(wsDES.Cells(section2Top, 1), wsDES.Cells(section2Bottom, maxColumns)))

    ' Footer
    Call Apply_Outline_Border(wsDES.Range(wsDES.Cells(footerTop, 1), wsDES.Cells(footerBottom, maxColumns)))
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

Sub Finalize_DES_Sheet(wsDES As Worksheet, wsProjectInfo As Worksheet, maxColumns As Long, section2Offset As Long, routeCount As Long)
    ' AutoFit the Item and Number rows for first section
    wsDES.Rows(3).EntireRow.AutoFit
    wsDES.Rows(4).EntireRow.AutoFit
    
    ' AutoFit the Item and Number rows for second section (only if they have content)
    If wsDES.Cells(3 + section2Offset, 2).Value <> "" Then wsDES.Rows(3 + section2Offset).EntireRow.AutoFit
    If wsDES.Cells(4 + section2Offset, 2).Value <> "" Then wsDES.Rows(4 + section2Offset).EntireRow.AutoFit

    ' Place footer immediately after the second section's "Total" row
    ' Section 2 ends at: section2Offset + 8 + routeCount
    Dim footerStartRow As Long
    footerStartRow = section2Offset + 8 + routeCount + 1

    Call Add_Footer_Info(wsDES, wsProjectInfo, maxColumns, footerStartRow)

    ' Apply bold borders: One for Section 1, and one for Section 2 + Footer combined
    Call Apply_Final_Borders(wsDES, maxColumns, section2Offset, footerStartRow)
End Sub

