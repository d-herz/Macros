Option Explicit

Sub GenerateDES()
    ' -------------------------
    ' 1. SPEED SWITCHES ON (Crucial for performance)

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False ' Ensure no alerts interrupt the process
    End With

    On Error GoTo ErrorHandler ' Add error handling for clean shutdown

    ' --- Variable Declarations ---
    Dim wsItemList As Worksheet
    Dim wsProjectInfo As Worksheet
    Dim wsDES As Worksheet
    Dim lastItemRow As Long
    Dim itemRow As Long
    Dim itemCount As Long, routeCount As Long
    Dim i As Long, j As Long
    Dim currentCategoryInList As String
    Dim itemData() As Variant           ' Stores ItemList data
    Dim missingBreakouts As String
    Dim routeList() As String
    Dim breakoutTabName As String
    Dim desIndex As Long, sectionRow As Long, maxColumns As Long
    
    ' Variables for Category Header Logic (Not static)
    Dim colPtr As Long
    Dim catStartCol As Long
    Dim prevCategoryName As String
    Dim lastItemDESRow As Long
    
    ' Variables for Array Reading
    Dim ItemListRange As Range
    Dim ItemListData As Variant
    
    ' Named Constants for ItemList columns (for readability)
    Const COL_ITEM_NUM As Long = 2   ' B
    Const COL_A_FLAG As Long = 3     ' C
    Const COL_DESC As Long = 4       ' D
    Const COL_UNIT As Long = 5       ' E

    ' -------------------------
    ' Set references
    Set wsItemList = ThisWorkbook.Sheets("ItemList")
    Set wsProjectInfo = ThisWorkbook.Sheets("ProjectInfo")

    ' -------------------------
    ' Delete existing DES sheets
    For i = ThisWorkbook.Sheets.Count To 1 Step -1
        If Left(ThisWorkbook.Sheets(i).Name, 3) = "DES" Then
            ThisWorkbook.Sheets(i).Delete
        End If
    Next i

    ' -------------------------
    ' Get routes from ProjectRoutes (skip blanks)
    With wsProjectInfo.ListObjects("ProjectRoutes")
        routeCount = 0
        For i = 1 To .ListRows.Count
            If Trim(.DataBodyRange(i, 1).Value) <> "" Then routeCount = routeCount + 1
        Next i
        
        ' Handle case where routeCount is 0 to avoid ReDim error
        ReDim routeList(1 To IIf(routeCount = 0, 1, routeCount))
        Dim idx As Long: idx = 1
        For i = 1 To .ListRows.Count
            If Trim(.DataBodyRange(i, 1).Value) <> "" Then
                routeList(idx) = .DataBodyRange(i, 1).Value
                idx = idx + 1
            End If
        Next i
    End With
    
    Dim sectionHeight As Long
    sectionHeight = 8 + routeCount
    Dim section2Offset As Long
    section2Offset = sectionHeight + 1 ' One blank row

    ' -------------------------
    ' Optimized: Read ItemList data into an array
    lastItemRow = wsItemList.Cells(wsItemList.Rows.Count, COL_ITEM_NUM).End(xlUp).Row
    
    If lastItemRow < 7 Then
        MsgBox "The ItemList sheet contains no items.", vbInformation
        GoTo FinalCleanUp
    End If
    
    ' Read columns B through E (Item # through Unit) into the array
    Set ItemListRange = wsItemList.Range(wsItemList.Cells(7, COL_ITEM_NUM), wsItemList.Cells(lastItemRow, COL_UNIT))
    ItemListData = ItemListRange.Value
    
    itemCount = 0
    currentCategoryInList = ""
    
    ' Loop through the ItemListData array (faster than looping through cells)
    For itemRow = LBound(ItemListData, 1) To UBound(ItemListData, 1)
        Dim cellB As Variant: cellB = ItemListData(itemRow, 1) ' Column B (1st column in the array)
        Dim cellE As Variant: cellE = ItemListData(itemRow, 4) ' Column E (4th column in the array)

        ' CATEGORY ROW
        If Not IsNumeric(cellB) And Trim(CStr(cellB)) <> "" And Trim(CStr(cellE)) = "" Then
            currentCategoryInList = Trim(CStr(cellB))
            
        ' ITEM ROW (The fix from before)
        ElseIf IsNumeric(cellB) And currentCategoryInList <> "" Then
            If LCase(Trim(CStr(cellE))) <> "est." Then
                itemCount = itemCount + 1
                ReDim Preserve itemData(1 To 5, 1 To itemCount) ' Only need 5 columns now

                itemData(1, itemCount) = CStr(cellB)                                   ' Item #
                itemData(2, itemCount) = Trim(CStr(ItemListData(itemRow, 2)))          ' A flag (Col C/2nd array col)
                itemData(3, itemCount) = ItemListData(itemRow, 3)                      ' Item description (Col D/3rd array col)
                itemData(4, itemCount) = ItemListData(itemRow, 4)                      ' Unit (Col E/4th array col)
                itemData(5, itemCount) = currentCategoryInList                         ' Category
            End If
        End If
    Next itemRow

    If itemCount = 0 Then
        MsgBox "No valid items were found in the ItemList.", vbInformation
        GoTo FinalCleanUp
    End If

    ' -------------------------
    ' Generate DES sheets initialization
    desIndex = 1
    sectionRow = 1
    colPtr = 2
    maxColumns = 26 ' up to column Z

    ' Create first DES sheet
    Set wsDES = ThisWorkbook.Sheets.Add
    wsDES.Name = "DES_" & desIndex
    wsDES.Cells.Font.Name = "Calibri"

    ' Add ROW HEADERS for Section 1 dynamically
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
                ' ***Read entire breakout sheet into a single array ***
                Set wsBreakout = ThisWorkbook.Sheets(breakoutTabName)
                lastBreakoutRow = wsBreakout.Cells(wsBreakout.Rows.Count, "K").End(xlUp).Row
                
                If lastBreakoutRow >= 1 Then
                    ' Read columns K and L into a 2D array
                    BreakoutData = wsBreakout.Range("K1:L" & lastBreakoutRow).Value
                Else
                    ' Handle empty breakout sheet case
                    ReDim BreakoutData(1 To 1, 1 To 2)
                End If

                Dim rowOffset As Long
                If sectionRow = 1 Then rowOffset = 0 Else rowOffset = section2Offset

                ' --- CATEGORY HEADER LOGIC ---
                If i = 1 Or itemData(5, i) <> itemData(5, i - 1) Then
                    Call Finalize_Category_Header(wsDES, catStartCol, colPtr, prevCategoryName, rowOffset)
                    If catStartCol > 0 Then colPtr = colPtr + 1 ' Move column pointer after merging
                    
                    prevCategoryName = itemData(5, i)
                    catStartCol = colPtr
                End If

                ' --- Write DES column values (Cell-by-cell writing remains for complex formatting) ---
                wsDES.Cells(3 + rowOffset, colPtr).NumberFormat = "@"
                wsDES.Cells(3 + rowOffset, colPtr).Value = CStr(itemData(1, i))

                ' A-flag
                If LCase(Trim(itemData(2, i))) = "a" Then
                    wsDES.Cells(2 + rowOffset, colPtr).Value = "A"
                Else
                    wsDES.Cells(2 + rowOffset, colPtr).ClearContents
                End If

                ' Description
                wsDES.Cells(4 + rowOffset, colPtr).Value = itemData(3, i)
                wsDES.Cells(4 + rowOffset, colPtr).WrapText = True

                ' Unit
                With wsDES.Cells(5 + rowOffset, colPtr)
                    .Value = UCase(Trim(itemData(4, i)))
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Font.Bold = True
                End With

                ' Route quantities - Uses the fast array lookup function
                For j = 1 To routeCount
                    wsDES.Cells(5 + j + rowOffset, colPtr).Value = GetQuantityFromArray(BreakoutData, routeList(j) & " Subtotal")
                    wsDES.Cells(5 + j + rowOffset, colPtr).HorizontalAlignment = xlCenter
                Next j

                ' Project totals - Uses the fast array lookup function
                wsDES.Cells(6 + routeCount + rowOffset, colPtr).Value = GetQuantityFromArray(BreakoutData, "ProjectWide Subtotal")
                wsDES.Cells(7 + routeCount + rowOffset, colPtr).Value = GetQuantityFromArray(BreakoutData, "Unassigned")
                wsDES.Cells(8 + routeCount + rowOffset, colPtr).Value = GetQuantityFromArray(BreakoutData, "Total")
                wsDES.Range(wsDES.Cells(6 + routeCount + rowOffset, colPtr), wsDES.Cells(8 + routeCount + rowOffset, colPtr)).HorizontalAlignment = xlCenter

                ' Rotate headers
                With wsDES.Range(wsDES.Cells(2 + rowOffset, colPtr), wsDES.Cells(4 + rowOffset, colPtr))
                    .Orientation = 90
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With

                ' Borders for item block (Re-applies border for the whole range up to maxColumns)
                Dim blockBottom As Long
                blockBottom = 8 + routeCount + rowOffset
                With wsDES.Range(wsDES.Cells(2 + rowOffset, 2), wsDES.Cells(blockBottom, maxColumns))
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                End With

                ' Background for subtotal/unassigned/total (Re-applies background for the whole range up to maxColumns)
                wsDES.Range(wsDES.Cells(6 + routeCount + rowOffset, 2), wsDES.Cells(8 + routeCount + rowOffset, maxColumns)).Interior.Color = RGB(223, 227, 229)

                ' Next column
                colPtr = colPtr + 1

                ' --- Section transition / new sheet ---
                If colPtr > maxColumns Then
                    If sectionRow = 1 Then
                        ' Finalize header for section 1, start section 2
                        Call Finalize_Category_Header(wsDES, catStartCol, colPtr, prevCategoryName, 0)
                        colPtr = 2
                        sectionRow = 3
                        catStartCol = 0
                        prevCategoryName = ""
                        Call Add_Row_Headers(wsDES, routeList, routeCount, section2Offset, maxColumns)
                    Else
                        ' Finalize header for section 2, start new sheet
                        Call Finalize_Category_Header(wsDES, catStartCol, colPtr, prevCategoryName, section2Offset)
                        desIndex = desIndex + 1
                        Set wsDES = ThisWorkbook.Sheets.Add(After:=wsDES)
                        wsDES.Name = "DES_" & desIndex
                        wsDES.Cells.Font.Name = "Calibri"
                        Call Add_Row_Headers(wsDES, routeList, routeCount, 0, maxColumns)
                        colPtr = 2
                        sectionRow = 1
                        catStartCol = 0
                        prevCategoryName = ""
                    End If
                End If

            Else
                ' --- FLAG MISSING BREAKOUT ---
                missingBreakouts = missingBreakouts & vbCrLf & "• " & breakoutTabName
            End If
        End If
    Next i


    ' --- SHOW MISSING BREAKOUTS ---
    If Len(missingBreakouts) > 0 Then
        MsgBox "The following breakout tabs were not found:" & vbCrLf & missingBreakouts, vbExclamation
    End If


    ' -------------------------
    ' Final category merge
    Dim finalOffset As Long
    If sectionRow = 1 Then
        finalOffset = 0
    Else
        finalOffset = section2Offset
    End If

    ' Pass colPtr to Finalize_Category_Header to ensure the last category range is correct
    Call Finalize_Category_Header(wsDES, catStartCol, colPtr, prevCategoryName, finalOffset)

    ' -------------------------
    ' Determine last used row dynamically
    lastItemDESRow = wsDES.Cells(wsDES.Rows.Count, "B").End(xlUp).Row

    ' -------------------------
    ' ADD FOOTER dynamically after last item
    Call Add_Footer_Info(wsDES, wsProjectInfo, maxColumns, lastItemDESRow + 2)

    MsgBox "Detailed Estimate Sheets generated successfully!", vbInformation

FinalCleanUp:
    ' -------------------------
    ' 2. SPEED SWITCHES OFF (Important to reset)
    ' -------------------------
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    
    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error occurred (" & Err.Number & "): " & Err.Description, vbCritical
    Resume FinalCleanUp ' Jump to cleanup to re-enable application settings

End Sub

' --- Helper Subroutines and Functions ---

Sub Add_Row_Headers(wsDES As Worksheet, routeList() As String, routeCount As Long, offset As Long, maxColumns As Long)
    ' Adds and formats the row headers in Column A

    Dim j As Long
    
    ' Use a single array write for all headers in Column A
    Dim headers() As Variant
    ReDim headers(1 To 8 + routeCount, 1 To 1)
    
    headers(2, 1) = "A"
    headers(3, 1) = "Item Number"
    headers(4, 1) = "Item"
    headers(5, 1) = "Unit"

    For j = 1 To routeCount
        headers(5 + j, 1) = routeList(j)
    Next j

    headers(6 + routeCount, 1) = "Subtotal"
    headers(7 + routeCount, 1) = "Unassigned"
    headers(8 + routeCount, 1) = "Total"
    
    ' Write the array to Column A
    wsDES.Cells(1 + offset, 1).Resize(UBound(headers, 1), 1).Value = headers

    ' Autofit column A
    wsDES.Columns("A").EntireColumn.AutoFit

    ' Format row headers (kept as range operation for complex formatting)
    With wsDES.Range("A" & (2 + offset) & ":A" & (8 + routeCount + offset))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        ' Background for subtotal/unassigned/total
        wsDES.Range("A" & (6 + routeCount + offset) & ":A" & (8 + routeCount + offset)).Interior.Color = RGB(223, 227, 229)
    End With

    ' Apply borders across empty columns too
    With wsDES.Range(wsDES.Cells(2 + offset, 2), wsDES.Cells(8 + routeCount + offset, maxColumns))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

End Sub

Sub Finalize_Category_Header(wsDES As Worksheet, ByVal catStartCol As Long, ByVal colPtr As Long, ByVal prevCategoryName As String, offset As Long)
    ' This is called on category change or at the end of the macro
    If catStartCol > 0 Then
        ' Merge only if more than one column was used by the category
        If colPtr - catStartCol > 0 Then
            wsDES.Range(wsDES.Cells(1 + offset, catStartCol), wsDES.Cells(1 + offset, colPtr - 1)).Merge
        End If
        With wsDES.Cells(1 + offset, catStartCol)
            .Value = prevCategoryName
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End With
    End If
End Sub

Sub Add_Footer_Info(wsDES As Worksheet, wsProjectInfo As Worksheet, maxColumns As Long, Optional startRow As Long = 27)
    Dim wsSummaryCDM As Worksheet
    Dim i As Long

    ' Use Application.DisplayAlerts = False earlier, no need for On Error Resume Next here.
    Set wsSummaryCDM = Nothing
    On Error Resume Next
    Set wsSummaryCDM = ThisWorkbook.Sheets("SummaryCDM")
    On Error GoTo 0
    
    If wsSummaryCDM Is Nothing Then
        MsgBox "Error: SummaryCDM sheet not found. Cannot populate footer data.", vbCritical
        Exit Sub
    End If

    Dim dataRow As Long
    
    ' Insert blank row above footer
    wsDES.Rows(startRow).Insert Shift:=xlDown
    With wsDES.Range(wsDES.Cells(startRow, 1), wsDES.Cells(startRow, maxColumns))
        .Merge
        .Interior.Color = RGB(4, 117, 188)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ' Footer now starts at startRow + 1
    dataRow = startRow + 1
    
    ' Use a single array for column A text for minor optimization
    Dim ColAText(1 To 4) As String
    ColAText(1) = "State Project No."
    ColAText(2) = "Project Name"
    ColAText(3) = "District"
    ColAText(4) = "Towns"
    
    ' Use a Data array for column B+ values for minor optimization
    Dim ColBData(1 To 4) As Variant
    ColBData(1) = wsSummaryCDM.Range("C7").Value
    ColBData(2) = wsSummaryCDM.Range("B5").Value
    ColBData(3) = wsProjectInfo.Range("D10").Value
    ColBData(4) = wsSummaryCDM.Range("C9").Value
    
    ' Loop to populate and format footer rows
    For i = 1 To 4
        ' Column A (Header)
        wsDES.Cells(dataRow + i - 1, "A").Value = ColAText(i)
        
        ' Column B to maxColumns (Data)
        With wsDES.Range(wsDES.Cells(dataRow + i - 1, 2), wsDES.Cells(dataRow + i - 1, maxColumns))
            .Merge
            .Value = ColBData(i)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Size = IIf(i = 1, 12, 10) ' First row is size 12, others size 10
            .Font.Bold = True
            .WrapText = True ' Apply wrap text for first two rows
        End With
    Next i
    
    ' Re-calculate dataRow after loop
    dataRow = dataRow + 3
    
    ' Format headers in column A
    With wsDES.Range(wsDES.Cells(startRow + 1, 1), wsDES.Cells(dataRow, 1))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    SheetExists = False ' Default to False
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        SheetExists = True
    End If
End Function

' *** Optimized Lookup Function ***
Function GetQuantityFromArray(DataArray As Variant, label As String) As Double
    ' Searches a 2-column array (K and L) for a label in the first column, 
    ' returning the value from the second column.
    
    Dim k As Long
    GetQuantityFromArray = 0 ' Default to 0
    
    If Not IsArray(DataArray) Then Exit Function ' Handle uninitialized/single-value arrays
    
    On Error GoTo HandleError ' Trap errors if array bounds are invalid

    ' Loop through the rows of the 2D array
    For k = LBound(DataArray, 1) To UBound(DataArray, 1)
        ' Check the label in the first column of the array (was Column K)
        If Trim(CStr(DataArray(k, 1))) = label Then 
            ' Return the value from the second column of the array (was Column L)
            GetQuantityFromArray = DataArray(k, 2)
            Exit Function
        End If
    Next k
    
    Exit Function
    
HandleError:
    GetQuantityFromArray = 0 ' Return 0 on error
End Function