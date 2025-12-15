Option Explicit
' Define constant for the shift to the second section

Sub Generate_DES_121525()

    Dim wsItemList As Worksheet
    Dim wsProjectInfo As Worksheet
    Dim wsDES As Worksheet
    Dim lastItemRow As Long
    Dim itemRow As Long
    Dim itemCount As Long, routeCount As Long
    Dim i As Long, j As Long, colPtr As Long
    Dim currentCategoryInList As String
    Dim itemData() As Variant
    Dim missingBreakouts As String
    Dim routeList() As String
    Dim breakoutTabName As String
    Dim quantity As Double
    Dim desIndex As Long, sectionRow As Long, maxColumns As Long
    ' Variables for Category Header Logic
    Static catStartCol As Long
    Static prevCategoryName As String
    Dim lastItemDESRow As Long

    ' -------------------------
    ' Set references
    Set wsItemList = ThisWorkbook.Sheets("ItemList")
    Set wsProjectInfo = ThisWorkbook.Sheets("ProjectInfo")

    ' -------------------------
    ' Delete existing DES sheets
    For i = ThisWorkbook.Sheets.count To 1 Step -1
        If Left(ThisWorkbook.Sheets(i).Name, 3) = "DES" Then
            Application.DisplayAlerts = False
            ThisWorkbook.Sheets(i).Delete
            Application.DisplayAlerts = True
        End If
    Next i

    ' -------------------------
    ' Get routes from ProjectRoutes (skip blanks)
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
    section2Offset = sectionHeight + 1 ' One blank row



    ' -------------------------
    ' Count items in ItemList and store data
    lastItemRow = wsItemList.Cells(wsItemList.Rows.count, "B").End(xlUp).Row
    itemCount = 0
    currentCategoryInList = ""

    ' Row 6 is assumed to be the final header row for ItemList. Items start from Row 7.
    For itemRow = 7 To lastItemRow
        Dim cellB As Variant: cellB = wsItemList.Cells(itemRow, "B").Value
        Dim cellE As Variant: cellE = wsItemList.Cells(itemRow, "E").Value

        ' CATEGORY ROW
        If Not IsNumeric(cellB) And Trim(cellB) <> "" And Trim(cellE) = "" Then
            currentCategoryInList = Trim(cellB)
        ' ITEM ROW
        ElseIf IsNumeric(cellB) And currentCategoryInList <> "" Then
            ' Only process if Unit (Column E) is NOT "est." - this prevents storing placeholder rows
            If LCase(Trim(cellE)) <> "est." Then
                itemCount = itemCount + 1
                ReDim Preserve itemData(1 To 7, 1 To itemCount)

                itemData(1, itemCount) = CStr(wsItemList.Cells(itemRow, "B").Text) 'Item #
                itemData(2, itemCount) = Trim(wsItemList.Cells(itemRow, "C").Value) ' A flag
                itemData(3, itemCount) = wsItemList.Cells(itemRow, "D").Value ' Item description
                itemData(4, itemCount) = wsItemList.Cells(itemRow, "E").Value 'Unit (May be blank for missing breakouts)
                itemData(5, itemCount) = currentCategoryInList ' Category

            End If
        End If
    Next itemRow

    ' -------------------------
    ' RESET STATIC VARIABLES
    catStartCol = 0
    prevCategoryName = ""

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
' -------------------------
' Items Loop (rewritten)
missingBreakouts = ""  ' reset before loop
colPtr = 2
sectionRow = 1
catStartCol = 0
prevCategoryName = ""

For i = 1 To itemCount
    ' Only process valid item numbers
    If Trim(itemData(1, i)) <> "" Then

        ' Construct expected breakout tab name
        breakoutTabName = CStr(itemData(1, i))
        If LCase(Trim(itemData(2, i))) = "a" Then breakoutTabName = breakoutTabName & "A"
        breakoutTabName = Replace(breakoutTabName, " ", "") ' remove stray spaces

        ' Optional debug: print sheet names being checked
        Debug.Print "Checking sheet: '" & breakoutTabName & "'"

        If SheetExists(breakoutTabName) Then
            ' -------------------------
            ' Existing DES-writing logic
            Dim wsBreakout As Worksheet
            Set wsBreakout = ThisWorkbook.Sheets(breakoutTabName)

            Dim rowOffset As Long
            If sectionRow = 1 Then
                rowOffset = 0
            Else
                rowOffset = section2Offset
            End If

            ' --- CATEGORY HEADER LOGIC ---
            If i = 1 Or itemData(5, i) <> itemData(5, i - 1) Then
                If catStartCol > 0 And colPtr - catStartCol > 0 Then
                    wsDES.Range(wsDES.Cells(1 + rowOffset, catStartCol), wsDES.Cells(1 + rowOffset, colPtr - 1)).Merge
                    With wsDES.Cells(1 + rowOffset, catStartCol)
                        .Value = prevCategoryName
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .Font.Bold = True
                        With .Borders
                            .LineStyle = xlContinuous
                            .Weight = xlMedium
                        End With
                    End With
                    colPtr = colPtr + 1
                End If
                prevCategoryName = itemData(5, i)
                catStartCol = colPtr
            End If

            ' --- Write DES column values ---
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

            ' Route quantities
            For j = 1 To routeCount
                wsDES.Cells(5 + j + rowOffset, colPtr).Value = GetQuantityFromBreakout(wsBreakout, routeList(j) & " Subtotal")
                wsDES.Cells(5 + j + rowOffset, colPtr).HorizontalAlignment = xlCenter
            Next j

            ' Project totals
            wsDES.Cells(6 + routeCount + rowOffset, colPtr).Value = GetQuantityFromBreakout(wsBreakout, "ProjectWide Subtotal")
            wsDES.Cells(7 + routeCount + rowOffset, colPtr).Value = GetQuantityFromBreakout(wsBreakout, "Unassigned")
            wsDES.Cells(8 + routeCount + rowOffset, colPtr).Value = GetQuantityFromBreakout(wsBreakout, "Total")
            wsDES.Range(wsDES.Cells(6 + routeCount + rowOffset, colPtr), wsDES.Cells(8 + routeCount + rowOffset, colPtr)).HorizontalAlignment = xlCenter

            ' Rotate headers
            With wsDES.Range(wsDES.Cells(2 + rowOffset, colPtr), wsDES.Cells(4 + rowOffset, colPtr))
                .Orientation = 90
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With

            ' Borders for item block
            Dim blockBottom As Long
            blockBottom = 8 + routeCount + rowOffset
            With wsDES.Range(wsDES.Cells(2 + rowOffset, 2), wsDES.Cells(blockBottom, maxColumns))
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
            End With

            ' Background for subtotal/unassigned/total
            wsDES.Range(wsDES.Cells(6 + routeCount + rowOffset, 2), wsDES.Cells(8 + routeCount + rowOffset, maxColumns)).Interior.Color = RGB(223, 227, 229)

            ' Next column
            colPtr = colPtr + 1

            ' --- Section transition / new sheet ---
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

    End If ' Trim(itemData(1,i))
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

    Call Finalize_Category_Header(wsDES, catStartCol, colPtr, prevCategoryName, finalOffset)

    ' -------------------------
    ' Determine last used row dynamically
    lastItemDESRow = wsDES.Cells(wsDES.Rows.count, "B").End(xlUp).Row

    ' -------------------------
    ' ADD FOOTER dynamically after last item
    Call Add_Footer_Info(wsDES, wsProjectInfo, maxColumns, lastItemDESRow + 2)

    MsgBox "Detailed Estimate Sheets generated successfully!", vbInformation

End Sub


' --- Helper Subroutines and Functions ---

Sub Add_Row_Headers(wsDES As Worksheet, routeList() As String, routeCount As Long, offset As Long, maxColumns As Long)
    ' Adds and formats the row headers in Column A

    Dim j As Long
    wsDES.Cells(2 + offset, 1).Value = "A"
    wsDES.Cells(3 + offset, 1).Value = "Item Number"
    wsDES.Cells(4 + offset, 1).Value = "Item"
    wsDES.Cells(5 + offset, 1).Value = "Unit"

    For j = 1 To routeCount
        wsDES.Cells(5 + j + offset, 1).Value = routeList(j)
    Next j

    wsDES.Cells(6 + routeCount + offset, 1).Value = "Subtotal"
    wsDES.Cells(7 + routeCount + offset, 1).Value = "Unassigned"
    wsDES.Cells(8 + routeCount + offset, 1).Value = "Total"

    ' Autofit column A
    wsDES.Columns("A").EntireColumn.AutoFit

    ' Format row headers
    With wsDES.Range("A" & (2 + offset) & ":A" & (8 + routeCount + offset))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        ' Background for subtotal/unassigned/total
        .Interior.Color = RGB(223, 227, 229)
    End With

    ' Apply borders across empty columns too
    With wsDES.Range(wsDES.Cells(2 + offset, 2), wsDES.Cells(8 + routeCount + offset, maxColumns))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

End Sub

Sub Finalize_Category_Header(wsDES As Worksheet, ByVal catStartCol As Long, ByVal colPtr As Long, ByVal prevCategoryName As String, offset As Long)
    If catStartCol > 0 Then
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
    wsDES.Cells(dataRow, "A").Value = "State Project No."
    wsDES.Cells(dataRow, "A").WrapText = True
    With wsDES.Range(wsDES.Cells(dataRow, 2), wsDES.Cells(dataRow, maxColumns))
        .Merge
        .Value = wsSummaryCDM.Range("C7").Value
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Font.Size = 12
        .Font.Bold = True
    End With

    dataRow = dataRow + 1
    wsDES.Cells(dataRow, "A").Value = "Project Name"
    With wsDES.Range(wsDES.Cells(dataRow, 2), wsDES.Cells(dataRow, maxColumns))
        .Merge
        .Value = wsSummaryCDM.Range("B5").Value
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Font.Size = 10
        .Font.Bold = True
    End With

    dataRow = dataRow + 1
    wsDES.Cells(dataRow, "A").Value = "District"
    With wsDES.Range(wsDES.Cells(dataRow, 2), wsDES.Cells(dataRow, maxColumns))
        .Merge
        .Value = wsProjectInfo.Range("D10").Value
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Font.Size = 10
        .Font.Bold = True
    End With

    dataRow = dataRow + 1
    wsDES.Cells(dataRow, "A").Value = "Towns"
    With wsDES.Range(wsDES.Cells(dataRow, 2), wsDES.Cells(dataRow, maxColumns))
        .Merge
        .Value = wsSummaryCDM.Range("C9").Value
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Font.Size = 10
        .Font.Bold = True
    End With

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

    On Error Resume Next

    Set ws = ThisWorkbook.Sheets(sheetName)

    SheetExists = Not ws Is Nothing

    On Error GoTo 0

End Function

Function GetQuantityFromBreakout(ws As Worksheet, label As String) As Double

    Dim rng As Range

    Set rng = ws.Columns("K").Find(What:=label, LookIn:=xlValues, LookAt:=xlWhole)

    If Not rng Is Nothing Then

        GetQuantityFromBreakout = ws.Cells(rng.Row, "L").Value

    Else

        GetQuantityFromBreakout = 0

    End If

End Function




