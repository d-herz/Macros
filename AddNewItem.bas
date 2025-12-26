Option Explicit

Sub AddNewItem()

    Dim ws As Worksheet
    Dim itemNum As String
    Dim itemName As String
    Dim prefix As String
    Dim foundHeader As Range
    Dim insertRow As Long
    Dim i As Long
    Dim lastRow As Long
    Dim category As String
    Dim categoryMap As Object
    Dim nextRow As Long
    Dim firstItemRow As Long
    Dim key As Variant
    Dim found As Boolean

    On Error GoTo CleanExit

    '==============================
    ' ItemList setup
    '==============================
    Set ws = ThisWorkbook.Sheets("ItemList")
    ws.Unprotect

    '==============================
    ' Prompt for item number
    '==============================
    itemNum = InputBox( _
        "Enter the item number:" & vbCrLf & _
        "- Standard items: 7 digits (e.g., 0406196)" & vbCrLf & _
        "- Drainage items with depth: 7 digits + .## (e.g., 0586001.10)", _
        "Add New Item")

    If itemNum = "" Then GoTo CleanExit

    If Not itemNum Like "#######" And Not itemNum Like "#######.##" Then
        MsgBox "Invalid item number. Please enter a 7-digit number, optionally with a 2-digit suffix.", vbExclamation
        GoTo CleanExit
    End If

    prefix = Left(itemNum, 2)

    '==============================
    ' Category mapping
    '==============================
    Set categoryMap = CreateObject("Scripting.Dictionary")
    categoryMap.Add "Earthwork Items", Array("02", "03")
    categoryMap.Add "Roadway Items", Array("04")
    categoryMap.Add "Drainage Items", Array("05", "06")
    categoryMap.Add "Incidental Construction Items", Array("07", "08", "09")
    categoryMap.Add "Traffic Control Items", Array("10", "11", "12", "18")
    categoryMap.Add "Traffic Signal Items", Array("82")
    categoryMap.Add "Non-Contract Items", Array("30")

    category = ""
    For Each key In categoryMap.Keys
        If Not IsError(Application.Match(prefix, categoryMap(key), 0)) Then
            category = key
            Exit For
        End If
    Next key

    If category = "" Then
        MsgBox "Category not found for item prefix " & prefix, vbExclamation
        GoTo CleanExit
    End If

    '==============================
    ' Locate category section
    '==============================
    Set foundHeader = ws.Cells.Find(What:=category, LookIn:=xlValues, LookAt:=xlWhole)
    If foundHeader Is Nothing Then
        MsgBox "Could not find category header: " & category, vbCritical
        GoTo CleanExit
    End If

    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    nextRow = lastRow + 1

    For i = foundHeader.Row + 1 To lastRow
        If ws.Cells(i, "B").Value Like "*Items" Then
            nextRow = i
            Exit For
        End If
    Next i

    firstItemRow = foundHeader.Row + 3
    If firstItemRow >= nextRow Then
        MsgBox "No template row found under " & category & ".", vbCritical
        GoTo CleanExit
    End If

    '==============================
    ' Duplicate check
    '==============================
    For i = foundHeader.Row + 1 To nextRow - 1
        If ws.Cells(i, "B").Text = itemNum Then
            MsgBox "Item " & itemNum & " already exists in " & category & ".", vbExclamation
            GoTo CleanExit
        End If
    Next i

    '==============================
    ' Determine insertion row
    '==============================
    insertRow = nextRow
    For i = firstItemRow To nextRow - 1
        If ws.Cells(i, "B").Value <> "" Then
            If ws.Cells(i, "B").Value > itemNum Then
                insertRow = i
                Exit For
            End If
        End If
    Next i

    '==============================
    ' Insert new item row
    '==============================
    ws.Rows(insertRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ws.Rows(firstItemRow).Copy
    ws.Rows(insertRow).PasteSpecial xlPasteFormats
    ws.Rows(insertRow).PasteSpecial xlPasteFormulas

    ws.Rows(insertRow).Hidden = False
    ws.Cells(insertRow, "B").NumberFormat = "@"
    ws.Cells(insertRow, "B").Value = itemNum

    '==============================
    ' Lookup item description
    '==============================
    On Error Resume Next
    itemName = Application.WorksheetFunction.XLookup( _
        itemNum, _
        ThisWorkbook.Sheets("_MasterItemBidList").Columns("A"), _
        ThisWorkbook.Sheets("_MasterItemBidList").Columns("C"), _
        "")
    On Error GoTo 0

    If itemName = "" Then itemName = "Description Not Found"

    '==============================
    ' Create breakout sheet
    '==============================
    Dim breakoutTemplate As Worksheet
    Dim newBreakout As Worksheet
    Dim sheetName As String
    Dim originalVisibility As XlSheetVisibility

    sheetName = itemNum

    On Error Resume Next
    Set newBreakout = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If newBreakout Is Nothing Then

        Set breakoutTemplate = ThisWorkbook.Sheets("_ItemBreakoutTemplate")
        originalVisibility = breakoutTemplate.Visible
        breakoutTemplate.Visible = xlSheetVisible

        If breakoutTemplate.ProtectContents Then breakoutTemplate.Unprotect
        breakoutTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        breakoutTemplate.Protect UserInterfaceOnly:=True
        breakoutTemplate.Visible = originalVisibility

        Set newBreakout = ActiveSheet
        newBreakout.Unprotect

        newBreakout.Range("F6").Formula = _
            "=HYPERLINK(""#'ItemList'!B" & insertRow & """,""Go Back to Item List"")"

        newBreakout.name = sheetName

        '==============================
        ' Route section generation
        '==============================
        Dim routeTable As ListObject
        Dim namedRouteCount As Long
        Dim sectionsNeeded As Long

        On Error Resume Next
        Set routeTable = ThisWorkbook.Sheets("ProjectInfo").ListObjects("ProjectRoutes")
        On Error GoTo 0

        If Not routeTable Is Nothing Then
            namedRouteCount = Application.WorksheetFunction.CountA( _
                routeTable.ListColumns("Route").DataBodyRange)

            If namedRouteCount >= 2 Then
                sectionsNeeded = namedRouteCount - 1
                Call AddRouteSections_Rev(sectionsNeeded, newBreakout)
            End If
        End If

        newBreakout.Protect UserInterfaceOnly:=True
        Call SortItemBreakoutTabs(False)

    Else
        MsgBox "A breakout tab for item " & itemNum & " already exists.", vbExclamation
    End If

    MsgBox "Item #" & itemNum & " added under " & category & ".", vbInformation

CleanExit:
    Application.CutCopyMode = False
    ws.Protect UserInterfaceOnly:=True
    Call UpdateEstimateMetaData
    Call LogEstimateChange("Macro: AddNewItem", "Item: #" & itemNum & " " & itemName & " Added")

    ' --- Mark DES as out of date ---
    DESOutOfDate = True


End Sub


