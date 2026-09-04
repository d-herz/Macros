' ====================================================
' Module: AddNewItemMod
' Author: DFH
' Created: January 2026
'
' Purpose:
' Adds a new bid item to the ItemList sheet and creates a corresponding Item Breakout worksheet based protected template.
'
' Key responsibilities:
' - Validate item number format
' - Determine item category from prefix (3-digit with 2-digit fallback)
' - Insert item in ItemList in sorted order within category
' - Prevent duplicate items
' - Create and configure breakout worksheet
' - Generate route sections if applicable
' - Log the change and update estimate metadata
'
' Assumptions:
' - ItemList category headers are in Column B
' - Item numbers are stored as text in Column B
' - Category sections are contiguous
' - "_ItemBreakoutTemplate" exists and is valid
' ====================================================

Option Explicit

Sub AddNewItem()

    FreezeUI
    On Error GoTo CleanExit

    Dim ws As Worksheet
    Dim itemNum As String
    Dim itemName As String
    Dim prefix2 As String
    Dim prefix3 As String
    Dim foundHeader As Range
    Dim insertRow As Long
    Dim i As Long
    Dim itemCreated As Boolean
    Dim lastRow As Long
    Dim category As String
    Dim categoryMap As Object
    Dim NextRow As Long
    Dim firstItemRow As Long
    Dim key As Variant
    Dim found As Boolean
    Dim originalSheet As Worksheet
    
    itemCreated = True
    Set originalSheet = ThisWorkbook.Sheets("ItemList")

    '==============================
    ' ItemList setup
    '==============================
    Set ws = ThisWorkbook.Sheets("ItemList")
    ws.Unprotect

    '==============================
    ' Prompt for item number
    ' Valid formats:
    '   - 7 digits (standard items)
    '   - 7 digits + ".##" (depth-based drainage items)
    '
    ' Item numbers are stored as text to preserve leading zeros.
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

    prefix2 = Left(itemNum, 2)
    prefix3 = Left(itemNum, 3)

    '==============================
    ' Category mapping
    '   Map item number prefixes to ItemList category headers.
    '   Explicit 3-digit prefixes are defined first to resolve overlaps 
    '   (e.g., Traffic Signals vs. Traffic Control in the 10/11 range).
    '   2-digit prefixes serve as fallbacks for non-overlapping categories.
    '==============================
    Set categoryMap = CreateObject("Scripting.Dictionary")
    
    ' Standard 2-digit categories
    categoryMap.Add "Earthwork Items", Array("02", "03")
    categoryMap.Add "Roadway Items", Array("04")
    categoryMap.Add "Drainage Items", Array("05", "06")
    categoryMap.Add "Incidental Construction Items", Array("07", "08", "09")
    categoryMap.Add "Non-Contract Items", Array("30")
    
    ' Overlapping 10/11/12/18 and 82 ranges separated into 3-digit & 2-digit rules
    categoryMap.Add "Traffic Signal Items", Array("100", "101", "102", "103", "104", "105", "106", "107", "108", "109", "110", "111", "82")
    categoryMap.Add "Traffic Control Items", Array("113", "114", "115", "116", "117", "118", "119", "12", "13" "18")

    category = ""
    
    ' --- Pass 1: Test 3-Digit Prefix Matches First ---
    For Each key In categoryMap.Keys
        If Not IsError(Application.Match(prefix3, categoryMap(key), 0)) Then
            category = key
            Exit For
        End If
    Next key

    ' --- Pass 2: Fall back to 2-Digit Prefix Matches if no 3-digit rule matched ---
    If category = "" Then
        For Each key In categoryMap.Keys
            If Not IsError(Application.Match(prefix2, categoryMap(key), 0)) Then
                category = key
                Exit For
            End If
        Next key
    End If

    If category = "" Then
        MsgBox "Category not found for item prefix " & prefix3 & " (or " & prefix2 & ").", vbExclamation
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
    NextRow = lastRow + 1

    For i = foundHeader.Row + 1 To lastRow
        If ws.Cells(i, "B").value Like "*Items" Then
            NextRow = i
            Exit For
        End If
    Next i

    firstItemRow = foundHeader.Row + 3
    If firstItemRow >= NextRow Then
        MsgBox "No template row found under " & category & ".", vbCritical
        GoTo CleanExit
    End If

    '==============================
    ' Duplicate check
    '==============================
    For i = foundHeader.Row + 1 To NextRow - 1
        If ws.Cells(i, "B").Text = itemNum Then
            MsgBox "Item " & itemNum & " already exists in " & category & ".", vbExclamation
            GoTo CleanExit
        End If
    Next i

    '==============================
    ' Determine insertion row
    '   Determine correct insertion row to maintain ascending sort order by item number within the category.
    '   Comparison is performed as text, relying on fixed-width item numbering to preserve numeric ordering.
    '==============================
    insertRow = NextRow
    For i = firstItemRow To NextRow - 1
        If ws.Cells(i, "B").value <> "" Then
            If ws.Cells(i, "B").value > itemNum Then
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
    ws.Cells(insertRow, "B").value = itemNum

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
    '   Create a new item breakout worksheet by copying the _ItemBreakoutTemplate
    '   Visibility and protection states are preserved and restored to avoid exposing internal/meta sheets
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

        newBreakout.Name = sheetName

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
                Call AddRouteSections(sectionsNeeded, newBreakout)
            End If
        End If

        newBreakout.Protect UserInterfaceOnly:=True
        
        Call SortItemBreakoutTabs(showMsg:=False, restoreSheet:=False) ' false, false is for not showing the msgbox and not restoring to the new breakout tab

    Else
        MsgBox "A breakout tab for item " & itemNum & " already exists.", vbExclamation
    End If

    MsgBox "Item #" & itemNum & " added under " & category & ".", vbInformation

CleanExit:
    Application.CutCopyMode = False
    ws.Protect UserInterfaceOnly:=True
    
    If itemCreated Then
        Call UpdateEstimateMetaData
        Call LogEstimateChange("Macro: AddNewItem", "Item: #" & itemNum & " " & itemName & " Added")
        ' --- Mark DES as out of date ---
        DESOutOfDate = True
    End If
    
    '------- Restore user back to original sheet (ItemList)----------
    originalSheet.Activate
    
    UnfreezeUI

End Sub