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
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("ItemList")

    ' Unprotect the sheet
    ws.Unprotect
    
    ' Prompt user for item number
    itemNum = InputBox("Enter the item number:" & vbCrLf & _
                   "- Standard items: 7 digits (e.g., 0406196)" & vbCrLf & _
                   "- Drainage items with depth: 7 digits + .## (e.g., 0586001.10)", _
                   "Add New Item")
    If itemNum = "" Then Exit Sub
    
    ' Validate input: accept 7 digits OR 7 digits + "." + 2 digits
    If Not itemNum Like "#######" And Not itemNum Like "#######.##" Then
        MsgBox "Invalid item number. Please enter a 7-digit number, optionally with a 2-digit suffix (e.g., 0586790 or 0586790.10).", vbExclamation
        Exit Sub
    End If

    
    prefix = Left(itemNum, 2)
    
    ' Create dictionary for category mapping
    Set categoryMap = CreateObject("Scripting.Dictionary")
    categoryMap.Add "Earthwork Items", Array("02", "03")
    categoryMap.Add "Roadway Items", Array("04")
    categoryMap.Add "Drainage Items", Array("05", "06")
    categoryMap.Add "Incidental Construction Items", Array("07", "08", "09")
    categoryMap.Add "Traffic Control Items", Array("10", "11", "12", "18")
    categoryMap.Add "Traffic Signal Items", Array("82")
    categoryMap.Add "Non-Contract Items", Array("30")
    
    ' Determine category based on prefix
    category = ""
    For Each key In categoryMap.Keys
        If Not IsError(Application.Match(prefix, categoryMap(key), 0)) Then
            category = key
            Exit For
        End If
    Next key
    
    If category = "" Then
        MsgBox "Category not found for item prefix " & prefix, vbExclamation
        Exit Sub
    End If
    
    ' Find category header
    Set foundHeader = ws.Cells.Find(What:=category, LookIn:=xlValues, LookAt:=xlWhole)
    If foundHeader Is Nothing Then
        MsgBox "Could not find category header: " & category, vbCritical
        Exit Sub
    End If

    ' Find next header (to define the end of the section)
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    nextRow = lastRow + 1
    For i = foundHeader.Row + 1 To lastRow
        If ws.Cells(i, "B").Value Like "*Items" Then
            nextRow = i
        Exit For
        End If
    Next i

' --- Skip 2 header rows under category header (for skipping "template" row) ---
firstItemRow = foundHeader.Row + 3
If firstItemRow >= nextRow Then
    MsgBox "No template row found under " & category & ".", vbCritical
    Exit Sub
End If

    
    ' Check for duplicates
    found = False
    For i = foundHeader.Row + 1 To nextRow - 1
        If ws.Cells(i, "B").Text = itemNum Then
            found = True
            Exit For
        End If
    Next i
    If found Then
        MsgBox "Item " & itemNum & " already exists in " & category & ".", vbExclamation
        Exit Sub
    End If
    
    ' Find insertion point (keep sorted by item number using string comparison)
    insertRow = nextRow
    For i = firstItemRow To nextRow - 1
        If ws.Cells(i, "B").Value <> "" Then
            If ws.Cells(i, "B").Value > itemNum Then
                insertRow = i
                Exit For
            End If
        End If
    Next i
    
    ' Insert new row
    ws.Rows(insertRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ' Copy formatting and formulas from template row (NOT the row above)
    ws.Rows(firstItemRow).Copy
    ws.Rows(insertRow).PasteSpecial Paste:=xlPasteFormats
    ws.Rows(insertRow).PasteSpecial Paste:=xlPasteFormulas
    
    ' Unhide new row (since all template rows hidden)
    ws.Rows(insertRow).Hidden = False
    
    ' Insert item number (preserve leading zeros)
    ws.Cells(insertRow, "B").NumberFormat = "@"
    ws.Cells(insertRow, "B").Value = itemNum
    
    ' Find Item Name from MasterItemBidList
    On Error Resume Next
    itemName = Application.WorksheetFunction.XLookup( _
                    itemNum, _
                    ThisWorkbook.Sheets("_MasterItemBidList").Columns("A"), _
                    ThisWorkbook.Sheets("_MasterItemBidList").Columns("C"), _
                    "")
    On Error GoTo 0

    If itemName = "" Then
        itemName = "Description Not Found"
    End If
    
    ' === Create Item Breakout Sheet ===
    Dim breakoutTemplate As Worksheet
    Dim newBreakout As Worksheet
    Dim sheetName As String
    Dim originalVisibility As XlSheetVisibility
    

    sheetName = itemNum  ' The item number becomes the sheet name

    ' Check if sheet already exists
    On Error Resume Next
    Set newBreakout = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If newBreakout Is Nothing Then
       ' Copy template
        Set breakoutTemplate = ThisWorkbook.Sheets("_ItemBreakoutTemplate")
        
        ' Store original visibility and temporarily unhide
        originalVisibility = breakoutTemplate.Visible
        breakoutTemplate.Visible = xlSheetVisible
        
        
        ' Unprotect Template only if needed — keeps user from messing with it
        If breakoutTemplate.ProtectContents Then
            breakoutTemplate.Unprotect
        End If
        
        ' Copy ItemBreakoutTemplate
        breakoutTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        
        ' Restore template visibility and Re-protect the template after copying
        breakoutTemplate.Protect UserInterfaceOnly:=True
        breakoutTemplate.Visible = originalVisibility
        
        ' Rename copied sheet
        Set newBreakout = ActiveSheet
        newBreakout.Protect UserInterfaceOnly:=True
        newBreakout.Unprotect
        
        ' === Update Back-to-ItemList hyperlink ===
        newBreakout.Range("F6").Formula = _
            "=HYPERLINK(""#'ItemList'!B" & insertRow & """, ""Go Back to Item List"")"

        ' Rename the new sheet
        On Error Resume Next
        newBreakout.name = sheetName
    
        If Err.Number <> 0 Then
            MsgBox "Could not name breakout sheet '" & sheetName & "'. " & vbCrLf & _
               "Check for invalid characters or duplicate names.", vbCritical
            Err.Clear
        End If
        
        ' Auto-sort item breakout tabs
        Call SortItemBreakoutTabs(False)
        On Error GoTo 0
        
        
    Else
         MsgBox "A breakout tab for item " & itemNum & " already exists.", vbExclamation
        
        
    End If
    
    
  
    
    Application.CutCopyMode = False
    MsgBox "Item #" & itemNum & " added under " & category & ".", vbInformation

    ws.Protect , UserInterfaceOnly:=True
    
    
      ' --- Update Last Updated in _MetaData
    Call UpdateEstimateMetaData
    
   
    
    ' Log the change in _MetaData
    Call LogEstimateChange("Macro: AddNewItem", "Item: #" & itemNum & " " & itemName & " Added")
    
End Sub



