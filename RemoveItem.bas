Option Explicit

' This macro is for removing an item from the "ItemList".
' It will also ask the user if they want to delete the associated item breakout tab (if it exists)

Sub RemoveItem()
    Dim ws As Worksheet
    Dim itemNum As String
    Dim itemName As String
    Dim suffix As String
    Dim fullSheetName As String
    Dim itemRow As Long
    Dim lastRow As Long
    Dim found As Boolean
    Dim breakoutSheet As Worksheet
    Dim deleteBreakout As VbMsgBoxResult
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("ItemList")
    
    ' Prompt user for item number
    itemNum = InputBox("Enter the item number to remove:", "Remove Item")
    If itemNum = "" Then Exit Sub
    
    ' Validate input: accept 7 digits OR 7 digits + "." + 2 digits
    If Not itemNum Like "#######" And Not itemNum Like "#######.##" Then
        MsgBox "Invalid item number. Please enter a 7-digit number, optionally with a 2-digit suffix (e.g., 0586790 or 0586790.10).", vbExclamation
        Exit Sub
    End If
    
    ' Unprotect the sheet
    ws.Unprotect
    
    ' Find the item in column B
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    found = False
    For itemRow = 1 To lastRow
        If ws.Cells(itemRow, "B").Text = itemNum Then
            found = True
            Exit For
        End If
    Next itemRow
    
    If Not found Then
        MsgBox "Item " & itemNum & " not found in ItemList.", vbExclamation
        ws.Protect , UserInterfaceOnly:=True
        Exit Sub
    End If
    
    ' Get Item Name (used in Logger)
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
    
    ' Get suffix from column C (if any) and build full sheet name
    suffix = ws.Cells(itemRow, "C").Text
    fullSheetName = itemNum & suffix
    
    ' Delete the item row
    ws.Rows(itemRow).Delete Shift:=xlUp
    
    ' Check if breakout sheet exists
    On Error Resume Next
    Set breakoutSheet = ThisWorkbook.Sheets(fullSheetName)
    On Error GoTo 0
    
    If Not breakoutSheet Is Nothing Then
        deleteBreakout = MsgBox("A breakout tab for item " & fullSheetName & " exists." & vbCrLf & _
                                "Do you want to delete the breakout tab as well?", vbYesNo + vbQuestion, "Delete Breakout Tab?")
        If deleteBreakout = vbYes Then
            Application.DisplayAlerts = False
            breakoutSheet.Delete
            Application.DisplayAlerts = True
        End If
    End If
    
    MsgBox "Item " & itemNum & " has been removed from the ItemList.", vbInformation
    
    ' Re-protect the sheet
    ws.Protect , UserInterfaceOnly:=True
    
    ' --- Update Last Updated in _MetaData
    Call UpdateEstimateMetaData
    
    ' Log the change in _MetaData
    Call LogEstimateChange("Macro: RemoveItem", "Item: #" & itemNum & " " & itemName & " Removed")

    
    
End Sub



