'================================================================================
' Module: RemoveItem
' Author: DFH
' Created: January 2026
'
' Purpose: Provides functionality to remove an item from the "ItemList" sheet.
'          Optionally deletes the associated item breakout sheet if it exists.
'
' Key Functionality:
'   - Prompts the user to enter an item number (7 digits, optionally with 2-digit suffix).
'   - Validates input and searches for the item in column B of "ItemList".
'   - Deletes the item row from the ItemList sheet.
'   - Checks for the existence of an associated breakout sheet and prompts the user to optionally deletes it.
'   - Logs the removal in the estimate metadata and marks the Detailed Estimate Sheets (DES) as out-of-date.
'   - Ensures proper sheet protection before and after the operation.
'
' Notes / Assumptions:
'   - ItemList sheet must exist and be unprotected before deletion.
'   - Breakout sheet names are expected to follow the convention: ItemNumber + Suffix (from column C).
'   - Uses _MasterItemBidList to retrieve item descriptions for logging.
'   - DESOutOfDate is a global flag indicating that Detailed Estimate Sheets need regeneration.
'================================================================================
Option Explicit

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
    Dim deleteBreakout As Long
    
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

    ' --- Mark DES as out of date ---
    DESOutOfDate = True

    
End Sub



