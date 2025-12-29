' This is for adding or removing an "A" to the item breakout tab name when a user adds the "A" to column C on ItemList (indicating a special provision)

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim itemNum As String
    Dim cell As Range
    Dim sheetFound As Boolean
    Dim logMessage As String
    Dim logChange As Boolean
    
    ' Only act if column C changed
    If Not Intersect(Target, Me.Columns("C")) Is Nothing Then
        Application.EnableEvents = False
        On Error GoTo ExitHandler
        
        For Each cell In Intersect(Target, Me.Columns("C"))
            itemNum = Me.Cells(cell.Row, "B").Value
            If itemNum = "" Then GoTo NextCell
            
            ' Ensure column C is always uppercase if it's "A"
            If Trim(UCase(cell.Value)) = "A" Then
                cell.Value = "A"
            Else
                cell.Value = ""  ' Clear any other entries
            End If
            
            ' Attempt to find the sheet for the item number
            sheetFound = False
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets(itemNum)      ' Try base name
            If Not ws Is Nothing Then sheetFound = True
            On Error GoTo 0
            
            If Not sheetFound Then
                On Error Resume Next
                Set ws = ThisWorkbook.Sheets(itemNum & "A")  ' Try with A
                If Not ws Is Nothing Then sheetFound = True
                On Error GoTo 0
            End If
            
            ' Update tab name based on column C
            If sheetFound Then
                If cell.Value = "A" Then
                    If Right(ws.name, 1) <> "A" Then ws.name = ws.name & "A"
                    ' --- Mark DES as out of date ---
                    DESOutOfDate = True

                    logChange = True
                    logMessage = "Item #"
    
                Else
                    If Right(ws.name, 1) = "A" Then ws.name = Left(ws.name, Len(ws.name) - 1)
                    ' --- Mark DES as out of date ---
                    DESOutOfDate = True
    
                End If
            Else
                MsgBox "Warning: No Breakout Tab for Item #" & itemNum & " was found.", vbExclamation, "Missing Item Breakout"
            End If
            
NextCell:
            Set ws = Nothing
        Next cell
        
ExitHandler:
        Application.EnableEvents = True
    End If
End Sub

