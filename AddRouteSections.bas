'This macro is for adding additinal route sections to an item breakout tab

Sub AddRouteSections()
    ' Macro to duplicate a designated template section (a "route")
    
    ' --- CONFIGURATION SECTION ---
    ' *** IMPORTANT: ADJUST THESE ROWS TO MATCH TEMPLATE SIZE ***
    Const TemplateStartRow As Long = 15 ' The first row of template section
    Const TemplateEndRow As Long = 28 ' The last row of template section
    
    ' The offset of the header row from the START of the template (16 - 16 = 0)
    Const HeaderRowOffset As Long = 0
    ' The column containing the item/route header formula
    Const HeaderColumn As String = "B"
    
    ' The starting row for the Q-cell reference (Q4 is for the first/original route)
    Const RouteNameStartRow As Long = 4
    
    ' The row offset within the template where the SECTION TOTAL is located (L26 is 11 rows from the start row 16 -> 26 - 15 = 11)
    Const SectionTotalRowOffset As Long = 11
    
    ' *** IMPORTANT: ADJUST THIS ROW IF YOUR SUBTOTAL SUMMARY STARTS ELSEWHERE ***
    ' This is the row where the "Route #1 Subtotal" currently resides (Row 31)
    Const SubtotalStartRow As Long = 31
    
    ' *** Row offset for the Project Wide Subtotal relative to SubtotalStartRow ***
    ' Assumes ProjectWide Subtotal starts 1 row after Route #1 Subtotal, so Row 32
    Const ProjectWideRowOffset As Long = 1
    ' -----------------------------
    
    Dim ws As Worksheet
    Dim TemplateRange As Range
    Dim SectionsToAdd As Variant
    Dim i As Long
    Dim NumRows As Long
    Dim TargetRow As Long
    Dim NewHeaderRow As Long
    Dim RouteNameRow As Long
    Dim DynamicHeaderFormula As String
    Dim NewSectionTotalRow As Long
    Dim SubtotalInsertionRow As Long
    Dim DynamicSubtotalTextFormula As String
    Dim DynamicSubtotalValueFormula As String
    
    Dim FinalSubtotalRow As Long
    Dim ProjectWideRow As Long
    Dim DynamicProjectWideFormula As String
    
    ' Set the worksheet to the one currently active
    Set ws = ActiveSheet

    ' Calculate the total number of rows in the template (30 - 16 + 1 = 15 rows total)
    NumRows = TemplateEndRow - TemplateStartRow + 1
    
    ' Define the template range to be copied
    On Error Resume Next
    Set TemplateRange = ws.Rows(TemplateStartRow & ":" & TemplateEndRow)
    On Error GoTo 0
    
    If TemplateRange Is Nothing Then
        MsgBox "Error: The template range specified (" & TemplateStartRow & " to " & TemplateEndRow & ") could not be selected.", vbCritical
        Exit Sub
    End If

    ' Prompt the user for the number of sections to add
    SectionsToAdd = Application.InputBox( _
        Prompt:="How many more Route Sections do you need?(e.g., enter '2' to add two duplicates):", _
        Title:="Add Route Sections", _
        Type:=1) ' Type 1 ensures the input is a number
    
    ' Check if the user cancelled (returns Boolean) or entered invalid data
    If VarType(SectionsToAdd) = vbBoolean Then
         MsgBox "Operation cancelled.", vbInformation
        Exit Sub
    End If
    
    If SectionsToAdd <= 0 Then
         MsgBox "Please enter a positive number of sections to add.", vbExclamation
        Exit Sub
    End If
    
    ' Ensure the number is treated as an integer
    SectionsToAdd = Int(SectionsToAdd)

    ' Start inserting the new route sections immediately after the original template block (i.e., row 31)
    TargetRow = TemplateEndRow + 1
    
    Application.ScreenUpdating = False ' Turn off screen updating to speed up the process

    ' Loop to copy and insert the template section repeatedly
    For i = 1 To SectionsToAdd
        ' 1. Copy the original template range (including formatting and formulas)
        TemplateRange.Copy
        
        ' 2. Insert the copied range at the TargetRow, shifting existing content down
        ws.Rows(TargetRow).Insert Shift:=xlDown
        
        ' 3. Paste the copied content into the newly inserted rows
        ws.Rows(TargetRow).PasteSpecial Paste:=xlPasteAll
        
        ' --- 3A: DYNAMIC HEADER UPDATE (Route Section Header) ---
        NewHeaderRow = TargetRow + HeaderRowOffset
        RouteNameRow = RouteNameStartRow + i
        DynamicHeaderFormula = "=CONCAT($C$9, "" for "", Q" & RouteNameRow & ")"
        ws.Range(HeaderColumn & NewHeaderRow).Formula = DynamicHeaderFormula
        
        ' --- 3B: DYNAMIC SUBTOTAL ROW INSERTION ---
        
        ' Calculate the row where the new Subtotal summary row will be inserted.
        ' SubtotalInsertionRow = Original Start + (Shift from Main Sections) + (Shift from Previous Subtotal Insertions)
        SubtotalInsertionRow = (SubtotalStartRow + 1) + (NumRows * i) + (i - 1)

        ' 1. Insert a new row at the calculated SubtotalInsertionRow
        ws.Rows(SubtotalInsertionRow).Insert Shift:=xlDown

        ' 2. Calculate the row where the new section's total value resides (L column)
        NewSectionTotalRow = TargetRow + SectionTotalRowOffset

        ' 3. Define the formulas for the new Subtotal row
        DynamicSubtotalTextFormula = "=CONCAT(Q" & RouteNameRow & ", "" Subtotal"")"
        DynamicSubtotalValueFormula = "=L" & NewSectionTotalRow
        
        ' 4. Apply formulas to the new row (K and L columns)
        ws.Range("K" & SubtotalInsertionRow).Formula = DynamicSubtotalTextFormula
        ws.Range("L" & SubtotalInsertionRow).Formula = DynamicSubtotalValueFormula
        
        ' 5. Copy formatting from the row above (the last route subtotal row)
        ws.Rows(SubtotalInsertionRow - 1).Copy
        ws.Rows(SubtotalInsertionRow).PasteSpecial Paste:=xlPasteFormats
        
        ' 6. Calculate the start row for the next main section insertion
        TargetRow = TargetRow + NumRows
    Next i
    
    ' --- 4. POST-LOOP: UPDATE PROJECT WIDE SUBTOTAL FORMULA ---
    
    If SectionsToAdd > 0 Then
        ' Calculate the row of the Project Wide Subtotal.
        ' It starts at SubtotalStartRow (35) + ProjectWideRowOffset (1) = 36
        ' And then shifted down by (NumRows * SectionsToAdd) + (SectionsToAdd)
        ProjectWideRow = (SubtotalStartRow + ProjectWideRowOffset) + (NumRows * SectionsToAdd) + SectionsToAdd

        ' Calculate the row of the *last* dynamic subtotal inserted.
        ' It is one row above the ProjectWideRow.
        FinalSubtotalRow = ProjectWideRow - 1
        
        ' Calculate the row of the *first* dynamic subtotal (Original Route 1 Subtotal).
        ' This row is only shifted down by the insertion of the new main route sections (NumRows * N).
        ' It is NOT shifted by the new subtotal rows since they are inserted below it.
        FirstSubtotalRow = SubtotalStartRow + (NumRows * SectionsToAdd)
        
        ' Construct the dynamic SUM formula: e.g., =SUM(L35:L52)
        ' L35 is the starting point for the first route subtotal.
        ' We assume all subtotal rows start at L35 and continue sequentially up to FinalSubtotalRow.
        DynamicProjectWideFormula = "=SUM(L" & FirstSubtotalRow & ":L" & FinalSubtotalRow & ")"
        
        ' Apply the new formula to the Project Wide Subtotal cell (Column L)
        ws.Range("L" & ProjectWideRow).Formula = DynamicProjectWideFormula
    End If
    
    Application.CutCopyMode = False ' Clear the clipboard
    Application.ScreenUpdating = True ' Restore screen updating
    
    MsgBox SectionsToAdd & " new route sections and corresponding subtotal rows have been successfully added to the '" & ws.Name & "' sheet! The Project Wide Subtotal has been updated.", vbInformation

End Sub

