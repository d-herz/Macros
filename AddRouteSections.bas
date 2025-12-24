Option Explicit

' This macro is for the "Add Route Section" button on the item breakouts
Public Sub AddRouteSections_UI()
    ' Manual entry point for button / Macros dialog
    ' Uses ActiveSheet and prompts user for count
    Call AddRouteSections(0, ActiveSheet)
End Sub


Sub AddRouteSections(Optional ByVal AutoCount As Long = 0, _
                     Optional ByRef TargetWs As Worksheet)

    '==============================
    ' CONFIGURATION
    '==============================
    Const TemplateStartRow As Long = 15
    Const TemplateEndRow As Long = 28
    Const HeaderRowOffset As Long = 0
    Const HeaderColumn As String = "B"
    Const RouteNameStartRow As Long = 4

    Const TotalAreaRowOffset As Long = 9
    Const TotalVolumeRowOffset As Long = 10
    Const TotalItemRowOffset As Long = 11
    Const SectionTotalRowOffset As Long = 11

    Const SubtotalStartRow As Long = 31
    Const ProjectWideRowOffset As Long = 1

    '==============================
    ' VARIABLES
    '==============================
    Dim ws As Worksheet
    Dim TemplateRange As Range
    Dim SectionsToAdd As Long
    Dim i As Long
    Dim numRows As Long
    Dim TargetRow As Long
    Dim NewHeaderRow As Long
    Dim RouteNameRow As Long
    Dim NewSectionTotalRow As Long
    Dim SubtotalInsertionRow As Long
    Dim FinalSubtotalRow As Long
    Dim ProjectWideRow As Long
    Dim FirstSubtotalRow As Long

    '==============================
    ' Target worksheet resolution
    '==============================
    If Not TargetWs Is Nothing Then
        Set ws = TargetWs
    Else
        Set ws = ActiveSheet
    End If

    '==============================
    ' Section count resolution
    '==============================
    If AutoCount > 0 Then
        SectionsToAdd = AutoCount
    Else
        SectionsToAdd = Application.InputBox( _
            "How many more Route Sections do you need?", _
            "Add Route Sections", _
            Type:=1)

        If SectionsToAdd <= 0 Then Exit Sub
    End If

    '==============================
    ' Template setup
    '==============================
    numRows = TemplateEndRow - TemplateStartRow + 1
    Set TemplateRange = ws.Rows(TemplateStartRow & ":" & TemplateEndRow)

    TargetRow = TemplateEndRow + 1
    Application.ScreenUpdating = False

    '==============================
    ' Main loop
    '==============================
    For i = 1 To SectionsToAdd

        TemplateRange.Copy
        ws.Rows(TargetRow).Insert Shift:=xlDown
        ws.Rows(TargetRow).PasteSpecial xlPasteAll

        NewHeaderRow = TargetRow + HeaderRowOffset
        RouteNameRow = RouteNameStartRow + i

        ws.Range(HeaderColumn & NewHeaderRow).Formula = _
            "=CONCAT($C$9,"" for "",Q" & RouteNameRow & ")"

        ' Totals
        ws.Range("B" & (TargetRow + TotalAreaRowOffset)).Formula = _
            "=CONCAT(""Total Area of "",$C$9,"" for "",Q" & RouteNameRow & ","" ="")"

        ws.Range("B" & (TargetRow + TotalVolumeRowOffset)).Formula = _
            "=CONCAT(""Total Volume of "",$C$9,"" for "",Q" & RouteNameRow & ","" ="")"

        ws.Range("B" & (TargetRow + TotalItemRowOffset)).Formula = _
            "=CONCAT(""Total "",$C$9,"" for "",Q" & RouteNameRow & ","" ="")"

        SubtotalInsertionRow = (SubtotalStartRow + 1) + (numRows * i) + (i - 1)
        ws.Rows(SubtotalInsertionRow).Insert Shift:=xlDown

        NewSectionTotalRow = TargetRow + SectionTotalRowOffset
        ws.Range("K" & SubtotalInsertionRow).Formula = "=Q" & RouteNameRow & " & "" Subtotal"""
        ws.Range("L" & SubtotalInsertionRow).Formula = "=L" & NewSectionTotalRow

        ws.Rows(SubtotalInsertionRow - 1).Copy
        ws.Rows(SubtotalInsertionRow).PasteSpecial xlPasteFormats

        TargetRow = TargetRow + numRows
    Next i

    '==============================
    ' Project-wide subtotal
    '==============================
    ProjectWideRow = (SubtotalStartRow + ProjectWideRowOffset) + _
                     (numRows * SectionsToAdd) + SectionsToAdd

    FinalSubtotalRow = ProjectWideRow - 1
    FirstSubtotalRow = SubtotalStartRow + (numRows * SectionsToAdd)

    ws.Range("L" & ProjectWideRow).Formula = _
        "=SUM(L" & FirstSubtotalRow & ":L" & FinalSubtotalRow & ")"

    Application.CutCopyMode = False
    Application.ScreenUpdating = True

End Sub


