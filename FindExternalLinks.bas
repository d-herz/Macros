' Finds External Links in a Workbook

Sub FindExternalLinks_NoErrors()
    Dim ws As Worksheet
    Dim nm As Name
    Dim shp As Shape
    Dim ch As ChartObject
    Dim pt As PivotTable
    Dim rng As Range
    Dim fc As FormatCondition
    Dim msg As String
    Dim cell As Range
    Dim formulaText As String

    msg = "External Links Found:" & vbCrLf

    ' Check defined names
    For Each nm In ThisWorkbook.Names
        If InStr(1, nm.RefersTo, "[") > 0 Then
            msg = msg & "Named Range: " & nm.Name & " ? " & nm.RefersTo & vbCrLf
        End If
    Next nm

    ' Check each sheet
    For Each ws In ThisWorkbook.Worksheets
        ' Cell formulas
        For Each cell In ws.UsedRange
            If cell.HasFormula Then
                If InStr(1, cell.Formula, "[") > 0 Then
                    msg = msg & "Cell " & ws.Name & "!" & cell.Address & " ? " & cell.Formula & vbCrLf
                End If
            End If
        Next cell

        ' Linked shapes
        For Each shp In ws.Shapes
            If shp.Type = msoLinkedPicture Or shp.Type = msoLinkedOLEObject Then
                msg = msg & "Linked Object in " & ws.Name & ": " & shp.Name & vbCrLf
            End If
        Next shp

        ' Charts (with error handling)
        For Each ch In ws.ChartObjects
            On Error Resume Next
            formulaText = ch.Chart.SeriesCollection(1).Formula
            If Err.Number = 0 Then
                If InStr(1, formulaText, "[") > 0 Then
                    msg = msg & "Chart in " & ws.Name & ": " & ch.Name & " ? " & formulaText & vbCrLf
                End If
            End If
            Err.Clear
            On Error GoTo 0
        Next ch

        ' PivotTables
        For Each pt In ws.PivotTables
            If InStr(1, pt.SourceData, "[") > 0 Then
                msg = msg & "PivotTable in " & ws.Name & ": " & pt.Name & " ? " & pt.SourceData & vbCrLf
            End If
        Next pt

        ' Conditional formatting
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeAllFormatConditions)
        On Error GoTo 0
        If Not rng Is Nothing Then
            For Each cell In rng
                For Each fc In cell.FormatConditions
                    On Error Resume Next
                    If InStr(1, fc.Formula1, "[") > 0 Then
                        msg = msg & "Conditional Format in " & ws.Name & "!" & cell.Address & " ? " & fc.Formula1 & vbCrLf
                    End If
                    On Error GoTo 0
                Next fc
            Next cell
        End If
        Set rng = Nothing
    Next ws

    If msg = "External Links Found:" & vbCrLf Then
        MsgBox "No external links found!", vbInformation
    Else
        MsgBox msg, vbInformation, "External Link Report"
    End If
End Sub


