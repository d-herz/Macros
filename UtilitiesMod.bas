Option Explicit
Private UIStack As Long


' Log change in the _MetaData hidden tab

Public Sub LogEstimateChange(actionText As String, Optional detailsText As String = "", Optional targetWB As Workbook = Nothing)
    
    Dim wb As Workbook
    Dim wsMeta As Worksheet
    Dim logTable As ListObject
    Dim newRow As ListRow
    Dim maxRows As Long
    Dim numRows As Long
    
    
     If targetWB Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetWB
    End If
    
    
    maxRows = 200   'Set your maximum number of log entries

    Set wsMeta = wb.Worksheets("_MetaData")
    Set logTable = wsMeta.ListObjects("tblChangeLog")
    
    'Insert new row at the top
    Set newRow = logTable.ListRows.Add(1)
    
    'Populate the row
    With newRow.Range
        .Cells(1, 1).value = Now                  'Timestamp
        .Cells(1, 2).value = UserName()          'Username
        .Cells(1, 3).value = actionText          'Action description
        .Cells(1, 4).value = detailsText         'Optional details
    End With
    
    'Check if we exceeded maxRows
    numRows = logTable.ListRows.count
    If numRows > maxRows Then
        'Delete rows at the bottom to maintain maxRows
        Dim i As Long
        For i = numRows To maxRows + 1 Step -1
            logTable.ListRows(i).Delete
        Next i
    End If

End Sub



' Update the LastUpdatedBy and LastUpdatedOn Meta Data

Public Sub UpdateEstimateMetaData(Optional targetWB As Workbook = Nothing)

    Dim wb As Workbook
    
    If targetWB Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetWB
    End If
    
    On Error Resume Next
    wb.Names("LastUpdatedBy").RefersToRange.value = UserName()
    wb.Names("LastUpdatedOn").RefersToRange.value = Now

End Sub

' Helper subs for freezing and unfreezing UI (for optimizing runtime and reducing screen flickering)

Public Sub FreezeUI()
    If UIStack = 0 Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    End If
    UIStack = UIStack + 1
End Sub

Public Sub UnfreezeUI()
    UIStack = UIStack - 1

    If UIStack <= 0 Then
        UIStack = 0
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub


