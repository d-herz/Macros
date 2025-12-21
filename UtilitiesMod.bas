Option Explicit

' Log change in the _MetaData hidden tab

Public Sub LogEstimateChange(actionText As String, Optional detailsText As String = "")
    
    Dim wsMeta As Worksheet
    Dim logTable As ListObject
    Dim newRow As ListRow
    Dim maxRows As Long
    Dim numRows As Long
    
    maxRows = 200   'Set your maximum number of log entries

    Set wsMeta = ThisWorkbook.Worksheets("_MetaData")
    Set logTable = wsMeta.ListObjects("tblChangeLog")
    
    'Insert new row at the top
    Set newRow = logTable.ListRows.Add(1)
    
    'Populate the row
    With newRow.Range
        .Cells(1, 1).Value = Now                  'Timestamp
        .Cells(1, 2).Value = UserName()          'Username
        .Cells(1, 3).Value = actionText          'Action description
        .Cells(1, 4).Value = detailsText         'Optional details
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

Public Sub UpdateEstimateMetaData()
    On Error Resume Next
    ThisWorkbook.Names("LastUpdatedBy").RefersToRange.Value = UserName()
    ThisWorkbook.Names("LastUpdatedOn").RefersToRange.Value = Now
End Sub

