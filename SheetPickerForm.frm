VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SheetPickerForm 
   Caption         =   "UserForm1"
   ClientHeight    =   1815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "SheetPickerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SheetPickerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()

    ' Center the form
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2

    Dim ws As Worksheet

    ' ---- Add core sheets first (in preferred order) ----
    With Me.ComboBox1
        .Clear
        .AddItem "Dash"
        .AddItem "ProjectInfo"
        .AddItem "SummaryCDM"
        .AddItem "SummaryDOT"
        .AddItem "ItemList"
    End With

    ' ---- Add all other visible sheets ----
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible <> xlSheetVeryHidden Then
            If Not ComboContains(Me.ComboBox1, ws.name) Then
                Me.ComboBox1.AddItem ws.name
            End If
        End If
    Next ws

    Me.ComboBox1.ListIndex = 0

End Sub


Private Sub CommandButton1_Click()
    Dim selectedSheet As String
    
    selectedSheet = Me.ComboBox1.value
    
    If selectedSheet <> "" Then
        On Error Resume Next
        Worksheets(selectedSheet).Activate
        Worksheets(selectedSheet).Range("A1").Select ' optional starting cell
        If Err.Number <> 0 Then MsgBox "Sheet not found!", vbExclamation
        On Error GoTo 0
        Unload Me
    End If
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ' User pressed Esc or clicked the X
        ' Do nothing except close the form
        Cancel = False
    End If
End Sub

