VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmResetPicker 
   Caption         =   "Reset Options"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3975
   OleObjectBlob   =   "frmResetPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmResetPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Reset Options"
    Me.Width = 270
    Me.Height = 310

    Dim L As Single: L = 20
    Dim T As Single: T = 16
    Dim W As Single: W = 220
    Dim H As Single: H = 18
    Dim gap As Single: gap = 26

    ' --- Title label ---
    Dim lblTitle As MSForms.Label
    Set lblTitle = Me.Controls.Add("Forms.Label.1", "lblTitle")
    With lblTitle
        .Caption = "Select items to reset:"
        .Left = L: .Top = T: .Width = W: .Height = H
        .Font.Bold = True
    End With
    T = T + gap + 4

    ' --- Checkboxes ---
    Dim chk As MSForms.CheckBox

    Set chk = Me.Controls.Add("Forms.CheckBox.1", "chkFormat")
    With chk
        .Caption = "Formatting": .Left = L: .Top = T: .Width = W: .Height = H: .Value = True
    End With
    T = T + gap

    Set chk = Me.Controls.Add("Forms.CheckBox.1", "chkList")
    With chk
        .Caption = "Lists": .Left = L: .Top = T: .Width = W: .Height = H: .Value = True
    End With
    T = T + gap

    Set chk = Me.Controls.Add("Forms.CheckBox.1", "chkObjects")
    With chk
        .Caption = "Objects": .Left = L: .Top = T: .Width = W: .Height = H: .Value = True
    End With
    T = T + gap

    Set chk = Me.Controls.Add("Forms.CheckBox.1", "chkTables")
    With chk
        .Caption = "Tables": .Left = L: .Top = T: .Width = W: .Height = H: .Value = True
    End With
    T = T + gap + 6

    Set chk = Me.Controls.Add("Forms.CheckBox.1", "chkStyleAll")
    With chk
        .Caption = "Styles (All)": .Left = L: .Top = T: .Width = W: .Height = H: .Value = True
    End With
    T = T + gap

    Set chk = Me.Controls.Add("Forms.CheckBox.1", "chkStyleBI")
    With chk
        .Caption = "Styles (Built-in only)": .Left = L: .Top = T: .Width = W: .Height = H: .Value = False
    End With
    T = T + gap + 10

    ' --- Select All / Clear All buttons ---
    Dim btn As MSForms.CommandButton

    Set btn = Me.Controls.Add("Forms.CommandButton.1", "btnSelAll")
    With btn
        .Caption = "Select All": .Left = L: .Top = T: .Width = 80: .Height = 24
    End With

    Set btn = Me.Controls.Add("Forms.CommandButton.1", "btnClrAll")
    With btn
        .Caption = "Clear All": .Left = L + 90: .Top = T: .Width = 80: .Height = 24
    End With
    T = T + 36

    ' --- Reset / Cancel buttons ---
    Set btn = Me.Controls.Add("Forms.CommandButton.1", "btnReset")
    With btn
        .Caption = "Reset": .Left = L + 30: .Top = T: .Width = 72: .Height = 26
        .Default = True
    End With

    Set btn = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    With btn
        .Caption = "Cancel": .Left = L + 112: .Top = T: .Width = 72: .Height = 26
        .Cancel = True
    End With
End Sub

' --- Button handlers ---

Private Sub btnReset_Click()
    Me.Tag = "OK"
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Me.Tag = "Cancel"
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Me.Tag = "Cancel"
        Me.Hide
        Cancel = True
    End If
End Sub

' --- Mutual exclusion: Styles (All) vs Styles (Built-in only) ---

Private Sub chkStyleAll_Change()
    If Me.Controls("chkStyleAll").Value = True Then Me.Controls("chkStyleBI").Value = False
End Sub

Private Sub chkStyleBI_Change()
    If Me.Controls("chkStyleBI").Value = True Then Me.Controls("chkStyleAll").Value = False
End Sub

' --- Select All / Clear All ---

Private Sub btnSelAll_Click()
    Me.Controls("chkFormat").Value = True
    Me.Controls("chkList").Value = True
    Me.Controls("chkObjects").Value = True
    Me.Controls("chkTables").Value = True
    Me.Controls("chkStyleAll").Value = True
End Sub

Private Sub btnClrAll_Click()
    Me.Controls("chkFormat").Value = False
    Me.Controls("chkList").Value = False
    Me.Controls("chkObjects").Value = False
    Me.Controls("chkTables").Value = False
    Me.Controls("chkStyleAll").Value = False
    Me.Controls("chkStyleBI").Value = False
End Sub
