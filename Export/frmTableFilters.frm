VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableFilters 
   Caption         =   "Table Filter Manager"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmTableFilters.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTableFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DoGet()
Public Event DoSet()
Public Event DoClear()
Public Event Completed()

Public Property Get Payload() As Variant
    Payload = CStr(Me.txtPayload.Value)
End Property

Public Property Let Payload(ByVal vNewValue As Variant)
    Me.txtPayload.Value = CStr(vNewValue)
    Me.cmbSet.Enabled = False
    'Me.cmbSet.Default = True
End Property

Private Sub cmbClear_Click()
    RaiseEvent DoClear
    Me.cmbSet.Enabled = True
End Sub

Private Sub cmbGet_Click()
    RaiseEvent DoGet
    Me.cmbSet.Enabled = False
    Me.cmbOK.Default = True
    Me.txtPayload.SetFocus
End Sub

Private Sub cmbOK_Click()
    RaiseEvent Completed
End Sub

Private Sub cmbSet_Click()
    RaiseEvent DoSet
    Me.cmbSet.Enabled = False
    Me.cmbOK.Default = True
    Me.cmbClear.Enabled = True
End Sub

Private Sub txtPayload_Change()
    If Len(txtPayload.Value) > 0 Then
        Me.cmbSet.Enabled = True
        Me.cmbSet.Default = True
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.cmbGet.Enabled = False
    Me.cmbSet.Enabled = False
    Me.cmbOK.Default = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    RaiseEvent Completed
End Sub

'Public Property Get CanClear() As Variant
    
'End Property

Public Property Let CanClear(ByVal vNewValue As Variant)
    Me.cmbClear.Enabled = CBool(vNewValue)
End Property
