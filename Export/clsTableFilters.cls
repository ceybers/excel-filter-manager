VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTableFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frm As frmTableFilters
Attribute frm.VB_VarHelpID = -1
Private table As ListObject

Private Sub Class_Initialize()
    Set frm = New frmTableFilters
End Sub

Private Sub Class_Terminate()
    Set frm = Nothing
End Sub

Public Sub Go(lo As ListObject)
    Debug.Assert Not lo Is Nothing
    Set table = lo
    CheckClear
    Call frm_DoGet
    
    frm.Show
End Sub

Private Sub frm_Completed()
    frm.Hide
End Sub

Private Sub frm_DoClear()
    If table.AutoFilter.FilterMode Then
        table.AutoFilter.ShowAllData
    End If
    CheckClear
End Sub

Private Sub frm_DoGet()
    frm.Payload = SerializeTableFilters(table)
End Sub

Private Sub frm_DoSet()
    DeserializeTableFilters table, frm.Payload
End Sub

Private Function CheckClear()
    frm.CanClear = table.AutoFilter.FilterMode
End Function
