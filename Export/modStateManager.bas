Attribute VB_Name = "modStateManager"
Option Explicit

Private Const SHEET_CODE_NAME = "FilterStateMgrPersistentData"

Private Type filterState
    tableName As String
    StateName As String
    Payload As String
End Type

Public Sub TESTState()
    Debug.Print "TESTState()"
    Dim states() As filterState
    Dim idx As Integer
    states = GetStates("Table1")
    
    Debug.Print "Printing " & UBound(states) & " filter state(s)"
    For idx = 1 To UBound(states)
        Debug.Print idx & ") " & states(idx).StateName
    Next idx
    
    Debug.Print ""
End Sub

Private Function GetStates(Optional tableName As String) As filterState()
    Dim ws As Worksheet
    Dim arr As Variant
    Dim states() As filterState
    Dim idx As Integer
    Dim offset As Integer
    Set ws = ThisWorkbook.Worksheets(SHEET_CODE_NAME)
    arr = ws.UsedRange.Value
    ReDim states(1 To UBound(arr, 1))
    For idx = 1 To UBound(arr, 1)
        Dim fs As filterState
        With fs
            .tableName = arr(idx, 1)
            .StateName = arr(idx, 2)
            .Payload = arr(idx, 3)
        End With
        If tableName = vbNullString Then
            states(idx) = fs
        Else
            If fs.tableName = tableName Then
                states(idx + offset) = fs
            Else
                offset = offset - 1
            End If
        End If
    Next idx
    ReDim Preserve states(1 To (UBound(arr, 1) + offset))
    GetStates = states
End Function
