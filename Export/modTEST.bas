Attribute VB_Name = "modTEST"
Option Explicit

Public Sub TESTForm()
    Dim lo As ListObject
    Set lo = Selection.ListObject
    If lo Is Nothing Then
        Set lo = ThisWorkbook.Worksheets(1).ListObjects(1)
    End If
    
    With New clsTableFilters
        .Go lo
    End With
End Sub

Public Sub TESTSave()
    Debug.Print "TESTSave()"
    Dim lo As ListObject
    Dim filtersCombined As String
    
    Set lo = ThisWorkbook.Worksheets(1).ListObjects(1)

    filtersCombined = SerializeTableFilters(lo)
    
    Debug.Print "SAVE="
    Debug.Print filtersCombined
    Debug.Print
End Sub

Public Sub TESTLoad()
    Debug.Print "TESTLoad()"
    Dim lo As ListObject
    Dim filtersCombined As String
    Dim idx As Integer
    
    Set lo = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    If lo.AutoFilter.FilterMode Then
        lo.AutoFilter.ShowAllData
    End If
    
    filtersCombined = "1,True,5,7,UFNSQkpERXcsUFNSQkpERXgsUFNSQkpERXksUFNSQkpEZz0sUFNSQkpEaz0=,;2,True,1,11,33,;3,False,0,0,,;4,False,0,0,,"
    filtersCombined = "1,True,1,0,PSRBJDI=,;2,False,0,0,,;3,False,0,0,,;4,False,0,0,,"
    filtersCombined = "1,True,2,2,PSRBJDEw,PSRBJDEx;2,False,0,0,,;3,False,0,0,,;4,False,0,0,,"
    filtersCombined = "1,True,2,1,PD4qYWFhKg==,PD5lZWU=;2,False,0,0,,;3,False,0,0,,;4,False,0,0,,"
    filtersCombined = "1,False,0,0,,;2,True,1,8,13561798,;3,False,0,0,,;4,False,0,0,,"
    filtersCombined = "1,False,0,0,,;2,True,1,12,,;3,False,0,0,,;4,False,0,0,,"
    filtersCombined = "1,False,0,0,,;2,True,1,9,22428,;3,False,0,0,,;4,False,0,0,,"
    filtersCombined = "1,False,0,0,,;2,True,1,13,,;3,False,0,0,,;4,False,0,0,,"
    filtersCombined = "1,False,0,0,,;2,True,1,10,1|2,;3,False,0,0,,;4,False,0,0,,"
    
    DeserializeTableFilters lo, filtersCombined
    
    Debug.Print "OK"
    Debug.Print ""
End Sub
