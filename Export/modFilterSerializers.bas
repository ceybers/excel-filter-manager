Attribute VB_Name = "modFilterSerializers"
Option Explicit

Private Type encFilter
    Index As Long
    On As Boolean
    count As Long
    Operator As XlAutoFilterOperator
    Criteria1 As String
    Criteria2 As String
End Type

Private Function ToString(f As encFilter) As String
    ToString = CStr(f.Index) & "," & f.On & "," & f.count & "," & f.Operator & "," & f.Criteria1 & "," & f.Criteria2
End Function

Private Function FromString(s As String) As encFilter
    Dim arr As Variant
    arr = Split(s, ",")
    Debug.Assert UBound(arr) = 5
    With FromString
        .Index = CLng(arr(0))
        .On = CBool(arr(1))
        .count = CLng(arr(2))
        .Operator = CLng(arr(3))
        .Criteria1 = CStr(arr(4))
        .Criteria2 = CStr(arr(5))
    End With
End Function

Private Function SaveColumn(f As Filter, idx As Long) As encFilter
    SaveColumn.Index = idx
    SaveColumn.On = f.On
    If SaveColumn.On = False Then Exit Function
    
    SaveColumn.Operator = f.Operator
    SaveColumn.count = f.count
    
    Select Case f.Operator
        Case XlAutoFilterOperator.xlFilterCellColor
            SaveColumn.Criteria1 = SerializeInterior(f.Criteria1)
        Case XlAutoFilterOperator.xlFilterFontColor
            SaveColumn.Criteria1 = SerializeFont(f.Criteria1)
        Case XlAutoFilterOperator.xlFilterIcon
            SaveColumn.Criteria1 = SerializeIcon(f.Criteria1)
        Case XlAutoFilterOperator.xlFilterDynamic
            SaveColumn.Criteria1 = f.Criteria1
        Case XlAutoFilterOperator.xlAnd
            SaveColumn.Criteria1 = SerializeString(f.Criteria1)
            SaveColumn.Criteria2 = SerializeString(f.Criteria2)
        Case XlAutoFilterOperator.xlOr
            SaveColumn.Criteria1 = SerializeString(f.Criteria1)
            SaveColumn.Criteria2 = SerializeString(f.Criteria2)
        Case XlAutoFilterOperator.xlFilterAutomaticFontColor
        Case XlAutoFilterOperator.xlFilterNoFill
        Case XlAutoFilterOperator.xlFilterNoIcon
        Case Else
            Select Case VarType(f.Criteria1)
                Case (vbArray + vbVariant)
                    SaveColumn.Criteria1 = SerializeVariantArray(f.Criteria1)
                Case Else
                    SaveColumn.Criteria1 = SerializeString(f.Criteria1)
            End Select
    End Select
End Function

Private Function LoadColumn(rng As Range, mf As encFilter)
    If mf.On = False Then Exit Function
    Select Case mf.Operator
        Case XlAutoFilterOperator.xlFilterCellColor
            rng.AutoFilter mf.Index, CLng(mf.Criteria1), mf.Operator
        Case XlAutoFilterOperator.xlFilterFontColor
            rng.AutoFilter mf.Index, CLng(mf.Criteria1), mf.Operator
        Case XlAutoFilterOperator.xlFilterIcon
            rng.AutoFilter mf.Index, DeserializeIcon(mf.Criteria1), mf.Operator
        Case XlAutoFilterOperator.xlFilterDynamic
            rng.AutoFilter mf.Index, CInt(mf.Criteria1), mf.Operator
        Case XlAutoFilterOperator.xlAnd
            rng.AutoFilter mf.Index, DeserializeString(mf.Criteria1), mf.Operator, DeserializeString(mf.Criteria2)
        Case XlAutoFilterOperator.xlOr
            rng.AutoFilter mf.Index, DeserializeString(mf.Criteria1), mf.Operator, DeserializeString(mf.Criteria2)
        Case XlAutoFilterOperator.xlFilterAutomaticFontColor
            rng.AutoFilter mf.Index, , mf.Operator
        Case XlAutoFilterOperator.xlFilterNoFill
            rng.AutoFilter mf.Index, , mf.Operator
        Case XlAutoFilterOperator.xlFilterNoIcon
            rng.AutoFilter mf.Index, , mf.Operator
        Case Else
            If mf.count > 2 Then ' Array
                rng.AutoFilter Field:=mf.Index, Operator:=mf.Operator, Criteria1:=DeserializeVariantArray(mf.Criteria1)
            Else
                rng.AutoFilter Field:=mf.Index, Criteria1:=DeserializeString(mf.Criteria1)
            End If
    End Select
End Function

Public Function SerializeTableFilters(lo As ListObject) As String
    Dim idx As Long
    Dim thisFilter As Filter
    Dim filtersAsString() As String
    
    ReDim filtersAsString(1 To lo.AutoFilter.Filters.count)
    
    For idx = 1 To lo.AutoFilter.Filters.count
        Set thisFilter = lo.AutoFilter.Filters(idx)
        filtersAsString(idx) = ToString(SaveColumn(thisFilter, idx))
    Next idx
    
    SerializeTableFilters = Join(filtersAsString, ";")
End Function

Public Function DeserializeTableFilters(lo As ListObject, Payload As String) As Boolean
    Dim idx As Long
    Dim dbr As Range
    Dim filtersDecombined() As String
    Dim filterCount As Long
    
    filtersDecombined = Split(Payload, ";") ' starts at 0!
    filterCount = UBound(filtersDecombined, 1) - LBound(filtersDecombined, 1) + 1
    
    Set dbr = lo.DataBodyRange
    
    For idx = 1 To filterCount
        Call LoadColumn(dbr, FromString(filtersDecombined(idx - 1)))
    Next idx
    
    DeserializeTableFilters = True
End Function

Private Function SerializeString(ByVal v As Variant) As String
    SerializeString = StringToBase64(v)
End Function

Private Function DeserializeString(ByVal v As Variant) As String
    DeserializeString = Base64toString(v)
End Function

Private Function SerializeVariantArray(ByVal v As Variant) As String
    Dim i As Integer
    
    For i = LBound(v, 1) To UBound(v, 1)
        v(i) = StringToBase64(v(i))
    Next i
    SerializeVariantArray = StringToBase64(Join(v, ","))
End Function

Private Function DeserializeVariantArray(Payload As String) As Variant
    Dim decodedPayload As String
    decodedPayload = Base64toString(Payload)
    Dim splitPayloads As Variant
    splitPayloads = Split(decodedPayload, ",")
    Dim count As Integer
    count = UBound(splitPayloads, 1) - LBound(splitPayloads, 1) + 1
    Dim offset As Variant
    ReDim offset(1 To count)
    Dim idx As Long
    For idx = LBound(splitPayloads, 1) To UBound(splitPayloads, 1)
        offset(idx + 1) = Base64toString(splitPayloads(idx))
    Next idx
    DeserializeVariantArray = offset
End Function

Private Function SerializeInterior(ByVal v As Variant) As String
    If VarType(v) = vbLong Then
        SerializeInterior = CLng(v)
    Else
        Debug.Assert TypeName(v) = "Interior"
        SerializeInterior = v.Color
    End If
End Function

Private Function SerializeFont(ByVal v As Variant) As String
    If VarType(v) = vbLong Then
        SerializeFont = CLng(v)
    Else
        Debug.Assert TypeName(v) = "Font"
        SerializeFont = v.Color
    End If
End Function

Private Function SerializeIcon(ByVal v As Variant) As String
    Debug.Assert TypeName(v) = "Icon"
    SerializeIcon = v.Parent.ID & "|" & v.Index
End Function

Private Function DeserializeIcon(ByVal v As Variant) As Variant
    Dim arr As Variant
    arr = Split(v, "|")
    Set DeserializeIcon = ThisWorkbook.IconSets(arr(0)).Item(arr(1))
End Function

' EOF
