Attribute VB_Name = "modFilterEnumHelpers"
Option Explicit

Private Function GetDynamicFilterEnumName(c As XlDynamicFilterCriteria) As String
    Select Case c
        Case xlFilterAboveAverage
            GetDynamicFilterEnumName = "Filter all above-average values."
        Case xlFilterAllDatesInPeriodApril
            GetDynamicFilterEnumName = "  Filter all dates in April."
        Case xlFilterAllDatesInPeriodAugust
            GetDynamicFilterEnumName = "Filter all dates in August."
        Case xlFilterAllDatesInPeriodDecember
            GetDynamicFilterEnumName = "Filter all dates in December."
        Case xlFilterAllDatesInPeriodFebruray
            GetDynamicFilterEnumName = "Filter all dates in February."
        Case xlFilterAllDatesInPeriodJanuary
            GetDynamicFilterEnumName = "Filter all dates in January."
        Case xlFilterAllDatesInPeriodJuly
            GetDynamicFilterEnumName = "Filter all dates in July."
        Case xlFilterAllDatesInPeriodJune
            GetDynamicFilterEnumName = "Filter all dates in June."
        Case xlFilterAllDatesInPeriodMarch
            GetDynamicFilterEnumName = "Filter all dates in March."
        Case xlFilterAllDatesInPeriodMay
            GetDynamicFilterEnumName = "Filter all dates in May."
        Case xlFilterAllDatesInPeriodNovember
            GetDynamicFilterEnumName = "Filter all dates in November."
        Case xlFilterAllDatesInPeriodOctober
            GetDynamicFilterEnumName = "Filter all dates in October."
        Case xlFilterAllDatesInPeriodQuarter1
            GetDynamicFilterEnumName = "Filter all dates in Quarter1."
        Case xlFilterAllDatesInPeriodQuarter2
            GetDynamicFilterEnumName = "Filter all dates in Quarter2."
        Case xlFilterAllDatesInPeriodQuarter3
            GetDynamicFilterEnumName = "Filter all dates in Quarter3."
        Case xlFilterAllDatesInPeriodQuarter4
            GetDynamicFilterEnumName = "Filter all dates in Quarter4."
        Case xlFilterAllDatesInPeriodSeptember
            GetDynamicFilterEnumName = "Filter all dates in September."
        Case xlFilterBelowAverage
            GetDynamicFilterEnumName = "Filter all below-average values."
        Case xlFilterLastMonth
            GetDynamicFilterEnumName = "Filter all values related to last month."
        Case xlFilterLastQuarter
            GetDynamicFilterEnumName = "Filter all values related to last quarter."
        Case xlFilterLastWeek
            GetDynamicFilterEnumName = "Filter all values related to last week."
        Case xlFilterLastYear
            GetDynamicFilterEnumName = "Filter all values related to last year."
        Case xlFilterNextMonth
            GetDynamicFilterEnumName = "Filter all values related to next month."
        Case xlFilterNextQuarter
            GetDynamicFilterEnumName = "Filter all values related to next quarter."
        Case xlFilterNextWeek
            GetDynamicFilterEnumName = "Filter all values related to next week."
        Case xlFilterNextYear
            GetDynamicFilterEnumName = "Filter all values related to next year."
        Case xlFilterThisMonth
            GetDynamicFilterEnumName = "Filter all values related to the current month."
        Case xlFilterThisQuarter
            GetDynamicFilterEnumName = "Filter all values related to the current quarter."
        Case xlFilterThisWeek
            GetDynamicFilterEnumName = "Filter all values related to the current week."
        Case xlFilterThisYear
            GetDynamicFilterEnumName = "Filter all values related to the current year."
        Case xlFilterToday
            GetDynamicFilterEnumName = "Filter all values related to the current date."
        Case xlFilterTomorrow
            GetDynamicFilterEnumName = "Filter all values related to tomorrow."
        Case xlFilterYearToDate
            GetDynamicFilterEnumName = "Filter all values from today until a year ago."
        Case xlFilterYesterday
            GetDynamicFilterEnumName = "Filter all values related to yesterday."
    End Select
End Function

Private Function FilterToString(f As Filter) As String
    If f.On = False Then
        FilterToString = "No filter"
        Exit Function
    End If
    
    Select Case f.Operator
        Case xlAnd
            FilterToString = "and(" & CStr(f.Criteria1) & "," & CStr(f.Criteria2) & ")"
        Case xlOr
            FilterToString = "or(" & CStr(f.Criteria1) & "," & CStr(f.Criteria2) & ")"
        Case xlTop10Items
            FilterToString = "highval(" & f.Criteria1 & ")"
        Case xlBottom10Items
            FilterToString = "lowval(" & f.Criteria1 & ")"
        Case xlTop10Percent
            FilterToString = "highpct(" & f.Criteria1 & ")"
        Case xlBottom10Percent
            FilterToString = "lowpct(" & f.Criteria1 & ")"
        Case xlFilterValues
            FilterToString = "Filter values"
        Case xlFilterCellColor
            FilterToString = "cellcolor(" & CStr(VariantToColor(f.Criteria1)) & ")"
        Case xlFilterFontColor
            FilterToString = "fontcolor(" & CStr(VariantToColor(f.Criteria1)) & ")"
        Case xlFilterIcon
            FilterToString = "icon(" & CStr(f.Criteria1.Parent.ID) & "x" & CStr(f.Criteria1.Index) & ")"
        Case xlFilterDynamic
            FilterToString = "dynamic(" & GetDynamicFilterEnumName(f.Criteria1) & ")"
        ' non documented
        Case xlFilterNoFill
            FilterToString = "cellcolor(nofill)"
        Case xlFilterAutomaticFontColor
            FilterToString = "fontcolor(default)"
        Case 0
            FilterToString = "eval(" & CStr(f.Criteria1) & ")"
        Case Else
            FilterToString = "OTHER"
            Debug.Assert False
    End Select
End Function

Private Function VariantToColor(v As Variant) As Long
    Select Case VarType(v)
        Case vbLong
            VariantToColor = v
        Case vbObject
            If TypeName(v) = "Interior" Then
                VariantToColor = v.Color
            ElseIf TypeName(v) = "Font" Then
                VariantToColor = v.Color
            Else
                Debug.Assert False
            End If
        Case Else
            Debug.Assert False
    End Select
End Function

Private Function OperatorToString(op As XlAutoFilterOperator, Optional crit As XlDynamicFilterCriteria) As String
    Select Case op
        Case xlAnd
            OperatorToString = "Logical AND of Criteria1 and Criteria2"
        Case xlOr
            OperatorToString = "Logical OR of Criteria1 or Criteria2"
        Case xlTop10Items
            OperatorToString = "Highest-valued items displayed"
        Case xlBottom10Items
            OperatorToString = "Lowest-valued items displayed"
        Case xlTop10Percent
            OperatorToString = "Highest-valued items displayed"
        Case xlBottom10Percent
            OperatorToString = "Lowest-valued items displayed"
        Case xlFilterValues
            OperatorToString = "Filter values"
        Case xlFilterCellColor
            OperatorToString = "Color of the cell"
        Case xlFilterFontColor
            OperatorToString = "Color of the font"
        Case xlFilterIcon
            OperatorToString = "Filter icon"
        Case xlFilterDynamic
            OperatorToString = "Dynamic filter"
        Case Else
            OperatorToString = "OTHER"
    End Select
End Function

Private Function PrintFilter(f As Filter) As String
    If f.On Then
        PrintFilter = OperatorToString(f.Operator)
    Else
        PrintFilter = "No filter"
    End If
End Function
