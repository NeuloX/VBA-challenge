
Attribute VB_Name = "Visuals"
Option Explicit
Dim rowrange As Long
Sub format()
Attribute format.VB_ProcData.VB_Invoke_Func = " \n14"
    'set percent change column to percentage format
    Columns(11).NumberFormat = "0.00\%"
    Range(Cells(2, 17), Cells(3, 17)).NumberFormat = "0.00\%"
    Columns.AutoFit
    


End Sub
Sub conditional_format()
Attribute conditional_format.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
    
    '
    rowrange = Cells(Rows.Count, 10).End(xlUp).Row
    Range(Cells(2, 10), Cells(rowrange, 10)).Select
    'Conditional Formating
    'if values are (0>value) negative get to red
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16777024
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 8290026
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    'if values are (0++) positive set green
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
        , Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = -0.499984740745262
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False

End Sub
