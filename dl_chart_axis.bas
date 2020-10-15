Attribute VB_Name = "dl_chart_axis"
Option Explicit


'
' function to change chart axis autimatically
'

Function setChartAxis(sheetName As String, chartName As String, MinOrMax As String, _
    ValueOrCategory As String, PrimaryOrSecondary As String, Value As Variant)

'Create variables
Dim cht As Chart
Dim valueAsText As String

'Set the chart to be controlled by the function
Set cht = Application.Caller.Parent.Parent.Sheets(sheetName) _
    .ChartObjects(chartName).Chart

'Set Value of Primary axis
If (ValueOrCategory = "Value" Or ValueOrCategory = "Y") _
    And PrimaryOrSecondary = "Primary" Then

    With cht.Axes(xlValue, xlPrimary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'Set Category of Primary axis
If (ValueOrCategory = "Category" Or ValueOrCategory = "X") _
    And PrimaryOrSecondary = "Primary" Then

    With cht.Axes(xlCategory, xlPrimary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'Set value of secondary axis
If (ValueOrCategory = "Value" Or ValueOrCategory = "Y") _
    And PrimaryOrSecondary = "Secondary" Then

    With cht.Axes(xlValue, xlSecondary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'Set category of secondary axis
If (ValueOrCategory = "Category" Or ValueOrCategory = "X") _
    And PrimaryOrSecondary = "Secondary" Then
    With cht.Axes(xlCategory, xlSecondary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'If is text always display "Auto"
If IsNumeric(Value) Then valueAsText = Value Else valueAsText = "Auto"

'Output a text string to indicate the value
setChartAxis = ValueOrCategory & " " & PrimaryOrSecondary & " " _
    & MinOrMax & ": " & valueAsText

End Function





