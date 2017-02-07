# excel-save-file-with-date



Private Sub CommandButton1_Click()
Dim pName, fName, tName As String

pName = ThisWorkbook.Path
pName = pName & "\portfolio_history\"

fName = "portfolio "
fName = fName & Year(Date) & "_"

If Month(Date) < 10 Then
    fName = fName & "0" & Month(Date) & "_"
    Else
    fName = fName & Month(Date) & "_"
End If

If Day(Date) < 10 Then
    fName = fName & "0" & Day(Date) & ".xlsm"
    Else
    fName = fName & Day(Date) & ".xlsm"
End If

tName = pName & fName
ThisWorkbook.SaveCopyAs tName

MsgBox "File saved as: " & tName, vbOKOnly, "Portfolio"

End Sub

