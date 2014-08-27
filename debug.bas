Attribute VB_Name = "Misc"
Function checkval(data As String) As Boolean
checkval = True
For temp = 0 To frmMain.lstwchr.ListCount - 1
wchr = Right(frmMain.lstwchr.List(temp), 1)
If InStr(1, data, wchr) > 0 Then
checkval = False
GoTo endf
End If
Next

endf:
End Function


Public Function validate(Text As String) As Boolean
validate = True
For temp = 0 To frmMain.ActiveForm.pro.ListCount - 1
If UCase(Split(frmMain.ActiveForm.pro.List(temp), ".")(0)) = UCase(Split(Text, ".")(0)) Then
validate = False
GoTo endf
End If
validate = True
Next
endf:
End Function
