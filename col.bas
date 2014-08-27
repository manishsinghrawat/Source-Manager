Attribute VB_Name = "Special"
Dim intnum As Integer
Dim sellen As Integer

Function storetemp()
On Error Resume Next
frmMain.rtftext.Text = vbNullString
frmMain.rtftext.Font.name = frmMain.ActiveForm.rtftext.Font.name
frmMain.rtftext.Font.Size = frmMain.ActiveForm.rtftext.Font.Size
frmMain.rtftext.Text = frmMain.ActiveForm.rtftext.Text
sellen = frmMain.ActiveForm.rtftext.SelLength
intnum = frmMain.ActiveForm.rtftext.SelStart
End Function

Function appcol(num As Integer)

For temp = 1 To UBound(keywords)
keyw = Split(keywords(temp), "//\\")(0)
word = Split(keywords(temp), "//\\")(1)

laststart = -200
frmMain.rtftext.SelStart = 0
frmMain.rtftext.SelLength = 0

Do Until frmMain.rtftext.SelStart = laststart
'sets variable which checks that if repeated finding has been done
laststart = frmMain.rtftext.SelStart

'this function is used to find a string in a substring
frmMain.rtftext.find keyw, frmMain.rtftext.SelStart + frmMain.rtftext.SelLength

'checks if selection is changed or not or checks if new substring is found
If frmMain.rtftext.SelLength = 0 Then GoTo nexts

'set colour
frmMain.rtftext.SelColor = word


Loop
nexts:
Next
End Function

Function appind()
On Error Resume Next
frmMain.rtftext.SelStart = 1
frmMain.rtftext.SelLength = 0

If frmMain.ActiveForm.rtftext.SelStart = 2 Then
intnum = frmMain.ActiveForm.rtftext.SelStart
Exit Function
End If

intm = Split(frmMain.rtftext.Text, vbNewLine)


Dim texts() As String
If UBound(intm) > -1 Then
ReDim texts(UBound(intm))
Else
Exit Function
End If

For temp = 0 To UBound(intm)
texts(temp) = intm(temp)
Next

curpos = frmMain.ActiveForm.rtftext.SelStart
pos = 0

For temp = 0 To UBound(texts)
Do While InStr(1, texts(temp), vbTab) > 0
If pos < curpos Then
curpos = curpos - 1

End If

temp1 = Left(intm(temp), InStr(1, intm(temp), vbTab) - 1)
temp2 = Right(intm(temp), Len(intm(temp)) - InStr(1, intm(temp), vbTab))
intm(temp) = temp1 + temp2
texts(temp) = temp1 + temp2
Loop
pos = pos + Len(intm(temp)) + 2
Next

Level = 0

Dim result() As String
ReDim result(UBound(texts))
pos = 0
first = False

'set level of tabs
For temp = 0 To UBound(texts)
cur = Level
For mem = 1 To Len(texts(temp))
If Mid(texts(temp), mem, 1) = "{" Then
Level = Level + 1
End If
If Mid(texts(temp), mem, 1) = "}" Then
Level = Level - 1
cur = Level
End If
Next

'create required tabs
vertab = vbNullString
For bui = 1 To cur
vertab = vbTab + vertab
Next
If pos <= curpos Then
curpos = curpos + cur
pos = pos + cur
End If

result(temp) = vertab + texts(temp)
pos = pos + Len(texts(temp)) + 2
Next

resstr = vbNullString
For temp = 0 To UBound(result)
If first = False Then
first = True
resstr = result(temp)
Else
resstr = resstr + vbNewLine + result(temp)
End If

Next

frmMain.rtftext.Text = resstr
intnum = curpos
endf:
End Function

Function addspc()
num = intnum
strs = frmMain.rtftext.Text

mem = 1
'this section is used for special < & > signs after #include preprocessing directive
Do While Not InStr(mem, UCase(strs), UCase("#include")) = 0
numbr = numbr + 1 'increments values of > and <  w/o spaces
numbr1 = numbr1 + 1

mem = InStr(mem, UCase(strs), UCase("#include")) + 8
Loop

If Len(strs) = 0 Then Exit Function
For temp = 1 To UBound(operators)
Do While Not InStr(1, UCase(strs), UCase(operators(temp))) = 0
tex = InStr(1, UCase(strs), UCase(operators(temp)))

'this section checks whether space is already added or not
If Not tex = 1 Then 'to stop from running if tex is zero as tex-1 evaluates to 0
If Mid(strs, tex - 1, 1) = Space(1) Then
prechar = vbNullString
Else
prechar = Space(1)
'increments value of variable which holds current position
If tex <= num Then num = num + 1
End If
End If

If Mid(strs, tex + Len(operators(temp)), 1) = Space(1) Then
postchar = vbNullString
Else
postchar = Space(1)
'increments value of variable which holds current position
If tex + Len(operators(temp)) <= num - Len(prechar) Then
num = num + 1
End If
End If
'checking section ends here


'this section is used to check whether < & > signs are disabled or not
If numbr > 0 Then
If operators(temp) = ">" Then
num = num - Len(prechar) - Len(postchar)
prechar = vbNullString
postchar = vbNullString
numbr = numbr - 1
End If
End If


If numbr1 > 0 Then
If operators(temp) = "<" Then
num = num - Len(prechar) - Len(postchar)
prechar = vbNullString
postchar = vbNullString
numbr1 = numbr1 - 1
End If
End If

'this section is used to convert user understandable form into coded labguage so as to prevent reconversion
If tex + Len(operators(temp)) <= num - Len(prechar) - Len(postchar) Then
num = num + Len("~~" + CStr(temp) + "~~") - Len(operators(temp))
End If


'this is string which
'1 adds spaces prefix and postfix
       'extracts left part                                          extracts right part
strs = Left(strs, tex - 1) + prechar + "~~" + CStr(temp) + "~~" + postchar + Right(strs, Len(strs) - (tex - 1) - Len(operators(temp)))
Loop
Next

For temp = 1 To UBound(operators)
Do While Not InStr(1, strs, "~~" + CStr(temp) + "~~") = 0
tex = InStr(1, strs, "~~" + CStr(temp) + "~~")

'this section converts coded language into user understandable one
If tex + Len("~~" + CStr(temp) + "~~") <= num Then
num = num - Len("~~" + CStr(temp) + "~~") + Len(operators(temp))
End If

'this is string which
'1 adds spaces prefix and postfix
       'extracts left part                                          extracts right part
strs = Left(strs, tex - 1) + operators(temp) + Right(strs, Len(strs) - (tex - 1) - Len("~~" + CStr(temp) + "~~"))
Loop
Next

'decoding section ends here
frmMain.rtftext.Text = strs
intnum = num

End Function

Function appall()
If autocol = False And autospc = False And autoind = False Then
Else
If disabled = False Then
storetemp
If autoind = True Then appind
If autospc = True Then addspc
If autocol = True Then appcol 10000
restoretext
End If
End If
End Function

Function restoretext()
On Error Resume Next
frmMain.ActiveForm.rtftext.TextRTF = frmMain.rtftext.TextRTF
frmMain.ActiveForm.rtftext.SelStart = intnum
frmMain.ActiveForm.rtftext.SelLength = sellen
End Function
