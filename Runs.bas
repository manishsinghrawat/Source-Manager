Attribute VB_Name = "Runs"
Function BuildFile(name As String, data As String)
On Error GoTo err
'name is with extention we will use it as it is
Dim fso As New FileSystemObject
Dim txt As TextStream
frmMain.sbStatusBar.Panels(1).Text = "Building ... " + name
mem = "c:\windows\temp\cpp"
If fso.FolderExists(mem) = False Then fso.CreateFolder mem
Set txt = fso.OpenTextFile(mem + "\" + name, ForWriting, True)
frmMain.rtftext.TextRTF = data
txt.Write (frmMain.rtftext.Text)

GoTo endf
err:
MsgBox err.Description, vbCritical, "Critical Error"
endf:
End Function

Function consEXE(compiler As String, name As String)
On Error GoTo err
Dim fso As New FileSystemObject

'name will be converted into file name
texts = Split(name, ".")
tems = fso.GetDriveName(workpro)
If fso.FileExists("c:\windows\temp\cpp\" + texts(0) + ".exe") = True Then
fso.DeleteFile ("c:\windows\temp\cpp\" + texts(0) + ".exe")
End If

If fso.FileExists("c:\windows\temp\" + texts(0) + ".txt") = True Then fso.DeleteFile "c:\windows\temp\" + texts(0) + ".txt"
If fso.FileExists("c:\windows\temp\cpp\" + texts(0) + ".exe") = True Then fso.DeleteFile "c:\windows\temp\cpp\" + texts(0) + ".exe"

constr = vbNullString
tex = workpro + "\compiler\" + compiler
mem1 = fso.GetDriveName(tex)
Do While Not fso.GetParentFolderName(tex) = tex
mem = fso.GetFolder(tex).ShortName
tex = fso.GetParentFolderName(tex)
constr = mem + "\" + constr
Loop
constr = mem1 + constr

inputs = Split(name, ".")

Shell constr + "exe\cr" + UCase(Split(name, ".")(1)) + "exe.bat " + constr + " " + inputs(0) + " " + tems, vbMaximizedFocus


Load Wait
filess = inputs(0)
Wait.Show vbModal

GoTo endf
err:
MsgBox "Error : Possible Cause : " + err.Description, vbCritical, "Critical Error"
endf:
End Function

Function RunEXE(compiler As String, name As String)
'name will be converted into file name
Dim fso As New FileSystemObject
name = fso.GetFileName(name)
texts = Split(name, ".")
tems = fso.GetDriveName(workpro)

constr = vbNullString
tex = workpro + "\compiler\" + compiler
mem1 = fso.GetDriveName(tex)
Do While Not fso.GetParentFolderName(tex) = tex
mem = fso.GetFolder(tex).ShortName
tex = fso.GetParentFolderName(tex)
constr = mem + "\" + constr
Loop
constr = mem1 + constr

inputs = Split(name, ".")

Shell constr + "exe\run.bat " + inputs(0) + " " + frmMain.ActiveForm.getarg, vbNormalFocus
endf:
End Function

Function outRunEXE(compiler As String, name As String)
'name will be converted into file name
Dim fso As New FileSystemObject
Dim txt As TextStream
name = fso.GetFileName(name)
texts = Split(name, ".")
tems = fso.GetDriveName(workpro)

constr = vbNullString
tex = workpro + "\compiler\" + compiler
mem1 = fso.GetDriveName(tex)
Do While Not fso.GetParentFolderName(tex) = tex
mem = fso.GetFolder(tex).ShortName
tex = fso.GetParentFolderName(tex)
constr = mem + "\" + constr
Loop
constr = mem1 + constr

inputs = Split(name, ".")
Set txt = fso.OpenTextFile("c:\x.bat", ForWriting, True)
txt.WriteLine constr + "exe\run.bat " + inputs(0) + " " + frmMain.ActiveForm.getarg + ">c:\temp.dat"
txt.Close
Shell "c:\x.bat", vbMinimizedNoFocus
endf:
End Function


