Attribute VB_Name = "digitalMars"
Sub dmars(ID As String)
'declarations
Dim fso As New FileSystemObject
Dim ff As file
Dim txt As TextStream

'delete temporary execuable if available
If fso.FileExists("c:\windows\temp\tex.exe") = True Then
fso.DeleteFile "c:\windows\temp\tex.exe", True
End If


'create temporary source file
fso.CreateTextFile "c:\windows\temp\tex.cpp", True
Set ff = fso.GetFile("c:\windows\temp\tex.cpp")
Set txt = ff.OpenAsTextStream(ForWriting)

mem = frmMain.ActiveForm.pro.ListIndex
frmMain.ActiveForm.pro.ListIndex = ID
frmMain.ActiveForm.doit
txt.Write (frmMain.ActiveForm.rtfText.Text)
frmMain.ActiveForm.pro.ListIndex = mem

txt.Close

'run maker file which converts source into executable
Shell CStr(workpro) + "\dm" + "\exe\make.bat " + workpro + "\dm", vbHide

Load Wait
Wait.Show vbModal

'debugging unit
If fso.FileExists("c:\windows\temp\tex.exe") = False Then
errr = MsgBox("Error,Do you want debugging to be started", vbYesNo, "Error")

If errr = vbYes Then
texts = Split(workpro, "\")
Shell CStr(workpro) + "\dm" + "\exe\cr.bat " + workpro + "\dm " + texts(0), vbNormalFocus
End If
End If

'run unit
If fso.FileExists("c:\windows\temp\tex.exe") = True Then
Load stp
stp.Show vbModal
Shell CStr(workpro) + "\dm" + "\exe\crun.bat", vbNormalFocus
End If
End Sub

Function bugs()
'declarations
Dim fso As New FileSystemObject
Dim ff As file
Dim txt As TextStream

'delete temporary execuable if available
If fso.FileExists("c:\windows\temp\tex.exe") = True Then
fso.DeleteFile "c:\windows\temp\tex.exe", True
End If


'create temporary source file
fso.CreateTextFile "c:\windows\temp\tex.cpp", True
Set ff = fso.GetFile("c:\windows\temp\tex.cpp")
Set txt = ff.OpenAsTextStream(ForWriting)
txt.Write (frmMain.ActiveForm.rtfText.Text)
txt.Close

Load Wait
Wait.Show vbModal

'run maker file which converts source into executable
Shell CStr(workpro) + "\dm" + "\exe\make.bat " + workpro + "\dm", vbHide

'debugging unit
If fso.FileExists("c:\windows\temp\tex.exe") = False Then
Beep
errr = MsgBox("Error,Do you want debugging to be started", vbYesNo, "Error")
If errr = vbYes Then
texts = Split(workpro, "\")
Shell CStr(workpro) + "\dm" + "\exe\cr.bat " + workpro + "\dm" + " " + texts(0), vbNormalFocus
End If
End If
End Function
