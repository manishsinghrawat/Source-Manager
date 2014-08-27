VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form addpat 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Patch"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add Patch"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   3600
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox files 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "File name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "addpat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo endf
Com.DialogTitle = "Patch File"
Com.Filter = "*.CMP (Compiler Files)|*.cmp"
CancelError = True
Com.ShowOpen
files.Text = Com.filename
endf:
End Sub

Private Sub Command2_Click()
Dim fso As New FileSystemObject
Dim ff As File
Dim txt As TextStream
'check if file specified exists
If fso.FileExists(files.Text) = True Then

'create compiler folder if it does not exist
If fso.FolderExists("c:\windows\temp\compiler") = False Then
fso.CreateFolder "c:\windows\temp\compiler"
End If

'set file handler
Set ff = fso.GetFile(files.Text)
Short = fso.GetParentFolderName(ff.ShortPath) 'set file's short name
Set txt = fso.OpenTextFile(files.Text, ForReading, False)

'delete intermediate file iof it exists
If fso.FileExists("c:\windows\temp\data.txt") = True Then fso.DeleteFile ("c:\windows\temp\data.txt")

'execute batch file which creates result file and sets up compiler in compiler folder
Shell fso.GetFolder(workpro).ShortPath + "\patch.bat " + fso.GetFolder(workpro).ShortPath + " " + fso.GetDriveName(workpro) + " " + fso.GetFile(files.Text).ShortPath + " " + fso.GetFolder(workpro).ShortPath + "\compiler\", vbMaximizedFocus

'load result showing form
Load Info
Info.setfile ("c:\windows\temp\data.txt")  'set intermediate data file
Info.termination ("All OK")                'set termination string
Info.Show vbModal

'set textstream for reading
Set txt = fso.OpenTextFile(workpro + "\compiler\compiler.dat", ForReading, False)

'this section checks whether the compiler we are installing is already registered
Temx = txt.ReadLine
For temp = 1 To 100
If Temx = GetSetting(App.Title, "compiler", "CMP" + CStr(temp), vbNullString) Then
MsgBox "Sorry other compiler with same name is already registered. Aborted", vbCritical, "Registration Conflict"
GoTo endf
End If
Next
'//////checking section ends here//////

'Retrives information from compiler information
'file and adds to registry under compiler key

For temp = 1 To 100
If vbNullString = GetSetting(App.Title, "compiler", "CMP" + CStr(temp), vbNullString) Then
SaveSetting App.Title, "compiler", "CMP" + CStr(temp), Temx
SaveSetting App.Title, "compiler", "DIR" + CStr(temp), txt.ReadLine
SaveSetting App.Title, "compiler", "HANDLES" + CStr(temp), txt.ReadLine
SaveSetting App.Title, "compiler", "TYP" + CStr(temp), txt.ReadLine
SaveSetting App.Title, "compiler", "RES" + CStr(temp), txt.ReadLine
GoTo endx
End If
Next
endx:
txt.Close
'split extensions compiler can be associated with
texts = Split(GetSetting(App.Title, "compiler", "HANDLES" + CStr(temp), vbNullString), "/")
cmpno = temp

'This section associates compiler files with extensions
For temp = 0 To UBound(texts)
'checks if extension is already registered
For tex = 1 To 100
If texts(temp) = GetSetting(App.Title, "association", "ext" + CStr(tex), vbNullString) Then
msg = MsgBox("Installing : " + Temx + vbNewLine + "Extension : " + texts(temp) + vbNewLine + "Clashes with " + GetSetting(App.Title, "association", "asso" + CStr(tex), vbNullString) + vbNewLine + "Do you want to really associate compiler with " + texts(temp), vbQuestion + vbYesNo, "Conflict")
If msg = vbYes Then
register CStr(texts(temp)), CStr(Split(GetSetting(App.Title, "compiler", "typ" + CStr(cmpno), vbNullString), "/")(temp)), CStr(GetSetting(App.Title, "compiler", "CMP" + CStr(cmpno), vbNullString))
End If
GoTo nexts
End If

If vbNullString = GetSetting(App.Title, "association", "ext" + CStr(tex), vbNullString) Then
'registers if extension is not registered
register CStr(texts(temp)), CStr(Split(GetSetting(App.Title, "compiler", "typ" + CStr(cmpno), vbNullString), "/")(temp)), CStr(GetSetting(App.Title, "compiler", "CMP" + CStr(cmpno), vbNullString))
GoTo nexts
End If
Next

nexts:
Next

Else
MsgBox "File not Found : Patching Aborted!", , "Critical Error"
GoTo endf
End If

GoTo endf
err:
MsgBox "Error : " + err.Number + " : " + err.Description, vbCritical, "Critical Error"
endf:
End Sub

Function register(ext As String, des As String, asso As String)
'registers extension into registry
For temp = 1 To 100
If ext = GetSetting(App.Title, "association", "ext" + CStr(temp), vbNullString) Then
SaveSetting App.Title, "association", "ext" + CStr(temp), ext
SaveSetting App.Title, "association", "des" + CStr(temp), des
SaveSetting App.Title, "association", "asso" + CStr(temp), asso
GoTo endx
End If

If vbNullString = GetSetting(App.Title, "association", "ext" + CStr(temp), vbNullString) Then
SaveSetting App.Title, "association", "ext" + CStr(temp), ext
SaveSetting App.Title, "association", "des" + CStr(temp), des
SaveSetting App.Title, "association", "asso" + CStr(temp), asso
GoTo endx
End If
Next

endx:
End Function

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub files_Click()
If KeyAscii = vbKeyReturn Then
Command2_Click
End If
End Sub

