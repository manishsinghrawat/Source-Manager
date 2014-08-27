VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fsm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File System Manager"
   ClientHeight    =   4515
   ClientLeft      =   4785
   ClientTop       =   3525
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   5160
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox co 
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   3930
      Hidden          =   -1  'True
      Left            =   120
      Pattern         =   "*.h"
      TabIndex        =   8
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Add Custom External Header"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Open in Header Viewer"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Create New Header"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Edit"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert into project"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.ComboBox cmbcomp 
      Height          =   315
      ItemData        =   "fsm.frx":0000
      Left            =   120
      List            =   "fsm.frx":0002
      TabIndex        =   0
      Text            =   "cmbcomp"
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "fsm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbcomp_click()
'clicking on compiler choose combo box runs this block
On Error GoTo error
setvis
File1.Pattern = "*.h" 'sets pattern of files to show
'sets path to compiler path
File1.path = workpro + "\compiler\" + co.List(cmbcomp.ListIndex) + "\include"

File1.ListIndex = 0
GoTo endf:
error:
'clears file box if error occurs
File1.Pattern = "*.sdgdsgsdg"
endf:
End Sub

Private Sub Command1_Click()
'this block if executed if user chooses to add a header file to active project
If Not File1.ListIndex = -1 Then
msg = MsgBox("Do you want to add " + File1.filename + " to current project?", vbYesNo)
If msg = vbYes Then
If frmMain.ActiveForm Is Nothing = False Then
frmMain.ActiveForm.exthea (File1.path + "\" + File1.filename)
End If
End If
End If
End Sub

Private Sub Command2_Click()
'this block is used to delete a header file form standard include dirtectory if specified compiler
If Not File1.ListIndex = -1 Then
Dim fso As New FileSystemObject
msg = MsgBox("Are you sure you want to delete " + File1.filename + "?", vbYesNo)
If msg = vbYes Then
fso.DeleteFile (File1.path + "\" + File1.filename)
End If
End If
File1.Refresh
End Sub

Private Sub Command3_Click()
'unload fsm
Unload Me
End Sub

Private Sub Command4_Click()
'block for editing of header file choosed in file list box
If Not File1.ListIndex = -1 Then
msg = MsgBox("Are you Sure? Editing may cause header no longer work.", vbYesNo)
If msg = vbYes Then
Dim ed As New Editor
Load ed
ed.Caption = File1.path + "\" + File1.filename
Dim fso As New FileSystemObject
Dim txt As TextStream
Set txt = fso.OpenTextFile(File1.path + "\" + File1.filename, ForReading, False)
ed.Text1.Text = txt.ReadAll
ed.Show vbModal
End If
End If
End Sub

Private Sub Command5_Click()
'creates a new file and opens up for editing
inputs = InputBox("Enter new Header Name(Extension Required)", "New Header", "NewH.h")
fval = False
For temp = 0 To File1.ListCount - 1
If UCase(File1.List(temp)) = UCase(inputs) Then
fval = True
GoTo nex1
End If
Next

nex1:
If fval = False Then
If Not Len(inputs) = 0 Then
Load Editor
Editor.Caption = File1.path + "\" + inputs
Editor.Show vbModal
End If
Else
MsgBox ("Input Error! : File Already Exists")
End If
File1.Refresh
End Sub

Private Sub Command6_Click()
MsgBox ("Not available in current version!")
End Sub

Private Sub Command7_Click()
'add an external header file and add it to standard include directory of choosed compiler
On Error GoTo endf
cdlg.DialogTitle = "Add External Custom Header"
cdlg.Filter = "*.h (C++ Header Files)|*.h"
cdlg.CancelError = True
cdlg.ShowOpen

Dim fso As New FileSystemObject

fval = False
For temp = 0 To File1.ListCount - 1
If UCase(File1.List(temp)) = UCase(fso.GetFileName(cdlg.filename)) Then
fval = True
GoTo nex1
End If
Next
nex1:

If fval = False Then

fso.CopyFile cdlg.filename, File1.path + "\" + fso.GetFileName(cdlg.filename), False
Else
MsgBox ("Input Error! : File Already Exists")
End If
endf:

File1.Refresh
End Sub

Private Sub Form_Load()
'retrives information from compiler registeration area
File1.path = "c:\windows"
'load's Compiler
tex = vbNullString
For temp = 1 To 100
If Not "x/x" = GetSetting(App.Title, "compiler", "CMP" + CStr(temp), "x/x") Then
cmbcomp.AddItem (GetSetting(App.Title, "compiler", "CMP" + CStr(temp), vbNullString))
co.AddItem (GetSetting(App.Title, "compiler", "DIR" + CStr(temp), vbNullString))
End If
Next

On Error Resume Next
cmbcomp.ListIndex = 0

File1.path = workpro + "\compiler\" + co.List(cmbcomp.ListIndex) + "\include"

setvis

File1.ListIndex = 0
End Sub

Function setvis()
Dim bools As Boolean
bools = True
If cmbcomp.ListCount = 0 Then
If Cmpcomp.ListIndex = -1 Then
bools = False
End If
End If
Command1.Enabled = bools
Command2.Enabled = bools
Command3.Enabled = bools
Command4.Enabled = bools
Command5.Enabled = bools
Command6.Enabled = bools
Command7.Enabled = bools
End Function
