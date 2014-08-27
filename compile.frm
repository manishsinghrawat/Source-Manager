VERSION 5.00
Begin VB.Form compiles 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compile Options"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Arguments With Compiler"
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   4695
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Compile options"
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   4695
      Begin VB.OptionButton Option1 
         Caption         =   "Compile All Files"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Compile Current File"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Compile Selected File"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Output Folder"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   285
         Left            =   3360
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton cmdcompile 
      Caption         =   "Compile"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   1815
   End
End
Attribute VB_Name = "compiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdcompile_Click()
On Error GoTo endf
For tem = 0 To 2
If Option1(tem).Value = True Then
temp = tem
End If
Next

Dim fso As New FileSystemObject
If fso.FolderExists(Text1.Text) = True Then

If temp = 0 Then
status = 2
frmMain.ActiveForm.selcomp Text1.Text
End If

If temp = 1 Then
status = 1
frmMain.ActiveForm.compiles Text1.Text
End If

If temp = 2 Then
status = 3
frmMain.ActiveForm.compall Text1.Text
End If

Else
msg = MsgBox("Folder does not exist. Do you want to create it?", vbYesNo, "Error")
If msg = vbYes Then
fso.CreateFolder Text1.Text
cmdcompile_Click
End If
End If
Unload Me
endf:
End Sub

Private Sub Command1_Click()
Load folbro1
folbro1.Show vbModal
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim fso As New FileSystemObject

Option1(1).Enabled = False
For tex = 0 To 100
If UCase(fso.GetExtensionName(frmMain.ActiveForm.pro.List(frmMain.ActiveForm.pro.ListIndex))) = UCase(asso(tex, 0)) Then
Option1(1).Enabled = True
Option1(1).Value = True
End If
Next

If frmMain.ActiveForm.pro.SelCount = 0 Then
Option1(0).Enabled = False
End If
Text1.Text = cfolder
End Sub

