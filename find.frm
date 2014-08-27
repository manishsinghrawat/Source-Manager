VERSION 5.00
Begin VB.Form find 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Find Next"
      Default         =   -1  'True
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox names 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Find What"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   160
      Width           =   1335
   End
End
Attribute VB_Name = "find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdfind_Click()
find
End Sub

Function find()
frmMain.ActiveForm.rtftext.find names.Text, frmMain.ActiveForm.rtftext.SelStart + frmMain.ActiveForm.rtftext.SelLength
ftext = names.Text
End Function

Private Sub names_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
find
End If
End Sub

Private Sub Form_Load()
names.Text = ftext
End Sub


