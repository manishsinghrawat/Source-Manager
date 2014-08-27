VERSION 5.00
Begin VB.Form frmreplace 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Replace"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   480
      Width           =   2415
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Whole File"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Selected"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdrall 
      Caption         =   "Replace All"
      Height          =   285
      Left            =   3720
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdreplace 
      Caption         =   "Replace"
      Height          =   285
      Left            =   3720
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Find Next"
      Default         =   -1  'True
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   3720
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Replace With"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Find What"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   165
      Width           =   1335
   End
End
Attribute VB_Name = "frmreplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdfind_Click()
frmMain.ActiveForm.rtftext.find Text1.Text, frmMain.ActiveForm.rtftext.SelStart + frmMain.ActiveForm.rtftext.SelLength
ftext = Text1.Text
reptext = Text2.Text
End Sub

Private Sub cmdrall_Click()
mem = frmMain.ActiveForm.rtftext.SelStart
If Option1.Value = True Then
st = frmMain.ActiveForm.rtftext.SelStart
en = frmMain.ActiveForm.rtftext.SelStart + frmMain.ActiveForm.rtftext.SelLength
End If

frmMain.ActiveForm.rtftext.SelStart = 0
For temp = 0 To 99
frmMain.ActiveForm.rtftext.find Text1.Text, frmMain.ActiveForm.rtftext.SelStart

If frmMain.ActiveForm.rtftext.SelLength > 0 Then
If Not frmMain.ActiveForm.rtftext.SelStart > en - 1 Then
frmMain.ActiveForm.rtftext.SelText = Text2.Text
End If

If Option2.Value = True Then
frmMain.ActiveForm.rtftext.SelText = Text2.Text
End If
End If
Next

frmMain.ActiveForm.rtftext.SelStart = mem
ftext = Text1.Text
reptext = Text2.Text
End Sub

Private Sub cmdreplace_Click()
If frmMain.ActiveForm.rtftext.SelLength = 0 Then
frmMain.ActiveForm.rtftext.find Text1.Text, frmMain.ActiveForm.rtftext.SelStart + frmMain.ActiveForm.rtftext.SelLength
Else
frmMain.ActiveForm.rtftext.SelText = Text2.Text
frmMain.ActiveForm.rtftext.find Text1.Text, frmMain.ActiveForm.rtftext.SelStart + frmMain.ActiveForm.rtftext.SelLength
End If
ftext = Text1.Text
reptext = Text2.Text
End Sub

Private Sub Form_Load()
If frmMain.ActiveForm.rtftext.SelLength > 0 Then
Option1.Value = True
Else
Option2.Value = True
End If
Text1.Text = ftext
Text2.Text = reptext
End Sub

