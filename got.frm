VERSION 5.00
Begin VB.Form got 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Goto Line"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4275
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Goto"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Line number (Must be an integer)"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "got"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If IsNumeric(Text1.Text) Then

'get line number of error
Dim mem1 As Integer
mem1 = Text1.Text

'split multiline text into array
texts = Split(frmMain.ActiveForm.rtftext.Text, vbNewLine)

If Text1.Text - 1 < UBound(texts) Then
'get length of string array before line number
length = 0
For temp = 0 To mem1 - 2
length = length + Len(texts(temp)) + 1
Next

frmMain.ActiveForm.rtftext.find texts(mem1 - 1), length - 2
Unload Me
Else
MsgBox "Invalid line number", vbCritical
End If

Else
MsgBox "Invalid line number", vbCritical
End If


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Command1_Click
End If
End Sub

