VERSION 5.00
Begin VB.Form Editor 
   AutoRedraw      =   -1  'True
   Caption         =   "Editor"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Discard Changes"
      Height          =   400
      Left            =   1800
      TabIndex        =   2
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   400
      Left            =   120
      TabIndex        =   1
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   6135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim closed As Boolean
Private Sub Command1_Click()
Dim fso As New FileSystemObject
Dim txt As TextStream
Set txt = fso.OpenTextFile(Me.Caption, ForWriting, True)
txt.Write (Text1.Text)
closed = True
Unload Me
End Sub

Private Sub Command2_Click()
closed = True
Unload Me
End Sub

Private Sub Form_resize()
On Error Resume Next
Text1.Width = Me.Width - 2 * Text1.Left - 100
Text1.Height = Me.Height - 2 * Text1.top - 100 - 900
Command1.top = Text1.top + Text1.Height + 100
Command2.top = Text1.top + Text1.Height + 100
Command1.Left = Text1.Left
Command2.Left = Text1.Left + Command1.Width + 100

End Sub

Private Sub Form_Unload(cancel As Integer)
If closed = False Then
cancel = True
End If
End Sub
