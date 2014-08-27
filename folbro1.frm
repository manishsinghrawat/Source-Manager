VERSION 5.00
Begin VB.Form folbro1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Browse For Folder"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3735
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Discard"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   4080
      Width           =   1455
   End
End
Attribute VB_Name = "folbro1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
compiles.Text1.Text = Dir1.path
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
End Sub
