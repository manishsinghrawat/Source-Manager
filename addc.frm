VERSION 5.00
Begin VB.Form addc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Compiler"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Discard"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Handles"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Folder Name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Compiler Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "addc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
For temp = 1 To 100
If "x/x" = GetSetting(App.Title, "compiler", "CMP" + CStr(temp), "x/x") Then
Exit For
End If
Next
If Not temp = 100 Then
SaveSetting App.Title, "compiler", "CMP" + CStr(temp), Text1.Text
SaveSetting App.Title, "compiler", "DIR" + CStr(temp), Text2.Text
SaveSetting App.Title, "compiler", "HANDLES" + CStr(temp), Text3.Text
SaveSetting App.Title, "compiler", "TYP" + CStr(temp), "C++ File"
SaveSetting App.Title, "compiler", "RES" + CStr(temp), "EXE"

End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
