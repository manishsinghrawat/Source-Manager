VERSION 5.00
Begin VB.Form comed 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compiler Editor"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Uninstall this compiler"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   3600
      TabIndex        =   2
      Top             =   0
      Width           =   3615
      Begin VB.ListBox ass 
         Appearance      =   0  'Flat
         Height          =   1335
         IntegralHeight  =   0   'False
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox fol 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox nam 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Associated with"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Compiler Folder"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Compiler Name"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.ListBox List2 
      Height          =   3960
      Left            =   8520
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "comed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
'load's Compiler
tex = vbNullString
For temp = 1 To 100
If Not "x/x" = GetSetting(App.Title, "compiler", "CMP" + CStr(temp), "x/x") Then
List1.AddItem (GetSetting(App.Title, "compiler", "CMP" + CStr(temp), vbNullString))
List2.AddItem (GetSetting(App.Title, "compiler", "DIR" + CStr(temp), vbNullString))
End If
Next
On Error Resume Next
List1.ListIndex = 0
End Sub


Private Sub List1_Click()
fol.Text = List2.List(List1.ListIndex)
nam.Text = List1.List(List1.ListIndex)

ass.Clear
'main process Call
For tex = 0 To 100
If nam.Text = asso(tex, 1) Then
ass.AddItem asso(tex, 0)
End If
Next

End Sub
