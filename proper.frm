VERSION 5.00
Begin VB.Form proper 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Project Properties"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   5415
      Begin VB.Label Label7 
         Caption         =   "Last Modified :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label Label6 
         Caption         =   "Created       :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label Label5 
         Caption         =   "File Size     :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.Label Label4 
         Caption         =   "Header Files"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   5175
      End
      Begin VB.Label Label3 
         Caption         =   "Cpp Files"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   5175
      End
      Begin VB.Label Label2 
         Caption         =   "Total Files"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "Project Name"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5055
      End
   End
End
Attribute VB_Name = "proper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label1.Caption = Label1.Caption + " : " + frmMain.ActiveForm.Caption
Label2.Caption = Label2.Caption + "  : " + CStr(frmMain.ActiveForm.pro.ListCount)

Dim fso As New FileSystemObject
mem = 0
For temp = 0 To frmMain.ActiveForm.pro.ListCount - 1
If UCase(fso.GetExtensionName(frmMain.ActiveForm.pro.List(temp))) = "CPP" Then
mem = mem + 1
End If
Next
Label3.Caption = Label3.Caption + "    : " + CStr(mem)

mem = 0
For temp = 0 To frmMain.ActiveForm.pro.ListCount - 1
If UCase(fso.GetExtensionName(frmMain.ActiveForm.pro.List(temp))) = "H" Then
mem = mem + 1
End If
Next
Label4.Caption = Label4.Caption + " : " + CStr(mem)

On Error Resume Next
Dim ff As file
Set ff = fso.GetFile(frmMain.ActiveForm.Caption)
Label5.Caption = Label5.Caption + CStr(ff.Size)
Label6.Caption = Label6.Caption + CStr(ff.DateCreated)
Label7.Caption = Label7.Caption + CStr(ff.DateLastModified)
End Sub

