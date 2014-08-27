VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form eded 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editor Settings"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Restore Defaults"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin MSComctlLib.Slider err 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393216
      Max             =   100
   End
   Begin MSComctlLib.Slider ex 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   50
      SmallChange     =   10
      Max             =   100
   End
   Begin VB.Label Label2 
      Caption         =   "ErrorBox Width"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Project Explorer Width"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "eded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
ex.Value = (2 / 7) * frmMain.ActiveForm.Width
err.Value = (2 / 7) * frmMain.ActiveForm.Height
frmMain.ActiveForm.seterrboxwid (err.Value)
frmMain.ActiveForm.setexpwid (ex.Value)
frmMain.ActiveForm.resize
End Sub

Private Sub err_Change()
frmMain.ActiveForm.seterrboxwid (err.Value)
frmMain.ActiveForm.resize
SaveSetting App.Title, "editor", "errbox", err.Value
End Sub

Private Sub ex_Change()
frmMain.ActiveForm.setexpwid (ex.Value)
frmMain.ActiveForm.resize
SaveSetting App.Title, "editor", "exbox", ex.Value
End Sub

Private Sub Form_Load()
ex.Max = frmMain.ActiveForm.Width * 0.7
err.Max = frmMain.ActiveForm.Height * 0.7
ex.Value = frmMain.ActiveForm.getexwid
err.Value = frmMain.ActiveForm.geterrwid
End Sub
