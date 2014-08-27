VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   4920
   ClientLeft      =   2565
   ClientTop       =   1380
   ClientWidth     =   5505
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox sas 
      Caption         =   "Show Tips at Startup"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      Height          =   4215
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5295
      Begin VB.CheckBox autoi 
         Caption         =   "Auto Indent"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   2640
         Width           =   2775
      End
      Begin VB.CheckBox autosp 
         Caption         =   "Auto Spaces"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   2280
         Width           =   2775
      End
      Begin VB.CheckBox dis 
         Caption         =   "Disable all Services"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox comfol 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   3975
      End
      Begin VB.CommandButton brow 
         Caption         =   "Browse..."
         Height          =   255
         Left            =   4200
         TabIndex        =   11
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox sdeb 
         Caption         =   "Show Debug Menu"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Value           =   1  'Checked
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Default Compile Folder"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   2055
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   2040
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   4455
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub brow_Click()
Load folbro
folbro.Show vbModal
End Sub

Private Sub cmdApply_Click()
debmenu = sdeb.Value
cfolder = comfol.Text
disabled = dis.Value
autospc = autosp.Value
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
debmenu = CBool(sdeb.Value)
cfolder = comfol.Text
disabled = dis.Value
autospc = autosp.Value
autoind = autoi.Value
ShowAtStart = sas.Value
Unload Me
End Sub


Private Sub Form_Load()
temp = CInt(debmenu)
If temp <> 0 Then temp = 1
sdeb.Value = temp

comfol.Text = cfolder

temp = CInt(autospc)
If temp <> 0 Then temp = 1
autosp.Value = temp

temp = CInt(disabled)
If temp <> 0 Then temp = 1
dis.Value = temp

temp = CInt(autoind)
If temp <> 0 Then temp = 1
autoi.Value = temp

temp = CInt(ShowAtStart)
If temp <> 0 Then temp = 1
sas.Value = temp


End Sub

