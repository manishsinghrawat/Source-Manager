VERSION 5.00
Begin VB.Form stp 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Timer tmr 
      Interval        =   500
      Left            =   1200
      Top             =   840
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Showing Result..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "stp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub tmr_Timer()
tmr.Enabled = False
Unload Me
End Sub
