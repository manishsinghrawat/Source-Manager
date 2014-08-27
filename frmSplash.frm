VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmSplash.frx":0000
   PaletteMode     =   2  'Custom
   Picture         =   "frmSplash.frx":79FA
   ScaleHeight     =   4815
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox colon 
      Height          =   285
      Left            =   4320
      TabIndex        =   8
      Text            =   """"
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   800
      Left            =   6240
      Top             =   480
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmer : Manish"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   10
      Top             =   3120
      Width           =   4815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No compiler is included in this package compilers belong to their developers"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   4320
      Width           =   5415
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2280
      TabIndex        =   0
      Tag             =   "Version"
      Top             =   0
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label lblProductName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source Manager"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   33
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   840
      TabIndex        =   1
      Tag             =   "Product"
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   4815
      Left            =   0
      Top             =   0
      Width           =   7455
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " >> Searching For Compilers"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " >> Checking Compilers"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " >> Applying Options"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " >> Done"
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " >> Launching Mains"
      Height          =   255
      Index           =   4
      Left            =   4920
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label12c 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This is only Test Version of this project. Actual working is not fully verified"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   4080
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   360
      Picture         =   "frmSplash.frx":97B8
      Top             =   3960
      Width           =   735
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num As Integer

Private Sub Form_Click()
num = 4
Timer2_Timer
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = "Source Manager"
    End Sub


Sub chcol()
staR = 255
staG = 255
staB = 255

endR = 150
endG = 150
enb = 150

consR = (endR - staR) / Me.Width
consG = (endG - staG) / Me.Width
consB = (enb - staB) / Me.Width

For temp = 0 To Me.Width
colR = staR + (temp * consR)
colG = staG + (temp * consG)
colB = staB + (temp * consB)
Line (temp, 0)-(temp, Me.Height), RGB(colR, colG, colB)
Next
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Label1(num).Visible = True
num = num + 1
If num = 5 Then
Load entkey
entkey.Show
Timer2.Enabled = False
Unload Me
End If
End Sub
