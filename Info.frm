VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Info 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Information Dialog"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox data 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Info.frx":0000
   End
   Begin RichTextLib.RichTextBox shows 
      Height          =   6735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   11880
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Info.frx":008B
   End
   Begin VB.Timer timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2280
      Top             =   6960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Log to File"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6960
      Width           =   1815
   End
End
Attribute VB_Name = "Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim files As String
Dim done As Boolean
Dim ter As String

Private Sub Command1_Click()
Dim fso As New FileSystemObject
Dim txt As TextStream
Set txt = fso.OpenTextFile("c:\SM.inf", ForWriting, True)
txt.Write (data.Text)
MsgBox "Log Saved to C:\SM.inf", vbInformation
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
done = False
timer.Enabled = True
End Sub

Private Sub Form_Unload(cancel As Integer)
If done = False Then cancel = True
End Sub

Private Sub timer_Timer()
On Error Resume Next
If done = False Then
Dim fso As New FileSystemObject
Dim txt As TextStream

Set txt = fso.GetFile(files).OpenAsTextStream(ForReading)
data.Text = txt.ReadAll
data.find ter, 0, 10000000
If data.SelLength > 1 Then
done = True
End If
If Not data.Text = shows.Text Then
shows.Text = data.Text
End If
shows.SelStart = 10000000
End If

If done = True Then
shows.ScrollBars = rtfBoth
shows.Refresh
Command1.Enabled = True
Command2.Enabled = True
timer.Enabled = False
End If
End Sub

Function setfile(dat As String)
files = dat
End Function
Function termination(dat As String)
ter = dat
End Function


