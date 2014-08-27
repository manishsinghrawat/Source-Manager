VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form out 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Output Manager"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   1920
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   11430
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5040
      Top             =   120
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   3720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save Derived Text"
      Filter          =   "Text Files|*.txt"
      Orientation     =   2
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3240
      Top             =   0
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   7320
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   615
      Left            =   8880
      TabIndex        =   4
      Top             =   7080
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   7080
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">> ADD >>"
      Height          =   615
      Left            =   7080
      TabIndex        =   2
      Top             =   7080
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox rtf1 
      Height          =   6255
      Left            =   6000
      TabIndex        =   1
      Top             =   720
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   11033
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"out.frx":0000
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   11033
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"out.frx":0082
   End
   Begin VB.Label Label2 
      Caption         =   "Program Interactions Here"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   7080
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "User Selected Output"
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label lbl 
      Caption         =   "Current Program Status Screen"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "out"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
rtf1.SelText = rtf.Text + vbNewLine + vbNewLine
End Sub

Private Sub Command2_Click()
frmMain.ActiveForm.strrun
Timer1.Enabled = True
txt.SetFocus

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim fso As New FileSystemObject
fso.DeleteFile ("c:\temp.dat")
fso.DeleteFile ("c:\temp1.dat")

rtf.Text = vbNewLine + " No Output Currently Detected." + vbNewLine + vbNewLine + "  Possible Causes " + vbNewLine + vbNewLine + " 1) Program Not yet Started " + vbNewLine + " 2) Internal Error (File System)"
End Sub

Private Sub Form_Unload(cancel As Integer)
On Error Resume Next
Dim fso As New FileSystemObject
Dim txt As TextStream
msg = MsgBox("Do you want to Save Derived Output?", vbQuestion + vbYesNoCancel)
If msg = vbYes Then
cdlg.ShowSave
Set txt = fso.OpenTextFile(cdlg.filename, ForWriting, True)
txt.Write rtf1.Text
ElseIf msg = vbCancel Then
cancel = True
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim fso As New FileSystemObject
Dim txt As TextStream
If fso.FileExists("c:\temp.dat") = True Then
fso.CopyFile "c:\temp.dat", "c:\temp1.dat", True
Set txt = fso.OpenTextFile("c:\temp1.dat", ForReading, False)
temp = txt.ReadAll

If Not temp = rtf.Text Then
rtf.Text = temp
rtf.SelStart = Len(rtf.Text)
End If
Else
temp = vbNewLine + " No Output Currently Detected." + vbNewLine + vbNewLine + "  Possible Causes " + vbNewLine + vbNewLine + " 1) Program Not yet Started " + vbNewLine + " 2) Internal Error (File System)"

If Not temp = rtf.Text Then
rtf.Text = temp
End If

End If
Exit Sub
errs:
End Sub

Private Sub Timer2_Timer()
txt.SetFocus
Timer2.Enabled = False
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
Dim fso As New FileSystemObject
On Error GoTo endf
AppActivate ("e:\windows\system32\cmd.exe")
SendKeys CStr(Chr(KeyAscii))
Timer2.Enabled = True
Exit Sub
endf:
Timer1.Enabled = False
If fso.FileExists("c:\temp.dat") Then
fso.DeleteFile "c:\temp.dat"
fso.DeleteFile "c:\temp1.dat"
End If
rtf.Text = "Program Closed Or ended"
End Sub
