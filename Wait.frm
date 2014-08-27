VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Wait 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Wait"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   600
      Top             =   600
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2520
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   8
      Left            =   3600
      Top             =   480
   End
   Begin MSComctlLib.ProgressBar pro 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   60
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   800
      Width           =   4455
   End
   Begin VB.Label present 
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Please Wait ..."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Wait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim done As Boolean
Dim patch As Boolean

Function patched(data As Boolean)
patch = data
End Function

Private Sub Form_Unload(cancel As Integer)
Beep
If done = False Then cancel = True
End Sub

Private Sub Timer1_Timer()
Label2.Caption = CStr(Round((pro.Value / pro.Max) * 100)) + "% Complete"
GoTo nexts11
If status = 1 Then
present.Caption = "Creating ... " + frmMain.ActiveForm.pro.List(frmMain.ActiveForm.pro.ListIndex)
ElseIf status = 2 Then
For temp = 0 To frmMain.ActiveForm.pro.ListCount - 1
If frmMain.ActiveForm.pro.Selected(temp) = True Then
If pro.Value / 120 < (temp + 1) / frmMain.ActiveForm.pro.ListCount Then
mem = temp
GoTo nexts1
End If
End If
Next
nexts1:
present.Caption = "Creating ... " + frmMain.ActiveForm.pro.List(mem)

ElseIf status = 3 Then
For temp = 0 To frmMain.ActiveForm.pro.ListCount - 1
If pro.Value / 120 < (temp + 1) / frmMain.ActiveForm.pro.ListCount Then
mem = temp
GoTo nexts
End If
Next
nexts:
present.Caption = "Creating ... " + frmMain.ActiveForm.pro.List(mem)
End If

nexts11:

If Not pro.Value = 60 Then
pro.Value = pro.Value + 1
Else
done = True
If patch = True Then Unload Me
End If
End Sub

Private Sub Timer2_Timer()
Dim fso As New FileSystemObject
If fso.FileExists("c:\windows\temp\" + filess + ".txt") = True Then
Timer3.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
Unload Me
End Sub
