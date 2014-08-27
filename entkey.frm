VERSION 5.00
Begin VB.Form entkey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Licensing"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   1935
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continue"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter Key"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "This is valid for 30 days to continue afterwards you have to enter Key"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "entkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim done As Boolean

Private Sub Command1_Click()
Dim fso As New FileSystemObject
Dim txt As TextStream
top:

mem = InputBox("Enter key Here", "Licensing", "", (Screen.Width / 2) - 5000, (Screen.Height / 2) - 1500)
If Not mem = vbNullString Then
If mem = "1284-4802-4778-4568" Then
Set txt = fso.OpenTextFile("c:\windows\unonk.dat", ForWriting, True)
txt.WriteLine "lic"
MsgBox "Correct Key. Done", vbInformation, "Key"
Label3.Caption = ""
Command2_Click
Else
MsgBox "Wrong licensing key", vbInformation, "Key"
GoTo top
End If
End If
End Sub

Private Sub Command2_Click()
If Not Label3.Caption = "Your trial has expired" Then
frmMain.Show
done = False
Unload Me
End If
End Sub

Private Sub Command3_Click()
frmMain.unloads
Unload frmMain
done = False
Unload Me
End Sub

Private Sub Form_Load()
done = True
Dim fso As New FileSystemObject
Dim txt As TextStream

If fso.FileExists("c:\windows\unonk.dat") = False Then
Set txt = fso.OpenTextFile("c:\windows\unonk.dat", ForWriting, True)
txt.WriteLine (Date)
txt.WriteLine ("30")
txt.Close
End If

Set txt = fso.OpenTextFile("c:\windows\unonk.dat", ForReading, False)
temp = txt.ReadLine

If temp = "lic" Then
Command2_Click
GoTo endf
End If

temp1 = txt.ReadLine
txt.Close

If CDate(Date) > CDate(temp) Then
Set txt = fso.OpenTextFile("C:\windows\unonk.dat", ForWriting, True)
txt.WriteLine (Date)
txt.WriteLine (temp1 - (CDate(Date) - CDate(temp)))
txt.Close
End If



Set txt = fso.OpenTextFile("c:\windows\unonk.dat", ForReading, False)
temp = txt.ReadLine
temp1 = txt.ReadLine
txt.Close

If CDate(Date) < CDate(temp) Then
Label3.Caption = "Your trial has expired"
Command2.Enabled = False
Else
Label3.Caption = "You have " + temp1 + " days remaining"
If temp1 <= 0 Then
Label3.Caption = "Your trial has expired"
Command2.Enabled = False
End If
End If

endf:
End Sub

Private Sub Form_Unload(cancel As Integer)
cancel = done
End Sub
