VERSION 5.00
Begin VB.Form dorun 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Run menu"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Argument"
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton ok 
      Caption         =   "Start Debug"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
   End
   Begin VB.ListBox lstdeb 
      Appearance      =   0  'Flat
      Height          =   3870
      IntegralHeight  =   0   'False
      Left            =   4320
      MultiSelect     =   1  'Simple
      TabIndex        =   5
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton cmddeb 
      Caption         =   ">> Debug >>"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmddone 
      Caption         =   "Done"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdrun 
      Caption         =   "Run"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   855
   End
   Begin VB.ListBox lstrun 
      Appearance      =   0  'Flat
      Height          =   3870
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Error found in Following files"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Following files were Compiled Successfully"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "dorun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim debugs As Boolean
Private Sub cmddeb_Click()
debugs = Not debugs
If debugs = True Then
Me.Width = 8625
cmddeb.Caption = "<< Debug <<"
Else
Me.Width = 4395
cmddeb.Caption = ">> Debug >>"
End If
End Sub

Function readdat(compiler As String, inputs As String)
On Error GoTo endf
Dim fso As New FileSystemObject
Dim char As TextStream
Dim add As Boolean
Set char = fso.OpenTextFile(workpro + "\compiler\" + compiler + "\Misc\deb.con", ForReading)
strs1 = char.ReadLine
strs2 = char.ReadLine
stre = char.ReadLine
sepr = char.ReadLine

Dim txt As TextStream
Set txt = fso.OpenTextFile("c:\windows\temp\" + inputs + ".txt", ForReading, True)

Do While txt.AtEndOfStream = False
temp = txt.ReadLine
mems = Len(temp)
If Not InStr(mems, temp, ":") > 0 Then
If InStr(1, temp, strs1) > 0 Or InStr(1, temp, strs2) > 0 Then
If InStr(1, temp, stre) > 0 Then
If InStr(1, temp, sepr) > 0 Then

id = vbNullString
If UBound(Split(temp, strs1)) > 0 Then
filename = Split(temp, strs1)(0)
Else
filename = Split(temp, strs2)(0)
End If

If fso.GetExtensionName(filename) = "" Then
temp1 = Split(fso.GetFileName(filename), ".")(0)
With frmMain.ActiveForm.pro
    For tempx = 0 To .ListCount - 1
    If UCase(temp1) = UCase(Split(.List(tempx), ".")(0)) Then
    filename = .List(tempx)
    End If
    Next
End With
End If

done:
id = vbNullString

 
'check if filename already exists
For temx = 0 To lstdeb.ListCount - 1
If UCase(lstdeb.List(temx)) = UCase(fso.GetFileName(filename)) Then
id = temx
End If
Next

If id = vbNullString Then
frmMain.ActiveForm.errfiles.AddItem (vbNullString)
frmMain.ActiveForm.errfiles.List(frmMain.ActiveForm.errfiles.ListCount - 1) = fso.GetFileName(filename)
lstdeb.AddItem (vbNullString)
lstdeb.List(lstdeb.ListCount - 1) = fso.GetFileName(filename)
lstdeb.Selected(lstdeb.ListCount - 1) = True
frmMain.ActiveForm.errf.AddItem (vbNullString)
id = lstdeb.ListCount - 1
End If

mem = "EMPTY"
If UBound(Split(temp, strs1)) > 0 Then
mem = Split(Split(temp, strs1)(1), stre)(0)
End If

If mem = "EMPTY" Then
If UBound(Split(temp, strs2)) > 0 Then
mem = Split(Split(temp, strs2)(1), stre)(0)
End If
End If

Errr = Split(temp, sepr)(UBound(Split(temp, sepr)))
mem1 = CStr(frmMain.ActiveForm.errf.List(id))

add = True
For temp = 0 To UBound(Split(mem1, "/"))
If UCase(CStr(mem) + " : " + CStr(Errr)) = UCase(Split(mem1, "/")(temp)) Then
add = False
GoTo next1
End If
Next

next1:
'if adding=true then this statement is executed
If add = True Then
If mem1 = "" Or mem1 = 0 Then
mem2 = CStr(mem) + " : " + CStr(Errr)
Else
mem2 = mem1 + "/" + CStr(mem) + " : " + CStr(Errr)
End If
Else
'if adding=false then this statement is executed
mem2 = mem1
End If
frmMain.ActiveForm.errf.List(id) = mem2
                        
End If
End If
End If
End If
Loop


txt.Close
char.Close
endf:
End Function

Private Sub cancel_Click()
frmMain.ActiveForm.Frame.Visible = False
debugging = False
frmMain.ActiveForm.remake
Unload Me
End Sub

Private Sub cmddone_Click()
Unload Me
End Sub

Private Sub cmdrun_Click()
If Not lstdeb.ListIndex = -1 Then
RunEXE comp, lstrun.List(lstrun.ListIndex)
End If
End Sub

Private Sub Command1_Click()
mem = InputBox("Enter the arguments with EXE", "Argument", frmMain.ActiveForm.getarg)
If Not mem = vbNullString Then
frmMain.ActiveForm.setarg (CStr(mem))
End If
End Sub

Private Sub Form_Activate()
If lstdeb.ListCount = 0 Then
cmddeb.Enabled = False
ok.Enabled = False
End If
End Sub

Function done()
ok_Click
End Function

Private Sub Form_Load()
debugs = False
End Sub

Private Sub ok_Click()

frmMain.ActiveForm.errors.Clear

For temp = lstdeb.ListCount - 1 To 0 Step -1
If lstdeb.Selected(temp) = False Then
frmMain.ActiveForm.errfiles.RemoveItem (temp)
frmMain.ActiveForm.errf.RemoveItem (temp)
End If
Next

On Error Resume Next
frmMain.ActiveForm.errfiles.ListIndex = 0
frmMain.ActiveForm.errors.ListIndex = 0
If Not frmMain.ActiveForm.errfiles.ListCount = 0 Then
frmMain.ActiveForm.Frame.Visible = True
debugging = True
frmMain.ActiveForm.remake
End If
frmMain.ActiveForm.ferr
Unload Me
End Sub
