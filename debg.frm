VERSION 5.00
Begin VB.Form debg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Debug"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton cancel 
      Caption         =   "Discard"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton ok 
      Caption         =   "Debug"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   1455
   End
   Begin VB.ListBox deb 
      Appearance      =   0  'Flat
      Height          =   3750
      IntegralHeight  =   0   'False
      Left            =   120
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label lbl 
      Caption         =   "We are sorry for inconvienience but there was an error found in Debugging"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Check files to debug "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Following file are found to have errors"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "debg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function readdat(compiler As String, inputs As String)
On Error GoTo error
Dim fso As New FileSystemObject
Dim char As TextStream
Dim add As Boolean
Set char = fso.OpenTextFile(workpro + "\compiler\" + compiler + "\Misc\deb.con", ForReading)
strs1 = char.ReadLine
strs2 = char.ReadLine
stre = char.ReadLine
sepr = char.ReadLine

Dim txt As TextStream
Dim txt1 As TextStream

Set txt = fso.OpenTextFile("c:\windows\temp\" + inputs + ".txt", ForReading)
Set txt1 = fso.OpenTextFile("c:\windows\temp\" + inputs + ".txt", ForReading)
Text1.Text = txt1.ReadAll

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
For temx = 0 To deb.ListCount - 1
If UCase(deb.List(temx)) = UCase(fso.GetFileName(filename)) Then
id = temx
End If
Next

If id = vbNullString Then
frmMain.ActiveForm.errfiles.AddItem (vbNullString)
frmMain.ActiveForm.errfiles.List(frmMain.ActiveForm.errfiles.ListCount - 1) = fso.GetFileName(filename)
deb.AddItem (vbNullString)
deb.List(deb.ListCount - 1) = fso.GetFileName(filename)
deb.Selected(deb.ListCount - 1) = True
frmMain.ActiveForm.errf.AddItem (vbNullString)
id = deb.ListCount - 1
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

GoTo endf
error:
MsgBox (err.Description + " " + CStr(err.Number) + " : " + "Compiler not accessible!")
Unload Me
endf:
End Function


Private Sub cancel_Click()
frmMain.ActiveForm.Frame.Visible = False
debugging = False
frmMain.ActiveForm.remake
Unload Me
End Sub

Function done()
ok_Click
End Function

Private Sub Form_Activate()
If deb.ListCount = 0 Then
deb.Visible = False
Text1.Visible = True
lbl.Visible = True
Label1.Visible = False
Label2.Visible = False
ok.Enabled = False
Else
deb.Visible = True
Text1.Visible = False
lbl.Visible = False
Label1.Visible = True
Label2.Visible = True
ok.Enabled = True
End If
End Sub

Private Sub ok_Click()
On Error Resume Next
frmMain.ActiveForm.errors.Clear

For temp = deb.ListCount - 1 To 0 Step -1
If deb.Selected(temp) = False Then
frmMain.ActiveForm.errfiles.RemoveItem (temp)
frmMain.ActiveForm.errf.RemoveItem (temp)
End If
Next

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

