VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   AutoRedraw      =   -1  'True
   Caption         =   "frmDocument"
   ClientHeight    =   5205
   ClientLeft      =   2145
   ClientTop       =   2610
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5205
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.Timer ur 
      Interval        =   1000
      Left            =   480
      Top             =   3840
   End
   Begin VB.ListBox errf 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1320
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame 
      Caption         =   "Debug"
      Height          =   1485
      Left            =   120
      TabIndex        =   2
      Top             =   2145
      Visible         =   0   'False
      Width           =   4830
      Begin VB.CommandButton Closeit 
         Caption         =   "X"
         Height          =   255
         Left            =   4200
         TabIndex        =   6
         Top             =   120
         Width           =   495
      End
      Begin VB.ListBox errfiles 
         Appearance      =   0  'Flat
         Height          =   975
         IntegralHeight  =   0   'False
         ItemData        =   "frmDocument.frx":0442
         Left            =   120
         List            =   "frmDocument.frx":0444
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.ListBox errors 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         IntegralHeight  =   0   'False
         ItemData        =   "frmDocument.frx":0446
         Left            =   1320
         List            =   "frmDocument.frx":0448
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
   End
   Begin RichTextLib.RichTextBox rtftext 
      Height          =   1935
      Left            =   1320
      TabIndex        =   1
      Top             =   135
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3413
      _Version        =   393217
      BackColor       =   16777215
      HideSelection   =   0   'False
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   1e6
      TextRTF         =   $"frmDocument.frx":044A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3000
      Top             =   2400
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   1560
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   1
   End
   Begin VB.ListBox pro 
      Appearance      =   0  'Flat
      Height          =   1935
      IntegralHeight  =   0   'False
      ItemData        =   "frmDocument.frx":04CA
      Left            =   120
      List            =   "frmDocument.frx":04CC
      Style           =   1  'Checkbox
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim shows As Boolean
Dim exwid As Integer
'file current status
Dim filename() As String
Dim filedata() As String
Dim fileplc() As Integer
'turn time variables
Dim ttundo(100) As String
Dim ttredo(100) As String
Dim ttplc0(100) As Single
Dim ttplc1(100) As Single
'variable for cur file assignment
Dim pre As Integer
'argument variable
Dim arg As String
Dim blind As Boolean
'checks whether file is saved or not
Dim saved As Boolean
'width of errorbox
Dim errboxwid As Integer
Dim changed As Boolean
'undo redo temporary variable
Dim last As String
Dim urdis As Boolean

Function setsave(sav As Boolean)
saved = sav
End Function
Private Sub Closeit_Click()
Frame.Visible = False
debugging = False
remake
End Sub

Function ferr()
errFiles_Click
End Function

Private Sub errFiles_Click()
On Error Resume Next
errors.Clear
For temp = 0 To UBound(Split(errf.List(errfiles.ListIndex), "/"))
errors.AddItem (Split(errf.List(errfiles.ListIndex), "/")(temp))
Next

For temp = 0 To pro.ListCount - 1
If UCase(pro.List(temp)) = UCase(errfiles.List(errfiles.ListIndex)) Then
mem = temp
End If
Next

pro.ListIndex = mem
pro_Click
errors.ListIndex = 0
errors_Click
End Sub

Private Sub errors_Click()
On Error GoTo endf

'set file thro pro _click
For temp = 0 To pro.ListCount - 1
If UCase(pro.List(temp)) = UCase(errfiles.List(errfiles.ListIndex)) Then
mem = temp
End If
Next
If mem <> pro.ListIndex Then
pro.ListIndex = mem
pro_Click
End If

'get line number of error
Dim mem1 As Integer
mem1 = Split(errors.List(errors.ListIndex), ":")(0)

'split multiline text into array
texts = Split(rtftext.Text, vbNewLine)

'get length of string array before line number
length = 0
For temp = 0 To mem1 - 2
length = length + Len(texts(temp)) + 1
Next

rtftext.find texts(mem1 - 1), length - 2
endf:
End Sub

Private Sub Form_Activate()
frmMain.proex.Checked = pro.Visible
frmMain.tbToolBar.Buttons(12).Value = Abs(pro.Visible)
pre = 0
pro_Click
End Sub

Function cleararr(filen As String, filed As String, filep As Integer)
ReDim filename(0) As String
ReDim filedata(0) As String
ReDim fileplc(0) As Integer
filename(0) = filen
filedata(0) = filed
fileplc(0) = filep
End Function
Function addarrnam(data As String)
ReDim Preserve filename(UBound(filename) + 1)
filename(UBound(filename)) = data
End Function
Function addarrdata(data As String)
ReDim Preserve filedata(UBound(filedata) + 1)
filedata(UBound(filedata)) = data
End Function
Function addarrplc(plc As Integer)
ReDim Preserve fileplc(UBound(fileplc) + 1)
fileplc(UBound(fileplc)) = plc
End Function
Function getarrnam(id As Integer)
getarrnam = filename(id)
End Function
Function getarrdata(id As Integer)
getarrdata = filedata(id)
End Function
Function getarrplc(id As Integer)
getarrplc = fileplc(id)
End Function

Function getubounda()
getubounda = UBound(filename)
End Function
Function getuboundb()
getuboundb = UBound(filedata)
End Function
Function getuboundc()
getuboundc = UBound(fileplc)
End Function


Private Sub pro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu frmMain.pop, 1, X, Y
End If
End Sub


Public Sub rtftext_selChange()
If Not frmMain.ActiveForm Is Nothing Then
'split multiline text into array
texts = Split(frmMain.ActiveForm.rtftext.Text, vbNewLine)

ccklen = 0
nums = 0
curline = 1
For temp = 0 To UBound(texts)
ccklen = ccklen + Len(texts(temp)) + 2
If ccklen <= frmMain.ActiveForm.rtftext.SelStart Then
curline = curline + 1
nums = nums + Len(texts(temp)) + 2
'to avoid increment in size of ccklen greater than cur line
End If
Next

frmMain.pos.Text = "Ln " + CStr(curline) + " , Col " + CStr(frmMain.ActiveForm.rtftext.SelStart - nums)

If Not rtftext.TextRTF = last Then
ttredo(100) = vbNullString
End If

changed = True
End If
saved = False
End Sub



Private Sub rtftext_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
appall
End If
End Sub

Private Sub Timer1_Timer()
If pro.ListIndex = -1 Then pro.ListIndex = 0
End Sub


Private Sub Form_Unload(cancel As Integer)
On Error GoTo endf
If saved = False Then
xx = MsgBox("Do you want to save " + Me.Caption + "?", vbYesNoCancel)
If xx = vbCancel Then
cancel = True
ElseIf xx = vbYes Then
    
    Dim sFile As String
    If Left$(Me.Caption, 7) = "Project" Then
        With frmMain.dlgCommonDialog
            .DialogTitle = "Save"
            .CancelError = True
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = "C++ Projects (*.cpx)|*.cpx"
            .ShowSave
            If Len(.filename) = 0 Then
                Exit Sub
            End If
            sFile = .filename
        End With
        saves sFile
        
    Else
        sFile = Me.Caption
        saves sFile
    End If

End If
End If
endf:
End Sub

Public Sub pro_Click()
If Not pre = -1 Then
filedata(pre) = rtftext.TextRTF
fileplc(pre) = rtftext.SelStart
End If
rtftext.TextRTF = filedata(pro.ListIndex)
rtftext.SelStart = fileplc(pro.ListIndex)

pre = pro.ListIndex

frmMain.sel.Enabled = pro.SelCount

Dim fso As New FileSystemObject
frmMain.cur.Enabled = False
frmMain.curd.Enabled = False

For temp = 0 To 100
If UCase(fso.GetExtensionName(pro.List(pro.ListIndex))) = UCase(asso(temp, 0)) Then
frmMain.cur.Enabled = True
frmMain.curd.Enabled = True

For tex = 0 To 100
If compilers(tex, 0) = asso(temp, 1) Then
comp = compilers(tex, 1)
End If
Next
GoTo done
End If
Next
done:

On Error GoTo endf
'this statement loads keywords operators and indentation
Dim txt As TextStream
ReDim keywords(0)
ReDim operators(0)
Set txt = fso.OpenTextFile(App.path + "\compiler\" + comp + "\misc\keywords.con", ForReading)
temp = 1
Do While txt.AtEndOfStream = False
ReDim Preserve keywords(UBound(keywords) + 1)

keywords(temp) = txt.ReadLine
temp = temp + 1
Loop
txt.Close

'reads operator file
Set txt = fso.OpenTextFile(CStr(workpro) + "\compiler\" + comp + "\misc\operator.con")
temp = 1
Do While txt.AtEndOfStream = False
ReDim Preserve operators(UBound(operators) + 1)
operators(temp) = txt.ReadLine
temp = temp + 1
Loop
txt.Close

'reads indendation file
Set txt = fso.OpenTextFile(CStr(workpro) + "\compiler\" + comp + "\misc\operator.con")
indentation(0) = txt.ReadLine
indentation(1) = txt.ReadLine
txt.Close
endf:

xchange
If runrun = False Then appall
runrun = False

End Sub

Private Sub rtfText_Click()
change
saved = False
End Sub

Function xchange()
rtftext.Font.Size = rtftext.SelFontSize
rtftext.Font.name = rtftext.SelFontName
frmMain.FontSize.Text = CInt(rtftext.SelFontSize)
frmMain.Fontname.Text = rtftext.SelFontName

End Function

Private Sub change()
frmMain.FontSize.Text = CInt(rtftext.Font.Size)
frmMain.Fontname.Text = rtftext.Font.name
End Sub

Function remove()
pro.Refresh
mem = pro.ListIndex

For temp = mem To UBound(filename) - 1
pro.List(temp) = pro.List(temp + 1)
filename(temp) = filename(temp + 1)
filedata(temp) = filedata(temp + 1)
fileplc(temp) = fileplc(temp + 1)
Next

ReDim Preserve filename(UBound(filename) - 1)
ReDim Preserve filedata(UBound(filedata) - 1)
ReDim Preserve fileplc(UBound(fileplc) - 1)

rtftext.TextRTF = filedata(mem)
rtftext.SelStart = fileplc(mem)

pro.RemoveItem (pro.ListCount - 1)


End Function

Function remake()
Form_resize
pro_Click
End Function
Private Sub Form_Load()
changed = False
errboxwid = GetSetting(App.Title, "editor", "errbox", 3000)
exwid = GetSetting(App.Title, "editor", "exbox", 3000)
shows = projectex
pro.Visible = projectex

startup = "Main.cpp"
pre = 0
NEwS

    Form_resize
    On Error Resume Next
    saved = True
    last = frmMain.rtftext.TextRTF
End Sub
Function resize()
Form_resize
End Function
Function seterrboxwid(data As Integer)
errboxwid = data
End Function
Function setexpwid(data As Integer)
exwid = data
End Function
Function geterrwid()
geterrwid = errboxwid
End Function
Function getexwid()
getexwid = exwid
End Function
Private Sub Form_resize()

If shows = True Then
wid = exwid
Else
wid = 0
End If

On Error Resume Next
If debugging = True Then
    reduce = errboxwid
    Else
    reduce = 0
End If

    If shows = True Then
    pro.Width = wid
    rtftext.Move 200 + wid, 100, Me.ScaleWidth - 300 - wid, Me.ScaleHeight - 200 - reduce
    Else
    pro.Width = wid
    rtftext.Move 100 + wid, 100, Me.ScaleWidth - 200 - wid, Me.ScaleHeight - 200 - reduce
    End If
    pro.Height = rtftext.Height
    
    If debugging = True Then
    Frame.Visible = True
    Frame.top = pro.Height + pro.top + 100
    Frame.Left = pro.Left
    Frame.Width = Me.ScaleWidth - Frame.Left - 100
    Frame.Height = reduce - 100
    
    errfiles.Left = 100
    errfiles.Width = pro.Width - 100
    errfiles.Height = Frame.Height - errfiles.top - 200
    
    errors.Height = errfiles.Height
    errors.Left = errfiles.Left + errfiles.Width + errfiles.Left
    errors.Width = Frame.Width - (2 * errfiles.Left + errfiles.Width + 200)
    
    Closeit.top = 120
    Closeit.Left = Frame.Width - 700
    End If
    
End Sub

Function showex()
shows = Not shows
pro.Visible = shows
pro.Width = widthex
pro.top = 100
pro.Left = 100
pro.Height = Me.ScaleHeight - 100

Form_resize
End Function
Function setarg(Text As String)
arg = Text
End Function
Function getarg()
getarg = arg
End Function

'File system manager unit
Function NEwS()
ReDim filename(0)
ReDim filedata(0)
ReDim fileplc(0)
filename(0) = "Main.cpp"
filedata(0) = ""
fileplc(0) = 1
pre = 0
rebuild
pro_Click
End Function

Sub opens(path As String)

Dim fso As New FileSystemObject
Dim txt As TextStream
Dim txt1 As TextStream
Dim txt2 As TextStream
Set txt = fso.OpenTextFile(path, ForReading, False)
temp = False

Do While txt.AtEndOfStream = False
If temp = True Then
ReDim Preserve filename(UBound(filename) + 1)
ReDim Preserve filedata(UBound(filedata) + 1)
ReDim Preserve fileplc(UBound(fileplc) + 1)
Else
ReDim filename(0)
ReDim filedata(0)
ReDim fileplc(0)
temp = True
End If
result = vbNullString
filename(UBound(filename)) = txt.ReadLine
fileplc(UBound(fileplc)) = txt.ReadLine
mem = txt.ReadLine

Do While Not mem = "***///***"
If result = vbNullString Then
result = mem
Else
result = result + vbNewLine + mem
End If
mem = txt.ReadLine
Loop

filedata(UBound(filedata)) = result
Loop

txt.Close
rebuild
pro.ListIndex = 0
pre = -1
pro_Click
Exit Sub
endf:
MsgBox "Error occurred : Possible Cause : Invalid file specified", , "Critical"
NEwS
End Sub

Function saves(path As String)
On Error GoTo err
Dim fso As New FileSystemObject
Dim txt As TextStream
sav
If fso.FileExists(path) = False Then fso.CreateTextFile (path)
Set txt = fso.OpenTextFile(path, ForWriting, True)

For temp = 0 To UBound(filename)
txt.WriteLine filename(temp)
txt.WriteLine fileplc(temp)
txt.WriteLine filedata(temp)
txt.WriteLine ("***///***")
Next
GoTo endf
err:
MsgBox "Error occurred : Possible Cause : " + err.Description, vbCritical
endf:
End Function

'File system Ends
Function addfile()
tmem = pro.ListIndex
def = "Newcpp"
retry:
temp = InputBox("Enter the Cpp file name (Extension not required)", "New CPP File", def)

If validate(CStr(temp) + ".cpp") = True Then

If checkval(CStr(temp) + ".cpp") = True Then
If Not Len(temp) > 8 Then
If Not temp = vbNullString Then
ReDim Preserve filename(UBound(filename) + 1)
ReDim Preserve fileplc(UBound(fileplc) + 1)
ReDim Preserve filedata(UBound(filedata) + 1)
filename(UBound(filename)) = temp + ".cpp"
fileplc(UBound(fileplc)) = 0
filedata(UBound(filedata)) = ""
rebuild
End If
Else
MsgBox "File name greater than 8 characters", vbCritical
End If

Else
MsgBox "Invalid Character encountered. Please remove * / : ' < > ? \   spaces and try again", vbCritical
def = temp
GoTo retry
End If

Else
msg = MsgBox("File already exists do you want to add file with another name?", vbYesNo, "Conflict")
If msg = vbYes Then
def = temp
GoTo retry
End If
End If
pro.ListIndex = tmem
End Function

Function addcfile()
tmem = pro.ListIndex
def = "NewC"
retry:
temp = InputBox("Enter the C file name (Extension not required)", "New C File", def)

If validate(CStr(temp) + ".C") = True Then

If checkval(CStr(temp) + ".C") = True Then
If Not Len(temp) > 8 Then
If Not temp = vbNullString Then
ReDim Preserve filename(UBound(filename) + 1)
ReDim Preserve fileplc(UBound(fileplc) + 1)
ReDim Preserve filedata(UBound(filedata) + 1)
filename(UBound(filename)) = temp + ".c"
fileplc(UBound(fileplc)) = 0
filedata(UBound(filedata)) = ""
rebuild
End If
Else
MsgBox "File name greater than 8 characters", vbCritical
End If
Else
MsgBox "Invalid Character encountered. Please remove * / : ' < > ? \   spaces and try again", vbCritical
def = temp
GoTo retry
End If

Else
msg = MsgBox("File already exists do you want to add file with another name?", vbYesNo, "Conflict")
If msg = vbYes Then
def = temp
GoTo retry
End If

End If
pro.ListIndex = tmem
End Function
Function addhfile()
tmem = pro.ListIndex
def = "Newh"
retry:
temp = InputBox("Enter the Header file name (Extention not required)", "New Header File", def)

If validate(CStr(temp) + ".h") = True Then
If checkval(CStr(temp) + ".h") = True Then
If Not Len(temp) > 8 Then
If Not temp = vbNullString Then
ReDim Preserve filename(UBound(filename) + 1)
ReDim Preserve fileplc(UBound(fileplc) + 1)
ReDim Preserve filedata(UBound(filedata) + 1)
filename(UBound(filename)) = temp + ".h"
fileplc(UBound(fileplc)) = 0
filedata(UBound(filedata)) = ""
rebuild
End If
Else
MsgBox "File name greater than 8 characters", vbCritical
End If
Else
MsgBox "Invalid Character encountered. Please remove * / : ' < > ? \   spaces and try again", vbCritical
def = temp
GoTo retry
End If

Else
msg = MsgBox("File already exists do you want to add file with another name?", vbYesNo, "Conflict")
If msg = vbYes Then
def = temp
GoTo retry
End If
End If
pro.ListIndex = tmem
End Function

Function asfile(ext As String)
tmem = pro.ListIndex
'this block is used to add file of custom used choice extension
If Len(ext) = 0 Then Exit Function

def = ""
msg = vbYes

Do While msg = vbYes
msg = vbNo
temp = InputBox("Enter name of file you want to insert into project" + vbNullString + vbNewLine + "Extension : " + ext, "Add " + ext + " file", def)
If Len(temp) = 0 Then Exit Function
temp = temp + "." + ext

If Len(temp) = 0 Then Exit Function
If validate(CStr(temp)) = True Then
If checkval(CStr(temp)) = True Then
If Not Len(Split(temp, ".")(0)) > 8 Then
If Not temp = vbNullString Then
ReDim Preserve filename(UBound(filename) + 1)
ReDim Preserve fileplc(UBound(fileplc) + 1)
ReDim Preserve filedata(UBound(filedata) + 1)
filename(UBound(filename)) = temp
fileplc(UBound(fileplc)) = 0
filedata(UBound(filedata)) = ""
rebuild
End If
Else
MsgBox "File name greater than 8 characters", vbCritical
def = temp
msg = vbYes
End If
Else
MsgBox "Invalid Character encountered. Please remove * / : ' < > ? \   spaces and try again", vbCritical
def = temp
msg = vbYes
End If

Else
msg = MsgBox("File already exists do you want to add file with another name?", vbYesNo, "Conflict")
If msg = vbYes Then
def = temp
End If
End If
Loop
pro.ListIndex = tmem
End Function


Function cusfile()
tmem = pro.ListIndex
'this block is used to add file of custom used choice extension
def = ""
msg = vbYes

Do While msg = vbYes
msg = vbNo
temp = InputBox("Enter name of file you want to insert into project" + vbNewLine + vbNewLine + "Be sure that extension you enter is associated with any of compiler otherwise file will be unusable.", "Add custom file", def)
If Len(temp) = 0 Then Exit Function
If validate(CStr(temp)) = True Then
If checkval(CStr(temp)) = True Then
If Not Len(Split(temp, ".")(0)) > 8 Then
If Not temp = vbNullString Then
ReDim Preserve filename(UBound(filename) + 1)
ReDim Preserve fileplc(UBound(fileplc) + 1)
ReDim Preserve filedata(UBound(filedata) + 1)
filename(UBound(filename)) = temp
fileplc(UBound(fileplc)) = 0
filedata(UBound(filedata)) = ""
rebuild
End If
Else
MsgBox "File name greater than 8 characters", vbCritical
def = temp
msg = vbYes
End If
Else
MsgBox "Invalid Character encountered. Please remove * / : ' < > ? \   spaces and try again", vbCritical
def = temp
msg = vbYes
End If

Else
msg = MsgBox("File already exists do you want to add file with another name?", vbYesNo, "Conflict")
If msg = vbYes Then
def = temp
End If
End If
Loop
pro.ListIndex = tmem
End Function

Sub addext()
'block adds external file into activeproject
On Error GoTo endf
tmem = pro.ListIndex
Dim fso As New FileSystemObject
Dim txt As TextStream
Dim flname As String
Com.DialogTitle = "Add File"
Com.Filter = "All files(*.*)|*.*"
Com.CancelError = True
Com.ShowOpen
flname = Com.filename

If Len(Com.filename) = 0 Then GoTo endf
    newname = fso.GetFileName(flname)
If validate(fso.GetFileName(flname)) = False Then
    msg = MsgBox("File already exists do you want to add file with another name?", vbYesNo, "Conflict")
    If msg = vbYes Then
    newname = InputBox("Enter new filename (Extension Required!)", "Add", fso.GetFileName(flname))
    If newname = vbNullString Then GoTo endf
    ElseIf msg = vbNo Then
    GoTo endf
    End If
End If

ReDim Preserve filename(UBound(filename) + 1)
ReDim Preserve fileplc(UBound(fileplc) + 1)
ReDim Preserve filedata(UBound(filedata) + 1)
Set txt = fso.OpenTextFile(Com.filename, ForReading, False)
If Len(Split(newname, ".")(0)) > 8 Then
newname = Left(Split(newname, ".")(0), 8) + "." + Split(newname, ".")(1)
End If
filename(UBound(filename)) = newname
fileplc(UBound(fileplc)) = 0
filedata(UBound(filedata)) = txt.ReadAll
rebuild
pro.ListIndex = tmem
endf:
End Sub

Function exthea(flname As String)

tmem = pro.ListIndex
Dim fso As New FileSystemObject
Dim txt As TextStream

If Len(flname) = 0 Then GoTo endf
newname = fso.GetFileName(flname)
If validate(fso.GetFileName(flname)) = False Then
msg = MsgBox("File already exists do you want to add file with another name?", vbYesNo, "Conflict")
    If msg = vbYes Then
    newname = InputBox("Enter new filename (with extension)", "Add")
    If flname = vbNullString Then GoTo endf
    ElseIf msg = vbNo Then
    GoTo endf
    End If
End If

ReDim Preserve filename(UBound(filename) + 1)
ReDim Preserve fileplc(UBound(fileplc) + 1)
ReDim Preserve filedata(UBound(filedata) + 1)
Set txt = fso.OpenTextFile(flname, ForReading, False)
filename(UBound(filename)) = newname
fileplc(UBound(fileplc)) = 0
filedata(UBound(filedata)) = txt.ReadAll
rebuild
pro.ListIndex = tmem
endf:
End Function

Function rebuild()
pro.Clear
On Error Resume Next
For temp = 0 To 100
pro.AddItem (filename(temp))
Next
End Function
'/////////////////////////////////Compiling section///////////////////////////////

Sub curRUN()
Dim fso As New FileSystemObject
For temp = 0 To 100
If UCase(fso.GetExtensionName(pro.List(pro.ListIndex))) = UCase(asso(temp, 0)) Then
For tex = 0 To 100
If compilers(tex, 0) = asso(temp, 1) Then
compiler = compilers(tex, 1)
End If
Next
GoTo done
End If
Next
done:

If fso.FolderExists(workpro + "\compiler\" + comp) = True Then
frmMain.ActiveForm.errfiles.Clear
frmMain.ActiveForm.errors.Clear
frmMain.ActiveForm.errf.Clear

pro_Click
Load debg
For temp = 0 To pro.ListCount - 1
BuildFile filename(temp), filedata(temp)
Next

consEXE CStr(compiler), filename(pro.ListIndex)

mem = fso.FileExists("c:\windows\temp\cpp\" + Split(filename(pro.ListIndex), ".")(0) + ".exe")

If mem = False Then
debg.deb.Clear
debg.readdat CStr(compiler), CStr(Split(filename(pro.ListIndex), ".")(0))

If debmenu = True Then
debg.Show vbModal
Else
debg.done
End If
End If

If mem = True Then
Load stp
stp.Show vbModal
RunEXE CStr(compiler), filename(pro.ListIndex)
End If
Else
MsgBox "Error Found : Possible Cause : Corrupted or missing Compiler Files", vbCritical
End If
End Sub

Sub OutRUN()
Dim fso As New FileSystemObject
For temp = 0 To 100
If UCase(fso.GetExtensionName(pro.List(pro.ListIndex))) = UCase(asso(temp, 0)) Then
For tex = 0 To 100
If compilers(tex, 0) = asso(temp, 1) Then
compiler = compilers(tex, 1)
End If
Next
GoTo done
End If
Next
done:

If fso.FolderExists(workpro + "\compiler\" + comp) = True Then
frmMain.ActiveForm.errfiles.Clear
frmMain.ActiveForm.errors.Clear
frmMain.ActiveForm.errf.Clear

pro_Click
Load debg
For temp = 0 To pro.ListCount - 1
BuildFile filename(temp), filedata(temp)
Next

consEXE CStr(compiler), filename(pro.ListIndex)

mem = fso.FileExists("c:\windows\temp\cpp\" + Split(filename(pro.ListIndex), ".")(0) + ".exe")

If mem = False Then
debg.deb.Clear
debg.readdat CStr(compiler), CStr(Split(filename(pro.ListIndex), ".")(0))

If debmenu = True Then
debg.Show vbModal
Else
debg.done
End If
End If
If mem = True Then
Load out
out.Show vbModal

End If


Else
MsgBox "Error Found : Possible Cause : Corrupted or missing Compiler Files", vbCritical
End If

End Sub

Function strrun()
Dim fso As New FileSystemObject
For temp = 0 To 100
If UCase(fso.GetExtensionName(pro.List(pro.ListIndex))) = UCase(asso(temp, 0)) Then
For tex = 0 To 100
If compilers(tex, 0) = asso(temp, 1) Then
compiler = compilers(tex, 1)
End If
Next
GoTo done
End If
Next
done:
Load stp
stp.Show vbModal
outRunEXE CStr(compiler), filename(pro.ListIndex)
End Function

Function runall()
Dim compiler As String
Dim fso As New FileSystemObject
If fso.FolderExists(workpro + "\compiler\" + compiler) = True Then
frmMain.ActiveForm.errfiles.Clear
frmMain.ActiveForm.errors.Clear
frmMain.ActiveForm.errf.Clear

Dim errors As Boolean

pro_Click
Load dorun
dorun.lstrun.Clear
dorun.lstdeb.Clear
temp = pro.ListCount - 1

For temp = 0 To pro.ListCount - 1
BuildFile filename(temp), filedata(temp)
Next

For temp = 0 To UBound(filename)
'main process Call
For tex = 0 To 100
If UCase(fso.GetExtensionName(filename(temp))) = UCase(asso(tex, 0)) Then
For tex1 = 0 To 100
If compilers(tex1, 0) = asso(tex, 1) Then
compiler = compilers(tex1, 1)
End If
Next
consEXE compiler, filename(temp)
End If
Next
Next



For temp = 0 To UBound(filename)
errors = False
'main process Call
For tex = 0 To 100
If UCase(fso.GetExtensionName(filename(temp))) = UCase(asso(tex, 0)) Then
For tex1 = 0 To 100
If compilers(tex1, 0) = asso(tex, 1) Then
compiler = compilers(tex1, 1)
End If
Next
'check if exe file exists
mem = fso.FileExists("c:\windows\temp\cpp\" + Split(filename(temp), ".")(0) + ".exe")

'sets up list containing built file name
If fso.FileExists("c:\windows\temp\cpp\" + Split(filename(temp), ".")(0) + ".exe") = True Then
dorun.lstrun.AddItem (filename(temp))
End If

'initialises error reding section
If mem = False Then dorun.readdat compiler, CStr(Split(filename(temp), ".")(0))

End If
Next
Next

On Error Resume Next
dorun.lstrun.ListIndex = 0
dorun.Show vbModal
Else
MsgBox "Error Found : Possible Cause : Corrupted or missing Compiler Files", vbCritical
End If
End Function

Function selrun()
Dim compiler As String
Dim fso As New FileSystemObject
If fso.FolderExists(workpro + "\compiler\" + compiler) = True Then
frmMain.ActiveForm.errfiles.Clear
frmMain.ActiveForm.errors.Clear
frmMain.ActiveForm.errf.Clear

Dim errors As Boolean

pro_Click
Load dorun
dorun.lstrun.Clear
dorun.lstdeb.Clear

For temp = 0 To pro.ListCount - 1
BuildFile filename(temp), filedata(temp)
Next

For temp = 0 To UBound(filename)
If frmMain.ActiveForm.pro.Selected(temp) = True Then
'main process Call
For tex = 0 To 100
If UCase(fso.GetExtensionName(filename(temp))) = UCase(asso(tex, 0)) Then
For tex1 = 0 To 100
If compilers(tex1, 0) = asso(tex, 1) Then
compiler = compilers(tex1, 1)
End If
Next
consEXE compiler, filename(temp)
End If
Next
End If
Next



For temp = 0 To UBound(filename)
If frmMain.ActiveForm.pro.Selected(temp) = True Then
errors = False
'main process Call
For tex = 0 To 100
If UCase(fso.GetExtensionName(filename(temp))) = UCase(asso(tex, 0)) Then
For tex1 = 0 To 100
If compilers(tex1, 0) = asso(tex, 1) Then
compiler = compilers(tex1, 1)
End If
Next
'check if exe file exists
mem = fso.FileExists("c:\windows\temp\cpp\" + Split(filename(temp), ".")(0) + ".exe")

'sets up list containing built file name
If fso.FileExists("c:\windows\temp\cpp\" + Split(filename(temp), ".")(0) + ".exe") = True Then
dorun.lstrun.AddItem (filename(temp))
End If

'initialises error reding section
If mem = False Then dorun.readdat compiler, CStr(Split(filename(temp), ".")(0))

End If
Next
End If
Next




On Error Resume Next
dorun.lstrun.ListIndex = 0
dorun.Show vbModal

Else
MsgBox "Error Found : Possible Cause : Corrupted or missing Compiler Files", vbCritical
End If
End Function

Function debgc()
Dim compiler As String
pro_Click
compiler = comp
Dim fso As New FileSystemObject
If fso.FolderExists(workpro + "\compiler\" + compiler) = True Then
frmMain.ActiveForm.errfiles.Clear
frmMain.ActiveForm.errors.Clear
frmMain.ActiveForm.errf.Clear

pro_Click
Load debg
debg.deb.Clear

For temp = 0 To pro.ListCount - 1
BuildFile filename(temp), filedata(temp)
Next

'main process Call
consEXE compiler, filename(pro.ListIndex)

mem = fso.FileExists("c:\windows\temp\cpp\" + Split(filename(pro.ListIndex), ".")(0) + ".exe")

If mem = False Then
debg.readdat compiler, CStr(Split(filename(pro.ListIndex), ".")(0))
If debmenu = False Then
debg.done
Else
debg.Show vbModal
End If
Else
MsgBox "Check Done : All FIles OK : No Errors", vbInformation

End If
Else

MsgBox "Error Found : Possible Cause : Corrupted or missing Compiler Files", vbCritical
End If
End Function

Function deball(compiler As String)
Dim fso As New FileSystemObject
If fso.FolderExists(workpro + "\compiler\" + compiler) = True Then
frmMain.ActiveForm.errfiles.Clear
frmMain.ActiveForm.errors.Clear
frmMain.ActiveForm.errf.Clear

Dim errors As Boolean

pro_Click
Load debg
debg.deb.Clear

For temp = 0 To pro.ListCount - 1
BuildFile filename(temp), filedata(temp)
Next

For temp = 0 To UBound(filename)
'main process Call
For tex = 0 To 100
If UCase(fso.GetExtensionName(filename(temp))) = UCase(asso(tex, 0)) Then
For tex1 = 0 To 100
If compilers(tex1, 0) = asso(tex, 1) Then
compiler = compilers(tex1, 1)
End If
Next
consEXE compiler, filename(temp)
End If
Next
Next

For temp = 0 To UBound(filename)
errors = False
'main process Call
For tex = 0 To 100
If UCase(fso.GetExtensionName(filename(temp))) = UCase(asso(tex, 0)) Then
For tex1 = 0 To 100
If compilers(tex1, 0) = asso(tex, 1) Then
compiler = compilers(tex1, 1)
End If
Next
'check if exe file exists
mem = fso.FileExists("c:\windows\temp\cpp\" + Split(filename(temp), ".")(0) + ".exe")

'initialises error reding section
If mem = False Then debg.readdat compiler, CStr(Split(filename(temp), ".")(0))

End If
Next
Next

If mem = False Then
On Error Resume Next
If debmenu = False Then
debg.done
Else
debg.deb.ListIndex = 0
debg.Show vbModal
End If
Else
MsgBox "Check Done : All FIles OK : No Errors", vbInformation
End If
Else
MsgBox "Error Found : Possible Cause : Corrupted or missing Compiler Files", vbCritical
End If
End Function

Sub compiles(fol As String)
Dim fso As New FileSystemObject
If fso.FolderExists(workpro + "\compiler\" + comp) = True Then
frmMain.ActiveForm.errfiles.Clear
frmMain.ActiveForm.errors.Clear
frmMain.ActiveForm.errf.Clear

pro_Click
Load debg
For temp = 0 To pro.ListCount - 1
BuildFile filename(temp), filedata(temp)
Next

'main process Call
consEXE CStr(comp), filename(pro.ListIndex)


'cheks whether file exists or not
mem = fso.FileExists("c:\windows\temp\cpp\" + Split(filename(pro.ListIndex), ".")(0) + ".exe")

If Split(fol, "\")(UBound(Split(fol, "\"))) = vbNullString Then
fol = fol
Else
fol = fol + "\"
End If


If mem = False Then
 MsgBox "Error Found: Ignored!"
GoTo endf
Else
fso.CopyFile "c:\windows\temp\cpp\" + Split((filename(pro.ListIndex)), ".")(0) + ".exe", fol + Split((filename(pro.ListIndex)), ".")(0) + ".exe", True
End If

Else
MsgBox "Error Found : Possible Cause : Corrupted or missing Compiler Files", vbCritical
End If
endf:
Unload debg
End Sub


Function selcomp(fol As String)
Dim fso As New FileSystemObject
If fso.FolderExists(workpro + "\compiler\" + compiler) = True Then
frmMain.ActiveForm.errfiles.Clear
frmMain.ActiveForm.errors.Clear
frmMain.ActiveForm.errf.Clear

Dim errors As Boolean

pro_Click
Load dorun
dorun.lstrun.Clear
dorun.lstdeb.Clear

For temp = 0 To pro.ListCount - 1
BuildFile filename(temp), filedata(temp)
Next

If Split(fol, "\")(UBound(Split(fol, "\"))) = vbNullString Then
fol = fol
Else
fol = fol + "\"
End If

For temp = 0 To UBound(filename)
If frmMain.ActiveForm.pro.Selected(temp) = True Then
'main process Call
For tex = 0 To 100
If UCase(fso.GetExtensionName(filename(temp))) = UCase(asso(tex, 0)) Then
For tex1 = 0 To 100
If compilers(tex1, 0) = asso(tex, 1) Then
compiler = compilers(tex1, 1)
End If
Next
consEXE CStr(compiler), filename(temp)
End If
Next
End If
Next

For temp = 0 To UBound(filename)
If frmMain.ActiveForm.pro.Selected(temp) = True Then
errors = False
'main process Call
For tex = 0 To 100
If UCase(fso.GetExtensionName(filename(temp))) = UCase(asso(tex, 0)) Then
For tex1 = 0 To 100
If compilers(tex1, 0) = asso(tex, 1) Then
compiler = compilers(tex1, 1)
End If
Next

'check if exe file exists
mem = fso.FileExists("c:\windows\temp\cpp\" + Split(filename(temp), ".")(0) + ".exe")

'initialises error reding section
If mem = False Then
errors = True
Else
fso.CopyFile "c:\windows\temp\cpp\" + Split((filename(temp)), ".")(0) + ".exe", fol + Split((filename(temp)), ".")(0) + ".exe", True
End If

End If
Next
End If
Next


If errors = True Then MsgBox "Error Found: Ignored!"

Else
MsgBox "Error Found : Possible Cause : Corrupted or missing Compiler Files", vbCritical
End If
Unload dorun
End Function

Function compall(fol As String)
Dim fso As New FileSystemObject
If fso.FolderExists(workpro + "\compiler\" + compiler) = True Then
frmMain.ActiveForm.errfiles.Clear
frmMain.ActiveForm.errors.Clear
frmMain.ActiveForm.errf.Clear

Dim errors As Boolean

pro_Click
Load dorun
dorun.lstrun.Clear
dorun.lstdeb.Clear
temp = pro.ListCount - 1

For temp = 0 To pro.ListCount - 1
BuildFile filename(temp), filedata(temp)
Next

If Split(fol, "\")(UBound(Split(fol, "\"))) = vbNullString Then
fol = fol
Else
fol = fol + "\"
End If

For temp = 0 To UBound(filename)
'main process Call
For tex = 0 To 100
If UCase(fso.GetExtensionName(filename(temp))) = UCase(asso(tex, 0)) Then
For tex1 = 0 To 100
If compilers(tex1, 0) = asso(tex, 1) Then
compiler = compilers(tex1, 1)
End If
Next
consEXE CStr(compiler), filename(temp)
End If
Next
Next

For temp = 0 To UBound(filename)
errors = False
'main process Call
For tex = 0 To 100
If UCase(fso.GetExtensionName(filename(temp))) = UCase(asso(tex, 0)) Then
For tex1 = 0 To 100
If compilers(tex1, 0) = asso(tex, 1) Then
compiler = compilers(tex1, 1)
End If
Next

'check if exe file exists
mem = fso.FileExists("c:\windows\temp\cpp\" + Split(filename(temp), ".")(0) + ".exe")

'initialises error reading section
If mem = False Then
errors = True
Else
fso.CopyFile "c:\windows\temp\cpp\" + Split((filename(temp)), ".")(0) + ".exe", fol + Split((filename(temp)), ".")(0) + ".exe", True
End If

End If
Next
Next


If errors = True Then MsgBox "Error Found: Ignored!"

Else
MsgBox "Error Found : Possible Cause : Corrupted or missing Compiler Files", vbCritical
End If
Unload dorun
End Function

Function sav()
pro_Click
End Function
Function doit()
pro_Click
End Function

Function getlastel()
getlastel = ttundo(100)
End Function
Function getlastel1()
getlastel1 = ttredo(100)
End Function

Function undoit()
ur.Enabled = False
For temp = 99 To 1 Step -1
ttundo(temp + 1) = ttundo(temp)
ttplc0(temp + 1) = ttplc0(temp)
Next

For temp = 2 To 100
ttredo(temp - 1) = ttredo(temp)
ttplc1(temp - 1) = ttplc1(temp)
Next
ttredo(100) = rtftext.TextRTF
ttplc1(100) = rtftext.SelStart

rtftext.TextRTF = ttundo(100)

rtftext.SelStart = ttplc0(100)
MsgBox (getlastel1)
last = rtftext.TextRTF
ur.Enabled = True


End Function

Function redoit()
ur.Enabled = False

For temp = 2 To 100
ttundo(temp - 1) = ttundo(temp)
ttplc0(temp - 1) = ttplc0(temp)
Next

ttundo(100) = rtftext.TextRTF
ttplc0(100) = rtftext.SelStart

rtftext.TextRTF = ttredo(100)
rtftext.SelStart = ttplc1(100)

For temp = 99 To 1 Step -1
ttredo(temp + 1) = ttredo(temp)
ttplc1(temp + 1) = ttplc1(temp)
Next

last = rtftext.TextRTF
ur.Enabled = True
End Function

Private Sub ur_Timer()
If ur.Enabled = True Then
If Not rtftext.TextRTF = last Then

For temp = 2 To 100
ttundo(temp - 1) = ttundo(temp)
ttplc0(temp - 1) = ttplc0(temp)
Next
ttundo(0) = vbNullString
ttplc0(0) = 1
ttundo(100) = rtftext.TextRTF
ttplc0(100) = rtftext.SelStart
last = rtftext.TextRTF

End If
End If
End Sub
