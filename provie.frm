VERSION 5.00
Begin VB.Form provie 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Project Viewer"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command6 
      Caption         =   "Rename"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Move Down"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Move Up"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.ListBox Lst 
      Appearance      =   0  'Flat
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "provie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim butt As Boolean
Dim data() As String
Dim plc() As Integer

Private Sub Command1_Click()
Dim mem As String
Dim mem1 As String
If Not Lst.ListIndex = -1 Then
If Not Lst.ListIndex = 0 Then
memnum = Lst.ListIndex
mem = Lst.List(memnum)
mem1 = data(memnum)
mem2 = plc(memnum)
Lst.List(memnum) = Lst.List(memnum - 1)
data(memnum) = data(memnum - 1)
plc(memnum) = plc(memnum - 1)
Lst.List(Lst.ListIndex - 1) = mem
data(memnum - 1) = mem1
plc(memnum - 1) = mem2
Lst.ListIndex = memnum - 1
End If
End If
End Sub
Private Sub Command2_Click()
If Not Lst.ListIndex = -1 Then
If Not Lst.ListIndex = Lst.ListCount - 1 Then
memnum = Lst.ListIndex
mem = Lst.List(memnum)
mem1 = data(memnum)
mem2 = plc(memnum)
Lst.List(memnum) = Lst.List(memnum + 1)
data(memnum) = data(memnum + 1)
plc(memnum) = plc(memnum + 1)
Lst.List(memnum + 1) = mem
data(memnum + 1) = mem1
plc(memnum + 1) = mem2
Lst.ListIndex = memnum + 1
End If
End If
End Sub

Private Sub Command3_Click()
If Lst.ListIndex = -1 Or Lst.ListCount = 0 Then
Else
msg = MsgBox("Are you Sure?", vbYesNo + vbExclamation)
If msg = vbYes Then
mem = Lst.ListIndex
Lst.RemoveItem mem

'remove item from data
For temp = mem To Lst.ListCount - 1
data(temp) = data(temp + 1)
plc(temp) = plc(temp + 1)
Next
ReDim Preserve data(UBound(data) - 1)
ReDim Preserve plc(UBound(plc) - 1)

If mem > Lst.ListCount - 1 Then
Lst.ListIndex = mem - 1
Else
Lst.ListIndex = mem
End If
End If
End If
End Sub

Private Sub Command4_Click()
mem = frmMain.ActiveForm.pro.List(frmMain.ActiveForm.pro.ListIndex)
frmMain.ActiveForm.cleararr Lst.List(0), data(0), plc(0)
For temp = 1 To Lst.ListCount - 1
frmMain.ActiveForm.addarrnam (Lst.List(temp))
frmMain.ActiveForm.addarrdata (data(temp))
frmMain.ActiveForm.addarrplc (plc(temp))
Next
frmMain.ActiveForm.rebuild
Unload Me

For temp = 0 To frmMain.ActiveForm.pro.ListCount - 1
If UCase(frmMain.ActiveForm.pro.List(temp)) = UCase(mem) Then
frmMain.ActiveForm.pro.ListIndex = temp
End If
Next
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
def = Lst.List(Lst.ListIndex)
retry:
If Not Lst.ListIndex = -1 Then
inputs = InputBox("Enter new name for file (Extension Required) : ", "Rename", def)
If Not inputs = vbNullSring Then
If validate(CStr(inputs)) = True Then
If checkval(CStr(inputs)) = True Then
mem = Lst.ListIndex
Lst.List(Lst.ListIndex) = inputs
Lst.ListIndex = mem
Else
MsgBox "Invalid Character encountered. Please remove * / : ' < > ? \   spaces and try again", vbCritical
def = inputs
GoTo retry
End If
Else
MsgBox "Critical error : File already exists", , "Conflict"
def = inputs
GoTo retry
End If
End If
End If
End Sub

Private Sub Form_Load()
ReDim Preserve data(frmMain.ActiveForm.pro.ListCount - 1) As String
ReDim Preserve plc(frmMain.ActiveForm.pro.ListCount - 1) As Integer
If Not frmMain.ActiveForm Is Nothing Then
Dim temp As Integer
For temp = 0 To frmMain.ActiveForm.getubounda
Lst.AddItem frmMain.ActiveForm.getarrnam(temp)
data(temp) = frmMain.ActiveForm.getarrdata(temp)
plc(temp) = frmMain.ActiveForm.getarrplc(temp)
Next
On Error Resume Next
Lst.ListIndex = 0
End If
End Sub


