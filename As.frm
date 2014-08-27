VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Assoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Association Editor"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8310
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "Add"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Discard Changes"
      Height          =   615
      Left            =   5520
      TabIndex        =   6
      Top             =   4680
      Width           =   2655
   End
   Begin VB.ListBox cmp 
      Appearance      =   0  'Flat
      Height          =   2370
      ItemData        =   "As.frx":0000
      Left            =   5520
      List            =   "As.frx":0002
      TabIndex        =   4
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save and Exit"
      Height          =   615
      Left            =   5520
      TabIndex        =   3
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set"
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Extension"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Currently Associated with"
         Object.Width           =   4498
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Associations"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Possible Associations"
      Height          =   255
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Assoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
lv.SelectedItem.SubItems(2) = cmp.List(cmp.ListIndex)
typ = GetSetting(App.Title, "compiler", "typ" + CStr(cmp.ItemData(cmp.ListIndex)), vbNullString)
ext = GetSetting(App.Title, "compiler", "handles" + CStr(cmp.ItemData(cmp.ListIndex)), vbNullString)

For tex = 0 To UBound(Split(ext, "/"))
If UCase(Split(ext, "/")(tex)) = UCase(lv.SelectedItem.Text) Then
lv.SelectedItem.SubItems(1) = Split(typ, "/")(tex)
End If
Next


End Sub

Private Sub Command2_Click()
temp = 1
Do While Not "x/x" = GetSetting(App.Title, "association", "ext" + CStr(temp), "x/x")
DeleteSetting App.Title, "association", "EXT" + CStr(temp)
DeleteSetting App.Title, "association", "des" + CStr(temp)
DeleteSetting App.Title, "association", "asso" + CStr(temp)
temp = temp + 1
Loop

For temp = 1 To lv.ListItems.Count
SaveSetting App.Title, "association", "EXT" + CStr(temp), lv.ListItems(temp).Text
SaveSetting App.Title, "association", "des" + CStr(temp), lv.ListItems(temp).SubItems(1)
SaveSetting App.Title, "association", "ASSO" + CStr(temp), lv.ListItems(temp).SubItems(2)
Next
frmMain.addasso
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
msg = MsgBox("Are you Sure you want to delete this association", vbYesNo, "Confirmation")
If msg = vbYes Then
lv.ListItems.remove (lv.SelectedItem.Index)
End If
lv_Click
End Sub

Private Sub Command5_Click()
ext = InputBox("Enter the extension", "New Association", vbNullString)

If Not ext = vbNullString Then
lv.ListItems.add , , ext
lv.ListItems(lv.ListItems.Count).SubItems(1) = "Not specified"
lv.ListItems(lv.ListItems.Count).SubItems(2) = "Nothing"
lv.ListItems(lv.ListItems.Count).Selected = True
End If
lv_Click
End Sub

Private Sub Form_Load()
temp = 1
Do While Not "x/x" = GetSetting(App.Title, "association", "EXT" + CStr(temp), "x/x")
lv.ListItems.add , , GetSetting(App.Title, "association", "ext" + CStr(temp), vbNullString)
lv.ListItems(temp).SubItems(1) = GetSetting(App.Title, "association", "DES" + CStr(temp), vbNullString)
lv.ListItems(temp).SubItems(2) = GetSetting(App.Title, "association", "ASSO" + CStr(temp), vbNullString)
temp = temp + 1
Loop

lv_Click
End Sub

Private Sub lv_Click()
On Error Resume Next
cmp.Clear
lsttyp.Clear
If lv.ListItems.Count = 0 Then Exit Sub
For temp = 0 To 100
texts = Split(compilers(temp, 3), "/")
For tex = 0 To UBound(texts)
'add compiler to possible compiler list for current association
If UCase(lv.SelectedItem.Text) = UCase(texts(tex)) Then
cmp.AddItem (compilers(temp, 0))
cmp.ItemData(cmp.ListCount - 1) = temp + 1
If UCase(lv.SelectedItem.SubItems(2)) = UCase(compilers(temp, 0)) Then
cmp.ListIndex = cmp.ListCount - 1
End If
End If
Next
Next

If cmp.ListIndex = -1 Then cmp.ListIndex = 0

End Sub
