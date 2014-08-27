VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Source Manager"
   ClientHeight    =   5580
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   11400
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0442
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2760
      Top             =   945
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5310
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14446
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "1/27/2009"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "6:41 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2550
      Top             =   2055
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1365
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0884
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0996
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":103A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":138C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D82
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E94
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FA6
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20B8
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21CA
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22DC
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23EE
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2500
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2612
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2724
            Key             =   "Macro"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2836
            Key             =   "run"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2948
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A5A
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B6C
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3322
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3674
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   24
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "butundo"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "butredo"
            Object.ToolTipText     =   "Redo"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "butfind"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "butpro"
            Object.ToolTipText     =   "Project Explorer"
            ImageIndex      =   22
            Style           =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "butview"
            Object.ToolTipText     =   "Project Viewer"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "buttool"
            Object.ToolTipText     =   "Tools"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "butins"
            Object.ToolTipText     =   "Insert"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Macro"
            Object.ToolTipText     =   "Compile"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Run"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Debug"
            Object.ToolTipText     =   "Debug"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   11400
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin RichTextLib.RichTextBox rtftext 
         Height          =   375
         Left            =   12480
         TabIndex        =   6
         Top             =   -15
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393217
         HideSelection   =   0   'False
         TextRTF         =   $"frmMain.frx":3D18
      End
      Begin VB.TextBox pos 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   75
         Width           =   1455
      End
      Begin VB.ListBox lstwchr 
         Height          =   255
         ItemData        =   "frmMain.frx":3D9A
         Left            =   165
         List            =   "frmMain.frx":3DB9
         TabIndex        =   7
         Top             =   315
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   1485
         TabIndex        =   5
         Top             =   60
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox Text 
         Height          =   285
         Left            =   1035
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   45
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.ComboBox Fontname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmMain.frx":3DE1
         Left            =   6795
         List            =   "frmMain.frx":3E06
         TabIndex        =   3
         Text            =   "Comic Sans MS"
         Top             =   30
         Width           =   2655
      End
      Begin VB.ComboBox FontSize 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmMain.frx":3EA5
         Left            =   9510
         List            =   "frmMain.frx":3ED9
         TabIndex        =   2
         Text            =   "10"
         Top             =   30
         Width           =   1575
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu spr 
         Caption         =   "-"
      End
      Begin VB.Menu imps 
         Caption         =   "&Import..."
      End
      Begin VB.Menu exp 
         Caption         =   "&Export..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu undo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu redo 
         Caption         =   "&Redo"
         Shortcut        =   {F2}
      End
      Begin VB.Menu sps 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnueditbar1 
         Caption         =   "-"
      End
      Begin VB.Menu selall 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu mnufind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnufindnext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnureplace 
         Caption         =   "&Replace"
         Shortcut        =   ^K
      End
      Begin VB.Menu sprs 
         Caption         =   "-"
      End
      Begin VB.Menu gotol 
         Caption         =   "&Goto Line"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu sss 
         Caption         =   "-"
      End
      Begin VB.Menu proex 
         Caption         =   "Project &Explorer"
         Checked         =   -1  'True
         Shortcut        =   ^Q
      End
      Begin VB.Menu setting 
         Caption         =   "&Options"
      End
      Begin VB.Menu edi 
         Caption         =   "Editor Settings"
      End
   End
   Begin VB.Menu project 
      Caption         =   "&Project"
      Begin VB.Menu cus 
         Caption         =   "&Header file"
         Shortcut        =   ^H
      End
      Begin VB.Menu addf 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu addcu 
         Caption         =   "Add &Custom File"
      End
      Begin VB.Menu ext 
         Caption         =   "Add &External file"
         Shortcut        =   ^E
      End
      Begin VB.Menu ftrhfruj 
         Caption         =   "-"
      End
      Begin VB.Menu viewer 
         Caption         =   "Project Viewer"
      End
      Begin VB.Menu propro 
         Caption         =   "Project Properties"
      End
   End
   Begin VB.Menu build 
      Caption         =   "&Build"
      Begin VB.Menu cur 
         Caption         =   "Run &Current File"
         Shortcut        =   {F5}
      End
      Begin VB.Menu sel 
         Caption         =   "Run &Selection"
         Shortcut        =   ^D
      End
      Begin VB.Menu run 
         Caption         =   "&Run All"
         Shortcut        =   ^R
      End
      Begin VB.Menu ii 
         Caption         =   "-"
      End
      Begin VB.Menu curd 
         Caption         =   "Debug Current File"
         Shortcut        =   {F6}
      End
      Begin VB.Menu alld 
         Caption         =   "Debug All"
         Shortcut        =   ^B
      End
      Begin VB.Menu iii 
         Caption         =   "-"
      End
      Begin VB.Menu scm 
         Caption         =   "&Compile Menu"
         Shortcut        =   ^M
      End
      Begin VB.Menu se 
         Caption         =   "-"
      End
      Begin VB.Menu arg 
         Caption         =   "Program &Arguments"
      End
      Begin VB.Menu sdm 
         Caption         =   "Show Debug &Menu"
      End
   End
   Begin VB.Menu outp 
      Caption         =   "&Output"
      Begin VB.Menu outmod 
         Caption         =   "&Run in Output Mode"
         Shortcut        =   {F7}
      End
      Begin VB.Menu outman 
         Caption         =   "&Output Manager"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu tools 
      Caption         =   "&Tools"
      Begin VB.Menu spac 
         Caption         =   "Auto &Spaces"
         Checked         =   -1  'True
      End
      Begin VB.Menu autind 
         Caption         =   "Auto &Indent"
         Checked         =   -1  'True
      End
      Begin VB.Menu iiii 
         Caption         =   "-"
      End
      Begin VB.Menu aspac 
         Caption         =   "Add Spaces"
         Shortcut        =   ^G
      End
      Begin VB.Menu inda 
         Caption         =   "Indent Text"
         Shortcut        =   ^I
      End
      Begin VB.Menu ssss 
         Caption         =   "-"
      End
      Begin VB.Menu upper 
         Caption         =   "Make Selection Uppercase"
         Shortcut        =   ^U
      End
      Begin VB.Menu lower 
         Caption         =   "Make Selection Lowercase"
         Shortcut        =   ^L
      End
      Begin VB.Menu untab 
         Caption         =   "Untabify Selection"
         Shortcut        =   ^T
      End
      Begin VB.Menu sprsg 
         Caption         =   "-"
      End
      Begin VB.Menu disable 
         Caption         =   "&Disable all Services"
      End
   End
   Begin VB.Menu options 
      Caption         =   "&Advanced"
      Begin VB.Menu associator 
         Caption         =   "Association &Editor"
      End
      Begin VB.Menu seperator 
         Caption         =   "-"
      End
      Begin VB.Menu patch 
         Caption         =   "&Add Patch"
      End
      Begin VB.Menu add 
         Caption         =   "Add &Compiler"
      End
      Begin VB.Menu ftyhftyhf 
         Caption         =   "-"
      End
      Begin VB.Menu fs 
         Caption         =   "&File System Manager"
      End
      Begin VB.Menu ed 
         Caption         =   "Compiler &Viewer"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
   Begin VB.Menu pop 
      Caption         =   "pop"
      Visible         =   0   'False
      Begin VB.Menu opn 
         Caption         =   "Open Viewer"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Dim first As Boolean


Private Sub add_Click()
Load addc
addc.Show vbModal
End Sub

Private Sub addcf_Click()
If Not ActiveForm Is Nothing Then Me.ActiveForm.addcfile
End Sub

Private Sub addcu_Click()
ActiveForm.cusfile
End Sub

Private Sub addf_Click(Index As Integer)
ActiveForm.asfile (addf(Index).Tag)
End Sub

Private Sub alld_Click()
status = 3
debugging = False
frmMain.ActiveForm.Frame.Visible = False
frmMain.ActiveForm.remake
ActiveForm.deball (comp)
End Sub

Private Sub arg_Click()
If Not ActiveForm Is Nothing Then
mem = InputBox("Enter the arguments with EXE", "Argument", frmMain.ActiveForm.getarg)
If Not mem = vbNullString Then
ActiveForm.setarg (CStr(mem))
End If
End If
End Sub

Private Sub auto_Click()
auto.Checked = Not auto.Checked
End Sub

Private Sub aspac_Click()
If Not ActiveForm Is Nothing Then
storetemp
addspc
restoretext
End If
End Sub

Private Sub associator_Click()
Load Assoc
Assoc.Show vbModal
refreshmenu
End Sub

Private Sub autind_Click()
autoind = Not autoind
autind.Checked = Not autind.Checked
End Sub

Private Sub comps_Click(Index As Integer)
For temp = 1 To comps.Count - 1
comps(temp).Checked = False
Next
comps(Index).Checked = True
comp = comps(Index).Tag
End Sub

Private Sub cur_Click()
If Not ActiveForm Is Nothing Then
status = 1
debugging = False
frmMain.ActiveForm.Frame.Visible = False
frmMain.ActiveForm.remake
Me.ActiveForm.curRUN
End If
End Sub


Private Sub curd_Click()
If Not ActiveForm Is Nothing Then
status = 1
debugging = False
frmMain.ActiveForm.Frame.Visible = False
frmMain.ActiveForm.remake
ActiveForm.debgc
End If
End Sub

Private Sub cus_Click()
ActiveForm.asfile ("h")
End Sub

Private Sub disable_Click()
disable.Checked = Not disable.Checked
disabled = disable.Checked
End Sub

Private Sub ed_Click()
Load comed
comed.Show vbModal
refreshmenu
End Sub

Private Sub edi_Click()
Load eded
eded.Show vbModal
End Sub

Private Sub exp_Click()
MsgBox ("Not yet complete!")
End Sub

Private Sub ext_Click()
If Not ActiveForm Is Nothing Then Me.ActiveForm.addext
End Sub

Private Sub file_Click()
If Not ActiveForm Is Nothing Then Me.ActiveForm.addfile
End Sub

Private Sub fontname_click()
On Error Resume Next
ActiveForm.rtftext.Font.name = Fontname.List(Fontname.ListIndex)
appall
ActiveForm.rtftext.SetFocus
End Sub

Private Sub fontsize_Click()
On Error Resume Next
ActiveForm.rtftext.Font.Size = FontSize.List(FontSize.ListIndex)
appall
ActiveForm.rtftext.SetFocus
End Sub

Private Sub fs_Click()
Load fsm
fsm.Show vbModal
End Sub


Private Sub gotol_Click()
Load got
got.Show vbModal
End Sub

Private Sub head_Click()
If Not ActiveForm Is Nothing Then Me.ActiveForm.addhfile
End Sub

Private Sub bootup()
If Len(Command) > 0 Then
LoadNewDoc
ActiveForm.Caption = Command
ActiveForm.opens (Command)
debugging = False
frmMain.ActiveForm.Frame.Visible = False
frmMain.ActiveForm.remake
frmMain.ActiveForm.setsave (True)
Else
LoadNewDoc
End If
refreshmenu
sdm.Checked = debmenu
End Sub

Function fcomp()

SaveSetting App.Title, "compiler", "CMP" + CStr(1), "Turbo C++"
SaveSetting App.Title, "compiler", "DIR" + CStr(1), "Turboc"
SaveSetting App.Title, "compiler", "HANDLES" + CStr(1), "cpp"
SaveSetting App.Title, "compiler", "TYP" + CStr(1), "C++ File"
SaveSetting App.Title, "compiler", "RES" + CStr(1), "EXE"



SaveSetting App.Title, "association", "EXT" + CStr(1), "cpp"
SaveSetting App.Title, "association", "des" + CStr(1), "C++ File"
SaveSetting App.Title, "association", "ASSO" + CStr(1), "Turbo C++"

End Function

Function addcom()
For temp = 1 To 100
compilers(temp - 1, 0) = vbNullString
compilers(temp - 1, 1) = vbNullString
compilers(temp - 1, 2) = vbNullString
compilers(temp - 1, 3) = vbNullString
compilers(temp - 1, 4) = vbNullString
Next
tex = vbNullString
temp = 1
Do While Not "x/x" = GetSetting(App.Title, "compiler", "CMP" + CStr(temp), "x/x")
compilers(temp - 1, 0) = GetSetting(App.Title, "compiler", "CMP" + CStr(temp), vbNullString)
compilers(temp - 1, 1) = GetSetting(App.Title, "compiler", "DIR" + CStr(temp), vbNullString)
compilers(temp - 1, 2) = GetSetting(App.Title, "compiler", "RES" + CStr(temp), vbNullString)
compilers(temp - 1, 3) = GetSetting(App.Title, "compiler", "HANDLES" + CStr(temp), vbNullString)
compilers(temp - 1, 4) = GetSetting(App.Title, "compiler", "typ" + CStr(temp), vbNullString)
temp = temp + 1
Loop
End Function

Function addasso()
For temp = 1 To 100
asso(temp - 1, 0) = vbNullString
asso(temp - 1, 1) = vbNullString
asso(temp - 1, 2) = vbNullString
Next

tex = vbNullString
For temp = 1 To 100
If Not "x/x" = GetSetting(App.Title, "Association", "EXT" + CStr(temp), "x/x") Then
asso(temp - 1, 0) = GetSetting(App.Title, "Association", "EXT" + CStr(temp), vbNullString)
asso(temp - 1, 1) = GetSetting(App.Title, "Association", "ASSO" + CStr(temp), vbNullString)
asso(temp - 1, 2) = GetSetting(App.Title, "Association", "DES" + CStr(temp), vbNullString)
tex = temp
End If
Next

If Not ActiveForm Is Nothing Then
ActiveForm.pro_Click
End If
End Function

Private Sub LoadNewDoc()

    Static lDocumentCount As Long
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Caption = "Project " & lDocumentCount
    frmD.Show
    frmMain.ActiveForm.setsave (True)
End Sub

Private Sub imps_Click()
MsgBox ("Not yet complete!")
End Sub

Private Sub inda_Click()
If Not ActiveForm Is Nothing Then
storetemp
appind
restoretext
End If
End Sub

Private Sub Label1_Click()

End Sub


Private Sub MDIForm_Activate()
If first = False Then
bootup
register
first = True

On Error Resume Next
Load frmTip
frmTip.Show vbModal
End If
End Sub

Function refreshmenu()
'unloads all menu subitems present in addition list
For temp = 1 To addf.Count - 1
Unload addf(temp)
Next

'checks if any one of association is present or not
If vbNullString = GetSetting(App.Title, "association", "ext1", vbNullString) Then
addf(0).Visible = False
Else
addf(0).Visible = True
End If

'adds items to adding list retrived for registry
temp = 1
Do While Not vbNullString = GetSetting(App.Title, "association", "des" + CStr(temp), vbNullString)
Load addf(temp)
addf(temp).Caption = GetSetting(App.Title, "association", "des" + CStr(temp), vbNullString) + " (*." + GetSetting(App.Title, "association", "ext" + CStr(temp), vbNullString) + ")"
addf(temp).Tag = GetSetting(App.Title, "association", "ext" + CStr(temp), vbNullString)
temp = temp + 1
Loop

'adds ending seperator to list
If Not vbNullString = GetSetting(App.Title, "association", "ext1", vbNullString) Then
Load addf(temp)
addf(temp).Caption = "-"
End If
End Function

Function register()
Dim fso As New FileSystemObject
Dim txt As TextStream
Set txt = fso.OpenTextFile("c:\windows\temp\reg.reg", ForWriting, True)
txt.WriteLine ("Windows Registry Editor Version 5.00")
txt.WriteLine ("")
txt.WriteLine ("[HKEY_CLASSES_ROOT\.cpx]")
txt.WriteLine ("@=" + frmSplash.colon.Text + "Consolidated C++ File" + frmSplash.colon.Text)
txt.WriteLine ("")
txt.WriteLine ("[HKEY_CLASSES_ROOT\.cpx\DefaultIcon]")

work = GetSetting(App.Title, "Main", "DIR", CurDir)
For temp = 1 To Len(work)
If Mid(work, 1, 1) = "\" Then
str1 = Left(work, temp - 1)
str2 = Right(work, Len(work) - temp)
work = str1 + "\\" + str2
temp = temp + 1
End If
Next

txt.WriteLine ("")
txt.WriteLine ("")
txt.WriteLine ("")
End Function

Private Sub MDIForm_Load()
If GetSetting(App.Title, "isfirst", "first", True) Then
fcomp
SaveSetting App.Title, "isfirst", "first", False
End If

pos.BackColor = frmDocument.BackColor
Me.Caption = App.Title
workpro = App.path
debmenu = GetSetting(App.Title, "options", "debm", True)
cfolder = GetSetting(App.Title, "options", "comfol", "C:\" + App.Title)

addcom
addasso
autocol = GetSetting(App.Title, "special", "autocol", True)
autospc = GetSetting(App.Title, "special", "autospc", True)
autoind = GetSetting(App.Title, "special", "autoind", True)
ShowAtStart = GetSetting(App.Title, "Options", "Show Tips at Startup", 1)
disabled = GetSetting(App.Title, "special", "disabled", False)
projectex = GetSetting(App.Title, "special", "proex", True)

disable.Checked = disabled
spac.Checked = autospc
autind.Checked = autoind

Dim fso As New FileSystemObject
If fso.FolderExists(workpro + "\compiler") = False Then
fso.CreateFolder (workpro + "\compiler")
End If

If fso.FolderExists("C:\windows\temp\Cpp") = False Then
fso.CreateFolder ("C:\windows\temp\Cpp")
End If

    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 10740)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    Me.WindowState = GetSetting(App.Title, "Settings", "window", 0)
    
End Sub

Private Sub MDIForm_Unload(cancel As Integer)
    
SaveSetting App.Title, "options", "debm", CBool(debmenu)
SaveSetting App.Title, "options", "comfol", cfolder

SaveSetting App.Title, "special", "autocol", autocol
SaveSetting App.Title, "special", "autospc", autospc
SaveSetting App.Title, "special", "autoind", autoind
SaveSetting App.Title, "options", "Show Tips at Startup", ShowAtStart
SaveSetting App.Title, "special", "disabled", disabled
SaveSetting App.Title, "special", "proex", projectex

    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.top
        If Not Me.Width < 10740 Then
                SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        End If
       If Not Me.Height < 6500 Then
                        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
        End If
        SaveSetting App.Title, "Settings", "window", Me.WindowState
    End If
    
unloads
End Sub
Function unloads()
Unload addc
Unload addpat
Unload comed
Unload compiles
Unload debg
Unload dorun
Unload Editor
Unload find
Unload folbro
Unload folbro1
Unload frmAbout
Unload frmOptions
Unload frmreplace
Unload frmSplash
Unload fsm
Unload provie
Unload stp
Unload Wait
Unload eded
Unload proper
End Function

Private Sub mnuFile_Click()
If ActiveForm Is Nothing Then
mnuFileSave.Enabled = False
mnuFileSaveAs.Enabled = False
mnuFileClose.Enabled = False
imps.Enabled = False
exp.Enabled = False
mnuFilePrint.Enabled = False
Else
mnuFileSave.Enabled = True
mnuFileSaveAs.Enabled = True
mnuFileClose.Enabled = True
imps.Enabled = True
exp.Enabled = True
mnuFilePrint.Enabled = True
End If
End Sub

Private Sub mnufind_Click()
If Not ActiveForm Is Nothing Then find.Show vbModal
End Sub

Private Sub mnufindnext_Click()
If Not ActiveForm Is Nothing Then frmMain.ActiveForm.rtftext.find ftext, frmMain.ActiveForm.rtftext.SelStart + frmMain.ActiveForm.rtftext.SelLength
End Sub

Private Sub mnureplace_Click()
If Not ActiveForm Is Nothing Then
Load frmreplace
frmreplace.Show vbModal
End If
End Sub


Private Sub opn_Click()
ActiveForm.doit
End Sub

Private Sub outmod_Click()
If Not ActiveForm Is Nothing Then
status = 1
debugging = False
frmMain.ActiveForm.Frame.Visible = False
frmMain.ActiveForm.remake
Me.ActiveForm.OutRUN
End If


End Sub

Private Sub patch_Click()
Load addpat
addpat.Show vbModal

End Sub

Private Sub proex_Click()
proex.Checked = Not proex.Checked
If tbToolBar.Buttons(12).Value = tbrUnpressed Then
tbToolBar.Buttons(12).Value = tbrPressed
Else
tbToolBar.Buttons(12).Value = tbrUnpressed
End If

projectex = proex.Checked

If Not ActiveForm Is Nothing Then ActiveForm.showex

End Sub


Private Sub propro_Click()
If Not ActiveForm Is Nothing Then
Load proper
proper.Show vbModal
End If
End Sub

Private Sub rem_Click()
ActiveForm.remove
End Sub


Private Sub run_Click()
If Not ActiveForm Is Nothing Then
status = 3
debugging = False
frmMain.ActiveForm.Frame.Visible = False
frmMain.ActiveForm.remake
ActiveForm.runall
End If
End Sub

Private Sub scm_Click()
If Not ActiveForm Is Nothing Then
Load compiles
compiles.Show vbModal
End If
End Sub

Private Sub sdm_Click()
sdm.Checked = Not sdm.Checked
debmenu = sdm.Checked
End Sub

Private Sub sel_Click()
If Not ActiveForm Is Nothing Then
status = 2
debugging = False
frmMain.ActiveForm.Frame.Visible = False
frmMain.ActiveForm.remake
ActiveForm.selrun
End If
End Sub

Private Sub selall_Click()
If Not ActiveForm Is Nothing Then
ActiveForm.rtftext.SelStart = 0
ActiveForm.rtftext.SelLength = Len(ActiveForm.rtftext.Text)
End If
End Sub

Private Sub setting_Click()
Load frmOptions
frmOptions.Show vbModal
End Sub

Private Sub spac_Click()
spac.Checked = Not spac.Checked
autospc = spac.Checked
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
       Case "butundo"
           
        Case "butredo"
            
        Case "butfind"
            mnureplace_Click
        Case "butpro"
            proex.Checked = Not proex.Checked
            projectex = proex.Checked
            If Not ActiveForm Is Nothing Then ActiveForm.showex
        Case "butview"
            viewer_Click
            Case "buttool"
            
            Case "butinsert"
            
       Case "Macro"
            Load compiles
            compiles.Show vbModal
        Case "Forward"
            cur_Click
        Case "Debug"
            curd_Click
        Case "Align Left"
            ActiveForm.rtftext.SelAlignment = rtfLeft
        Case "Center"
            ActiveForm.rtftext.SelAlignment = rtfCenter
        Case "Align Right"
            ActiveForm.rtftext.SelAlignment = rtfRight
    End Select
End Sub


Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.rtftext.SelRTF = Clipboard.GetText

End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtftext.SelRTF

End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtftext.SelRTF
    ActiveForm.rtftext.SelText = vbNullString

End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Print"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If ActiveForm.rtftext.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If err <> MSComDlg.cdlCancel Then
            ActiveForm.rtftext.SelPrint .hDC
        End If
    End With

End Sub

Private Sub mnuFileSaveAs_Click()
On Error GoTo endf
    Dim sFile As String
    Dim fso As New FileSystemObject
    

    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Save As"
        .CancelError = True
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "C++ Projects (*.cpx)|*.cpx"
        .ShowSave
        If Len(.filename) = 0 Then
            Exit Sub
        End If
        sFile = .filename
    End With
    If fso.FileExists(sFile) = True Then
        msg = MsgBox("Do you want to replace the existing file?", vbYesNoCancel, "Overwrite")
        If msg = vbYes Then
            ActiveForm.Caption = sFile
            ActiveForm.saves (sFile)
        ElseIf msg = vbNo Then
            mnuFileSave_Click
        End If
    Else
        ActiveForm.Caption = sFile
        ActiveForm.saves (sFile)
    End If
ActiveForm.setsave (True)
endf:
End Sub

Private Sub mnuFileSave_Click()
On Error GoTo endf
If Not ActiveForm Is Nothing Then
    Dim sFile As String
    If Left$(ActiveForm.Caption, 7) = "Project" Then
        With dlgCommonDialog
            .DialogTitle = "Save"
            .CancelError = True
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = "C++ Projects(*.cpx)|*.cpx"
            .ShowSave
            If Len(.filename) = 0 Then
                Exit Sub
            End If
            sFile = .filename
        End With
        
    Dim fso As New FileSystemObject
    If fso.FileExists(sFile) = True Then
        msg = MsgBox("Do you want to overwrite file?", vbYesNoCancel, "Overwrite")
        If msg = vbYes Then
            ActiveForm.Caption = sFile
            ActiveForm.saves (sFile)
        ElseIf msg = vbNo Then
            mnuFileSave_Click
        End If
    Else
        ActiveForm.Caption = sFile
        ActiveForm.saves (sFile)
    End If
               
               
               
    Else
      sFile = ActiveForm.Caption
            ActiveForm.saves (sFile)
           End If
    ActiveForm.setsave (True)
    End If
endf:
End Sub

Private Sub mnuFileClose_Click()
If Not ActiveForm Is Nothing Then Unload ActiveForm
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo endf
    Dim sFile As String

    
    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = True
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "C++ Projects (*.cpx)|*.cpx"
        .ShowOpen
        If Len(.filename) = 0 Then
                Exit Sub
        End If
        sFile = .filename
    
    If ActiveForm Is Nothing Then
        LoadNewDoc
        ActiveForm.Caption = sFile
        ActiveForm.opens (sFile)
    Else
        msg = MsgBox("Do you want to save " + ActiveForm.Caption + "?", vbYesNoCancel, "Open")
        If msg = vbNo Then
            ActiveForm.Caption = sFile
            ActiveForm.opens (sFile)
        ElseIf msg = vbYes Then
            mnuFileSave_Click
            ActiveForm.Caption = sFile
            ActiveForm.opens (sFile)
        End If
    End If
    
    End With
    debugging = False
    frmMain.ActiveForm.Frame.Visible = False
frmMain.ActiveForm.remake
frmMain.ActiveForm.setsave (True)
endf:
End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
    debugging = False
    frmMain.ActiveForm.Frame.Visible = False
frmMain.ActiveForm.remake
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
tools_click
mnuedit_click
If ActiveForm Is Nothing Then
mnuedit.Enabled = False
project.Enabled = False
build.Enabled = False
tools.Enabled = False
undo.Enabled = False
redo.Enabled = False
tbToolBar.Buttons(7).Enabled = False
tbToolBar.Buttons(8).Enabled = False
Else
mnuedit.Enabled = True
project.Enabled = True
build.Enabled = True
tools.Enabled = True

If frmMain.ActiveForm.getlastel = vbNullString Then
undo.Enabled = False
tbToolBar.Buttons(7).Enabled = False

Else
undo.Enabled = True
tbToolBar.Buttons(7).Enabled = True

End If
Text7.Text = frmMain.ActiveForm.getlastel1
If frmMain.ActiveForm.getlastel1 = vbNullString Then
redo.Enabled = False
tbToolBar.Buttons(8).Enabled = False
Else
tbToolBar.Buttons(8).Enabled = True
redo.Enabled = True
End If

End If

End Sub

Private Sub tools_click()
If disabled = 0 Then
disable.Checked = False
Else
disable.Checked = True
End If

If autospc = 0 Then
spac.Checked = False
Else
spac.Checked = True
End If
On Error Resume Next
mem = frmMain.ActiveForm.rtftext.SelLength
lower.Enabled = mem
upper.Enabled = mem
untab.Enabled = mem

End Sub


Private Sub upper_Click()
'retrive info from textbar
start = frmMain.ActiveForm.rtftext.SelStart
leng = frmMain.ActiveForm.rtftext.SelLength
res = UCase(frmMain.ActiveForm.rtftext.SelText)

'restore information to text bar
frmMain.ActiveForm.rtftext.SelText = res
frmMain.ActiveForm.rtftext.SelStart = start
frmMain.ActiveForm.rtftext.SelLength = leng
End Sub

Private Sub lower_Click()
'retrive info from textbar
start = frmMain.ActiveForm.rtftext.SelStart
leng = frmMain.ActiveForm.rtftext.SelLength
spl = LCase(frmMain.ActiveForm.rtftext.SelText)

'restore information to text bar
frmMain.ActiveForm.rtftext.SelText = spl
frmMain.ActiveForm.rtftext.SelStart = start
frmMain.ActiveForm.rtftext.SelLength = leng
End Sub

Private Sub untab_Click()
'retrive info from textbar
start = frmMain.ActiveForm.rtftext.SelStart
leng = frmMain.ActiveForm.rtftext.SelLength
spl = Split(frmMain.ActiveForm.rtftext.SelText, vbNewLine)

'process and convert selected into lcase

For temp = 0 To UBound(spl)

Tabs = 0
For mem = 1 To Len(spl(temp))
If Mid(spl(temp), mem, 1) = vbTab Then
Tabs = mem
End If
Next

If temp = 0 Then
res = Right(spl(temp), Len(spl(temp)) - Tabs)
Else
res = res + vbNewLine + Right(spl(temp), Len(spl(temp)) - Tabs)
End If

Next

'restore information to text bar
frmMain.ActiveForm.rtftext.SelText = res
frmMain.ActiveForm.rtftext.SelStart = start
frmMain.ActiveForm.rtftext.SelLength = leng
appall
End Sub
Private Sub viewer_Click()
If Not ActiveForm Is Nothing Then
Load provie
provie.Show vbModal
End If
End Sub

Sub mnuedit_click()

End Sub
