Attribute VB_Name = "mains"
Public ftext As String
Public reptext As String
Public workpro As String
Public widthex As Integer
Public comp As String
Public debugging As Boolean
Public errbox As Integer
Public debmenu As Boolean
Public cfolder As String
Public appl As TextBox
Public status  As Integer
Public runrun As Boolean
Public autospc As Boolean
Public autocol As Boolean
Public autoind As Boolean
Public disabled As Boolean
Public filess As String
Public ShowAtStart As Integer
Public compilers(100, 4) As String
Public asso(100, 2) As String
Public projectex As Boolean

'VARIABLES for keywords,operators and indendation
Public keywords() As String
Public operators() As String
Public indentation(1) As String

Sub Main()
autoind = True
runrun = False
debugging = False
widthex = 3000
errbox = 3000
Load frmMain
Load frmSplash
frmSplash.Show
End Sub
