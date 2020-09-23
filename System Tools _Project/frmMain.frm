VERSION 5.00
Begin VB.Form Tools 
   BackColor       =   &H80000013&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Windows Tools"
   ClientHeight    =   8355
   ClientLeft      =   5880
   ClientTop       =   1275
   ClientWidth     =   1575
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   1575
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H80000013&
      Caption         =   "NotePad"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MouseIcon       =   "frmMain.frx":0A02
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H80000013&
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MouseIcon       =   "frmMain.frx":170E
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":1A18
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H80000013&
      Caption         =   "Internet "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmMain.frx":241A
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":2724
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H80000013&
      Caption         =   "Reg.Edit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MouseIcon       =   "frmMain.frx":3126
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":3430
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H80000013&
      Caption         =   "Win Explorer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MouseIcon       =   "frmMain.frx":44B2
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":47BC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H80000013&
      Caption         =   "MS Dos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmMain.frx":51BE
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":54C8
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000013&
      Caption         =   "My Computer"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MouseIcon       =   "frmMain.frx":5ECA
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":61D4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000013&
      Caption         =   "Recycle Bin"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmMain.frx":7256
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":7560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000013&
      Caption         =   "Recent Files"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmMain.frx":7F62
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":826C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000013&
      Caption         =   "Network "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MouseIcon       =   "frmMain.frx":8C6E
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":8F78
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000013&
      Caption         =   "Desktop"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MouseIcon       =   "frmMain.frx":997A
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":9C84
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H80000013&
      Caption         =   "IE Cookies"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmMain.frx":AD06
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":B010
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H80000013&
      Caption         =   "Control Panel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MouseIcon       =   "frmMain.frx":BA12
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":BD1C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "Tools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwndOwner As Long, ByVal lpszPath As String, ByVal nFolder As Long, ByVal fCreate As Long) As Long

Private Const CSIDL_FONTS = &H14
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_COOKIES = &H21
Private Const CSIDL_HISTORY = &H22

Private Const NameSpace_MyComputer = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
Private Const NameSpace_RecycleBin = "::{645FF040-5081-101B-9F08-00AA002F954E}"
Private Const NameSpace_NetworkNeighborhood = "::{208D2C60-3AEA-1069-A2D7-08002B30309D}"
Private Const NameSpace_Dialup = "::{a4d92740-67cd-11cf-96f2-00aa00a11dd9}"
Private Const NameSpace_ControlPanel = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{21EC2020-3AEA-1069-A2DD-08002B30309D}"
Private Const NameSpace_Printers = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{2227A280-3AEA-1069-A2DE-08002B30309D}"
Private Const NameSpace_ScheduledTasks = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{D6277990-4C6A-11CF-8D87-00AA0060F5BF}"

Private Const MAX_PATH = 260
Private Sub OpenExplorerWindow(FolderName As String)
    Shell "explorer " & FolderName, vbNormalFocus
End Sub

Private Function TrimNull(Str1 As String) As String
    Dim Loc         As Integer
    
    Loc = InStr(Str1, Chr$(0))
    If Loc <> 0 Then
        TrimNull = Mid$(Str1, 1, Loc - 1)
    Else
        TrimNull = Str1
    End If
End Function

Private Function GetSpecialFolder(Folder As Long) As String
    Dim FolderPath          As String * MAX_PATH
    SHGetSpecialFolderPath 0, FolderPath, Folder, 0
    GetSpecialFolder = TrimNull(FolderPath)
End Function


Private Sub Command1_Click()
Unload Me
Form1.WindowState = vbNormal
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.ToolTipText = "Close"
End Sub

Private Sub Command10_Click()
OpenExplorerWindow NameSpace_ControlPanel
End Sub

Private Sub Command10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command10.ToolTipText = "Click for open the Control panel"
End Sub

Private Sub Command11_Click()
On Error Resume Next
Shell ("cmd")
End Sub

Private Sub Command11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command11.ToolTipText = "Command Prompt"
End Sub

Private Sub Command12_Click()
On Error Resume Next
Shell ("explorer")
End Sub

Private Sub Command12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command12.ToolTipText = "Windows Explorer"
End Sub

Private Sub Command13_Click()
On Error Resume Next
Shell ("regedit")
'MsgBox "The Application cannot be Executed"
MsgBox "Windows cannot access the specified drive,path or file."
End Sub

Private Sub Command13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command13.ToolTipText = "Registry Editor"
End Sub

Private Sub Command14_Click()
On Error Resume Next
Shell ("C:\Program Files\Internet Explorer\IEXPLORE.EXE")
End Sub

Private Sub Command14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command14.ToolTipText = "Internet Explorer"
End Sub

Private Sub Command15_Click()
On Error Resume Next
Shell ("calc")
End Sub

Private Sub Command15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command15.ToolTipText = "Open calculator"
End Sub

Private Sub Command16_Click()
On Error Resume Next
Shell ("notepad")
End Sub

Private Sub Command16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command16.ToolTipText = "MS NotePad"
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
End Sub

Private Sub Command3_Click()
OpenExplorerWindow NameSpace_MyComputer
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.ToolTipText = "Click for open My computer"
End Sub

Private Sub Command4_Click()
OpenExplorerWindow NameSpace_RecycleBin
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.ToolTipText = "Click for open the Recycle Bin"
End Sub

Private Sub Command5_Click()
    OpenExplorerWindow GetSpecialFolder(CSIDL_RECENT)
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.ToolTipText = "Recent Files Folder"
End Sub

Private Sub Command6_Click()
OpenExplorerWindow NameSpace_NetworkNeighborhood
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command6.ToolTipText = "My Network Places"
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Command8_Click()
OpenExplorerWindow GetSpecialFolder(CSIDL_DESKTOP)
End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command8.ToolTipText = "Click for open Desktop"
End Sub

Private Sub Command9_Click()
OpenExplorerWindow GetSpecialFolder(CSIDL_COOKIES)
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuHelp_Click()

End Sub

Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command9.ToolTipText = "Click for open Cookies folder"
End Sub

Private Sub Form_Load()

End Sub
