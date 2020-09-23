VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000013&
   Caption         =   "Friendssoft System Tools 2.0"
   ClientHeight    =   7665
   ClientLeft      =   3360
   ClientTop       =   1710
   ClientWidth     =   5985
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   5985
   Begin VB.CommandButton Command20 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SQL Client Configuration"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Configuration Editor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Startup Key"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Windows Tour"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Windows undate Manager"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Utility Manager"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tele Net"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Task Manager"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registry Editor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Private Charactor Editor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Printers and Faxes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Performance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Log Off"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Key Board Properties"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Display Properties"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Blutooth Transfer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clip Board Viewer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Check Disk"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Control Panel"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1995
      Left            =   0
      Picture         =   "Form1.frx":0D2E
      Top             =   0
      Width           =   6000
   End
   Begin VB.Menu Filemnu 
      Caption         =   "&File"
      Begin VB.Menu FileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu ToolsMnu 
      Caption         =   "&Tools"
      Begin VB.Menu ToolsControlPanel 
         Caption         =   "&Control Panel"
         Shortcut        =   {F1}
      End
      Begin VB.Menu ToolsCheckDisk 
         Caption         =   "&Chek Disk"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Tools1 
         Caption         =   "-"
      End
      Begin VB.Menu ToolsClipBoardViewer 
         Caption         =   "&Clip Board Viewer"
         Shortcut        =   {F3}
      End
      Begin VB.Menu ToolsBlueToothTransfer 
         Caption         =   "&Blue Tooth Transfer"
         Shortcut        =   {F4}
      End
      Begin VB.Menu ToolsDisplayProperties 
         Caption         =   "&Display Properties"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Tools2 
         Caption         =   "-"
      End
      Begin VB.Menu ToolsKeyBoardProperties 
         Caption         =   "&Key Board Properties"
         Shortcut        =   {F6}
      End
      Begin VB.Menu ToolsPerformance 
         Caption         =   "&Performance"
         Shortcut        =   {F7}
      End
      Begin VB.Menu ToolsPrintersFaxes 
         Caption         =   "&Printers & Faxes"
         Shortcut        =   {F8}
      End
      Begin VB.Menu ToolsPrivateCharactorEditor 
         Caption         =   "&Private Charactor Editor"
         Shortcut        =   {F9}
      End
      Begin VB.Menu ToolsRegistryEditor 
         Caption         =   "&Registry Editor"
         Shortcut        =   {F11}
      End
      Begin VB.Menu Tools3 
         Caption         =   "-"
      End
      Begin VB.Menu ToolsTaskManager 
         Caption         =   "&Task Manager"
         Shortcut        =   {F12}
      End
      Begin VB.Menu ToolsTeleNet 
         Caption         =   "&Tele Net"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu ToolsUtilityManager 
         Caption         =   "&Utility Manager"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu ToolsWindowsUpdateManager 
         Caption         =   "&Windows Update Manager"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu ToolsWindowsXPTour 
         Caption         =   "&WindowsXP Tour"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu ToolsStartupKey 
         Caption         =   "&Startup Key"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu Tools4 
         Caption         =   "-"
      End
      Begin VB.Menu ToolsConfigurationEditor 
         Caption         =   "&Configuration Editor"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu ToolsSQLClientConfiguration 
         Caption         =   "&SQL Client Configuration"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu Tools5 
         Caption         =   "-"
      End
      Begin VB.Menu ToolsLogOff 
         Caption         =   "&Log Off"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu Tools6 
         Caption         =   "-"
      End
      Begin VB.Menu ToolsOtherTools 
         Caption         =   "&Other Tools"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu HelpMnu 
      Caption         =   "&Help"
      Begin VB.Menu HelpStartRunCommands 
         Caption         =   "&Start-Run Commands"
         Shortcut        =   ^R
      End
      Begin VB.Menu HelpAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu Help 
         Caption         =   "-"
      End
      Begin VB.Menu HelpDownlodableSite 
         Caption         =   "&Downlodable Site"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "Control Panel", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command10_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "eudcedit", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command11_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "regedit", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command12_Click()

End Sub

Private Sub Command13_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "taskmgr", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command14_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "telnet", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command15_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "utilman", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command16_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "wupdmgr", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command17_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "tourstart", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command18_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "syskey", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command19_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "sysedit", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command2_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "chkdsk", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command20_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "cliconfg", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command21_Click()

End Sub

Private Sub Command3_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "clipbrd", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command4_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "C:\WINDOWS\system32\fsquirt", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command5_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "control.exe desktop", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command6_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "control.exe keyboard", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command7_Click()
Unload Me
On Error GoTo syserr
Shell "logoff", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command8_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "perfmon", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub Command9_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "control.exe printers", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub FileExit_Click()
End
End Sub
'**********************************************
'                                             *
'DESIGNED & PROGRAMED BY FAROOK FAHATH AHMED  *
'COMPANY : FRIENDSSOFT INC,                   *
'MUNAICHENAI.04,KINNIYA,SRI LANKA.            *
'T.P.: +94 77 1251413                         *
'E-MAIL : FRIENDSSOFTLK@GMAIL.COM             *
'FAHATH2008@GMAIL.COM                         *
'                                             *
'**********************************************
'Note : This computer programe protected by copyright law.If you
'like to redevelop this programe please informe us.
Private Sub Form_Load()

End Sub

Private Sub HelpAbout_Click()
frmAbout.Show
End Sub

Private Sub HelpDownlodableSite_Click()
On Error Resume Next
Shell "explorer http://kinniyaboys.achiever21in21.com"
End Sub

Private Sub HelpStartRunCommands_Click()
Form2.Show
End Sub

Private Sub ToolsBlueToothTransfer_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "C:\WINDOWS\system32\fsquirt", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsCheckDisk_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "chkdsk", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsClipBoardViewer_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "clipbrd", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsConfigurationEditor_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "sysedit", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsControlPanel_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "Control Panel", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsDisplayProperties_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "control.exe desktop", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsKeyBoardProperties_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "control.exe keyboard", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsLogOff_Click()
Unload Me
On Error GoTo syserr
Shell "logoff", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsOtherTools_Click()
Form1.WindowState = vbMinimized
Tools.Show
End Sub

Private Sub ToolsPerformance_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "perfmon", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsPrintersFaxes_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "control.exe printers", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsPrivateCharactorEditor_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "eudcedit", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsRegistryEditor_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "regedit", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsShutDown_Click()
Unload Me
On Error GoTo syserr
Shell "shutdown", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsSQLClientConfiguration_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "cliconfg", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsStartupKey_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "syskey", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsTaskManager_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "taskmgr", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsTeleNet_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "telnet", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsUtilityManager_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "utilman", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsWindowsUpdateManager_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "wupdmgr", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub

Private Sub ToolsWindowsXPTour_Click()
Form1.WindowState = vbMinimized
On Error GoTo syserr
Shell "tourstart", vbNormalFocus
Exit Sub
syserr: MsgBox Err.Description, vbCritical
End Sub
