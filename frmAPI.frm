VERSION 5.00
Begin VB.Form frmAPI 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "API"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAPI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done Learning About API"
      Height          =   975
      Left            =   0
      Picture         =   "frmAPI.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5640
      Width           =   5895
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H0000FF00&
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2655
      Left            =   3840
      TabIndex        =   37
      Top             =   3000
      Width           =   2055
      Begin VB.CommandButton Command6 
         Caption         =   "Mouse"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Keyboard"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Time"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Display"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H0000FF00&
      Caption         =   "System"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   3840
      TabIndex        =   36
      Top             =   2160
      Width           =   2055
      Begin VB.CommandButton Command2 
         Caption         =   "System"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H0000FF00&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   3840
      TabIndex        =   35
      Top             =   1320
      Width           =   2055
      Begin VB.CommandButton Command1 
         Caption         =   "Password"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H0000FF00&
      Caption         =   "Mouse"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   3840
      TabIndex        =   34
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton cmdDef 
         Caption         =   "Default"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdFlip 
         Caption         =   "Flip Buttons"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H0000FF00&
      Caption         =   "ShutDown"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1800
      TabIndex        =   33
      Top             =   4800
      Width           =   2055
      Begin VB.CommandButton cmdSD 
         Caption         =   "Shutdown"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H0000FF00&
      Caption         =   "Cursor"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   1800
      TabIndex        =   32
      Top             =   3480
      Width           =   2055
      Begin VB.CommandButton cdmCShow 
         Caption         =   "Show"
         Height          =   420
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdCHide 
         Caption         =   "Hide"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H0000FF00&
      Caption         =   "Desktop"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   1800
      TabIndex        =   31
      Top             =   2160
      Width           =   2055
      Begin VB.CommandButton cmdDHide 
         Caption         =   "Hide"
         Height          =   420
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdDShow 
         Caption         =   "Show"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H0000FF00&
      Caption         =   "Control Panel"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   1800
      TabIndex        =   29
      Top             =   840
      Width           =   2055
      Begin VB.CommandButton cmdHard 
         Caption         =   "HardWare"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdAR 
         Caption         =   "Add Remove"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H0000FF00&
      Caption         =   "Files"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1800
      TabIndex        =   28
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H0000FF00&
      Caption         =   "Explorer"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   0
      TabIndex        =   27
      Top             =   4800
      Width           =   1815
      Begin VB.CommandButton cmdEOpen 
         Caption         =   "Open"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0000FF00&
      Caption         =   "Minimize All"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   0
      TabIndex        =   26
      Top             =   3960
      Width           =   1815
      Begin VB.CommandButton cmdMin 
         Caption         =   "Minimize All"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0000FF00&
      Caption         =   "Taskbar"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   0
      TabIndex        =   25
      Top             =   2520
      Width           =   1815
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdHide 
         Caption         =   "Hide"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FF00&
      Caption         =   "CD-ROM"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   0
      TabIndex        =   24
      Top             =   1560
      Width           =   1815
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      Caption         =   "Ctrl Alt Del"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton cmdDisable 
         Caption         =   "Disable"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdEnable 
         Caption         =   "Enable"
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-NONE OF THIS WILL BE COMMENTED SORRY ABOUT THAT, Self-Explanitory
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
'-------------------------------------------------------------
Private Sub cdmCShow_Click()
Cursor_Show
End Sub

Private Sub cmdAR_Click()
Add_Remove
End Sub

Private Sub cmdCHide_Click()
Cursor_Hide
End Sub

Private Sub cmdDef_Click()
FlipMouseButtonsBack
End Sub

Private Sub cmdDHide_Click()
DesktopIconsHide
End Sub

Private Sub cmdDisable_Click()
ALT_CTRL_DEL_Disabled
End Sub

Private Sub cmdDone_Click()
'-Unload frmAPI
Unload Me
'-Make frmMain visible
frmMain.Visible = True
End Sub

Private Sub cmdDShow_Click()
DesktopIconsShow
End Sub

Private Sub cmdEnable_Click()
ALT_CTRL_DEL_Enabled
End Sub

Private Sub cmdEOpen_Click()
OpenExplore
End Sub

Private Sub cmdFind_Click()
FindFiles
End Sub

Private Sub cmdFlip_Click()
FlipMouseButtons
End Sub

Private Sub cmdHard_Click()
Add_HardWare
End Sub

Private Sub cmdHide_Click()
TaskBarHide
End Sub

Private Sub cmdMin_Click()
MinimizeAll
End Sub

Private Sub cmdOpen_Click()
OpenCDROM
End Sub

Private Sub cmdSD_Click()
ShutDown_DIALOG
End Sub

Private Sub cmdShow_Click()
TaskBarShow
End Sub

Private Sub Command1_Click()
Password_Settings
End Sub

Private Sub Command2_Click()
System_Settings
End Sub

Private Sub Command3_Click()
Display_Settings
End Sub

Private Sub Command4_Click()
Time_Date_Settings
End Sub

Private Sub Command5_Click()
Keyboard_Settings
End Sub

Private Sub Command6_Click()
Mouse_Settings
End Sub

