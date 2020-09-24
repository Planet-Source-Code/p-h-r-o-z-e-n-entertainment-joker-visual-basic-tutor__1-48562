VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTimer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Timer"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "frmTimer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done Learning About Timers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "frmTimer.frx":0BC2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   7455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Timer"
      Height          =   2775
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7455
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Left            =   120
         Top             =   1680
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "&Set Up Timer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         Picture         =   "frmTimer.frx":148C
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "&Run Timer"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1920
         Picture         =   "frmTimer.frx":2356
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top Timer"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3720
         Picture         =   "frmTimer.frx":2660
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Res&et Timer"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   5520
         Picture         =   "frmTimer.frx":296A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin MSComctlLib.ProgressBar pbLoad 
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click Me!"
         DataSource      =   "&H00000000&"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   2280
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
'-Unload frmTimer
Unload Me
'-make frmMain Visible
frmMain.Visible = True
End Sub

Private Sub cmdReset_Click()
'-Reset the pbLoad bar to 0
pbLoad.Value = 0
End Sub

Private Sub cmdRun_Click()
'-Enable the timer
Timer1.Enabled = True
End Sub

Private Sub cmdSet_Click()
'-You need to set up the timer before you can continue
'---------\/ if you set the timer you can continue
cmdRun.Enabled = True
cmdStop.Enabled = True
cmdReset.Enabled = True
'---------/\--------------------------------------
'-Declare varibles
Dim SetT As String
SetT = InputBox("Enter a number between 1-1000", "Set-up Timer", "100")
'-Setup how fast the timer is
Timer1.Interval = Val(SetT)
End Sub

Private Sub cmdStop_Click()
'-Disable the timer
Timer1.Enabled = False
End Sub

Private Sub Label1_Click()
'-Setup message box
MsgBox ("You need to setup the timer before you can go on. When setting up the timer you are setting how long of an interval there is before it does the code again.")
End Sub

Private Sub Timer1_Timer()
'-the timer will continue as long as it
'-does not reach the max value
If pbLoad.Value < pbLoad.Max Then
    pbLoad = Val(pbLoad.Value) + 10
End If
'-If it reaches the max value it disables the timer
If pbLoad >= pbLoad.Max Then
    Timer1.Enabled = False
End If
End Sub
