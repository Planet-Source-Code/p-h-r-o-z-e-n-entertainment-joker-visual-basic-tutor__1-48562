VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStartup 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn Visual Basic"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStartup.frx":030A
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLoad 
      Interval        =   1
      Left            =   0
      Top             =   2400
   End
   Begin MSComctlLib.ProgressBar PbarStartup 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblVote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Learn Visual Basic v1.5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   645
      TabIndex        =   2
      Top             =   1560
      Width           =   3285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "|P|h|r|o|z|e|n| Entertainment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'-If it messes up
On Error GoTo errhandle
'-Declare Varibles
Dim directory$
Dim MyString$
'-Loads "Vote.ini" from my prog's path and the Saves Folder
directory$ = App.Path & "\Saves\Vote.ini"
'-Opens the files and reads it
Open directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
    Wend
    '-Puts it in the text box
    lblVote.Caption = MyString$
    '-Closes the file
    Close #1
'-If it does mess up it will not do anything
errhandle:
End Sub

Private Sub tmrLoad_Timer()
If PbarStartup.Value < 100 Then
'-The Progressbar's value will have 2 added every 1
'-millisecond when the value of pbarstartup is less than 100
    PbarStartup = Val(PbarStartup.Value) + 5
End If

'-If the value of pbarstartup = 100 then the timer is disabled
If PbarStartup.Value = 100 Then
Timer1.Enabled = False
MsgBox ("Welcome To Learn Visual Basic v1.5 - Joker")
End If
    tmrLoad.Interval = 1000
    '-Unload frmStartup
    Unload Me
    '-Make frmMain Visible
    frmMain.Visible = True


End Sub
