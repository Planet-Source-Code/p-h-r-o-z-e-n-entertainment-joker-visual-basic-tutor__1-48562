VERSION 5.00
Begin VB.Form frmLoops 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loops"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "frmLoops.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done Learning About Loops"
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
      Left            =   0
      Picture         =   "frmLoops.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   6975
   End
   Begin VB.Frame Frame1 
      Caption         =   "For...Next..."
      Height          =   3375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6975
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   2640
         Picture         =   "frmLoops.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton cmdSetStep 
         Caption         =   "&Set Step"
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
         Left            =   120
         Picture         =   "frmLoops.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "&Run Loop"
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
         Left            =   120
         Picture         =   "frmLoops.frx":17A8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1440
         Width           =   2295
      End
      Begin VB.ListBox lstNum 
         Height          =   2985
         Left            =   4800
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtNum 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click Me!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Next X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "lstNum.Additem X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   720
         TabIndex        =   8
         Top             =   720
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "For X = 1 to"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmLoops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-Declare Global varibles
Dim S As String '-Used in sub cmdRun
Dim ST As String '-Used in sub cmdSetStep
Dim Num As Single

Private Sub cmdClear_Click()
'-Clear the listbox lstNum
lstNum.Clear
End Sub

Private Sub cmdDone_Click()
'-Unload frmLoops
Unload Me
'-Make frmMain visible
frmMain.Visible = True
End Sub

Private Sub cmdRun_Click()
'-Declare varibles
Dim X As Single
Dim ST As String
'-Gather info from textbox
Num = Val(txtNum.Text)
'-Setup loop
For X = 1 To Num
    lstNum.AddItem X '-put the numbers in the listbox
Next X
End Sub

Private Sub cmdSetStep_Click()
'-Declare varibles
Dim S As String
Dim X As Single
'-Setup input box to gather information
S = InputBox("Enter the Step value of the loop.", "Step Value", "1")
ST = Val(S)
'------------------------------------------------------------
'-Gather info from textbox
Num = Val(txtNum.Text)
'-Setup loop
For X = 1 To Num Step ST
    lstNum.AddItem X '-put the numbers in the listbox
Next X
End Sub

Private Sub Label4_Click()
'-Setup messagebox
MsgBox ("There are also For...Next... loops where you can say" & Chr(13) & "For X = 1 to 10 Step 2" & Chr(13) & "    lstNum.Additem X" & Chr(13) & "Next X" & Chr(13) & "Where the step part means it would count from 1 step 2 to 3 to 5 to 7... and so on.")
End Sub
