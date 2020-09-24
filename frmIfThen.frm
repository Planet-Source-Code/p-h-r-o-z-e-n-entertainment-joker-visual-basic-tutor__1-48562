VERSION 5.00
Begin VB.Form frmIfThen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "If Then Statments"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmIfThen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done Learning About If Then Statments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Picture         =   "frmIfThen.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   5295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Basic If Then Statment"
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdTest 
         Caption         =   "&Test Statment"
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
         Picture         =   "frmIfThen.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2280
         Width           =   5055
      End
      Begin VB.TextBox txt2 
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
         Height          =   330
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txt1 
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
         Height          =   345
         Left            =   480
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "End If"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "If it is False the MsgBox (""False"")"
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
         Left            =   480
         TabIndex        =   8
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ElseIf"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "If is it True then MsgBox (""True"")"
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
         Left            =   480
         TabIndex        =   6
         Top             =   720
         Width           =   3420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "then"
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
         Left            =   4680
         TabIndex        =   5
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "is greater than"
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
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "If "
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   210
      End
   End
End
Attribute VB_Name = "frmIfThen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
'-Unload frmIfThen
Unload Me
'-Make frmMain visible
frmMain.Visible = True
End Sub

Private Sub cmdTest_Click()
'-Declare Varibles
Dim Num1 As Integer
Dim Num2 As Integer
'-Get numbers to be tested
Num1 = Val(txt1.Text)
Num2 = Val(txt2.Text)
'-Setup If Then Statment
If Num1 > Num2 Then
    MsgBox ("True") '-If Num1 is greater than num2
ElseIf Num1 < Num2 Then
    MsgBox ("False") '-'-If Num1 is less than num2
ElseIf Num1 = Num2 Then
    MsgBox ("Null") '-If Num1 is equal than num2
End If
End Sub
