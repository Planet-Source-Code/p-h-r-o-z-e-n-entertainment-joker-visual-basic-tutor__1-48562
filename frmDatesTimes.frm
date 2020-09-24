VERSION 5.00
Begin VB.Form frmDatesTimes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dates and Times"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done learning About Dates and Times"
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
      Picture         =   "frmDatesTimes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dates and Times"
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   120
         Top             =   1680
      End
      Begin VB.Label lblD 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label lblTD 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         TabIndex        =   1
         Top             =   1200
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmDatesTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
'-Unload frmDatesTimes
Unload Me
'-Make frmMain visible
frmMain.Visible = True
End Sub

Private Sub Form_Load()
Call Timer1_Timer '-Goes to Timer1_Timer Sub
End Sub

Private Sub lblTD_Click()

End Sub

Private Sub Timer1_Timer()
lblD = Date '-Date = date
lblTD = Now '-Now = Date and time
End Sub
