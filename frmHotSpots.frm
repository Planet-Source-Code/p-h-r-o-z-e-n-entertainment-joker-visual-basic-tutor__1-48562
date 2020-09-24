VERSION 5.00
Begin VB.Form frmHotSpots 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hot Spots"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmHotSpots.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done Learning About Hot Spots"
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
      Picture         =   "frmHotSpots.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Put The Mouse Over The Diffrent Hot Spots"
      Height          =   1815
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6135
      Begin VB.Label Label1 
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
         Left            =   2040
         TabIndex        =   4
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblWhat 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3480
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblHi 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "HI"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblHI1 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblMB 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmHotSpots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
'-Unload frmHotSpots
Unload Me
'-Make frmMain visible
frmMain.Visible = True
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-Change the caption of lblWhat to nothing
lblWhat.Caption = ""
'-Change the colors back to normal for the word HI
lblHi.BackColor = &HC0C0C0  '-Gray
lblHi.ForeColor = &H0&      '-Black
End Sub

Private Sub Label1_Click()
'-Setup messagebox
MsgBox ("Hot spots are the MouseMove function, MouseMove is the same as having the mouse over the object.")
End Sub

Private Sub lblHi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHi.BackColor = &H0&      '-Black
lblHi.ForeColor = &HFF00&   '-Green
End Sub

Private Sub lblHI1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHi.BackColor = &HC0C0C0  '-Gray
lblHi.ForeColor = &H0&      '-Black
End Sub

Private Sub lblMB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-Setup message box when the mouse is over it
MsgBox ("Stop going over my hot spot")
End Sub

Private Sub lblWhat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-Change the caption of lblWhat to \/
lblWhat.Caption = "What Are You Doing???"
End Sub
