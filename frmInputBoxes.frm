VERSION 5.00
Begin VB.Form frmInputBoxes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Boxes"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "frmInputBoxes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "D&one Learning About Input Boxes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      Picture         =   "frmInputBoxes.frx":0CFA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3840
      Width           =   5775
   End
   Begin VB.Frame Frame4 
      Caption         =   "Genarate InputBox"
      Height          =   1935
      Left            =   2760
      TabIndex        =   9
      Top             =   1920
      Width           =   3015
      Begin VB.CommandButton cmdMakeInputBox 
         Caption         =   "Make &Input Box"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         Picture         =   "frmInputBoxes.frx":1004
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Defualt"
      Height          =   1935
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   2775
      Begin VB.CommandButton cmdSetDefault 
         Caption         =   "Set &Defualt"
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
         Picture         =   "frmInputBoxes.frx":37A6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtDefault 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Prompt"
      Height          =   1935
      Left            =   2760
      TabIndex        =   3
      Top             =   0
      Width           =   3015
      Begin VB.CommandButton cmdSetPrompt 
         Caption         =   "Set &Prompt"
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
         Left            =   120
         Picture         =   "frmInputBoxes.frx":4070
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtPrompt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Title"
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton cmdSetTitle 
         Caption         =   "Set &Title"
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
         Left            =   120
         Picture         =   "frmInputBoxes.frx":493A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtTitle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmInputBoxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
'-Unload frmInputBoxes
Unload Me
'-make frmMain visible
frmMain.Visible = True
End Sub

Private Sub cmdMakeInputBox_Click()
'-Declare varibles
Dim IB As String
'-Make the input box
IB = InputBox(Prmt, Tl, DF)
End Sub

Private Sub cmdSetDefault_Click()
'-Set the default of the input box to a varible
DF = txtDefault.Text
End Sub

Private Sub cmdSetPrompt_Click()
'-Set the prompt of the input box to a varible
Prmt = txtPrompt.Text
End Sub

Private Sub cmdSetTitle_Click()
'-Set the title of the input box to a varible
Tl = txtTitle.Text
End Sub
