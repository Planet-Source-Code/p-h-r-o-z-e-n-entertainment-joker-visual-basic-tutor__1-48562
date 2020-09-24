VERSION 5.00
Begin VB.Form frmArrays 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arrays"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "frmArrays.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done Learning About Arrays"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   0
      Picture         =   "frmArrays.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   7095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Arrys"
      Height          =   3375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton cmdLReset 
         Caption         =   "Rest &Label Array"
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
         Left            =   3600
         Picture         =   "frmArrays.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CommandButton cmdArray 
         Caption         =   "Array 2"
         Height          =   495
         Index           =   1
         Left            =   1080
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdArray 
         Caption         =   "Array 3"
         Height          =   495
         Index           =   2
         Left            =   1920
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdArray 
         Caption         =   "Array 4"
         Height          =   495
         Index           =   3
         Left            =   2760
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdArray 
         Caption         =   "Array 7"
         Height          =   495
         Index           =   4
         Left            =   1920
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdArray 
         Caption         =   "Array 6"
         Height          =   495
         Index           =   5
         Left            =   1080
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdArray 
         Caption         =   "Array 5"
         Height          =   495
         Index           =   6
         Left            =   240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdArray 
         Caption         =   "Array 8"
         Height          =   495
         Index           =   7
         Left            =   2760
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdCReset 
         Caption         =   "Reset &Command Button Array"
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
         Left            =   3600
         Picture         =   "frmArrays.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton cmdArray 
         Caption         =   "Array 1"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   735
      End
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
         Left            =   2280
         TabIndex        =   16
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label lblArray 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblArray 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   1800
         TabIndex        =   14
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblArray 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1800
         TabIndex        =   13
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblArray 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdArray_Click(Index As Integer)
'-Declare varibles
Dim X As Integer
'-Set up loop to make the command buttons
'-invisible when clicked on
For X = 0 To 7 '-For 0 because thats the first index of my
               '-command button and 7 for the last
    If cmdArray(X) Then '-If cmdArray(Index of button) clicked
                        '-Then if its visible make it
                        '-invisible
        If cmdArray(X).Visible = True Then
            cmdArray(X).Visible = False
        End If
    End If
'-Continue the loop
Next X
End Sub

Private Sub cmdCReset_Click()
'-Declare varibles
Dim X As Integer
'-Set up loop to make the command buttons
'-visible when the reset button is clicked
For X = 0 To 7 '-For 0 because thats the first index of my
               '-command button and 7 for the last
    cmdArray(X).Visible = True
Next X
End Sub

Private Sub cmdDone_Click()
'-Unload frmArrays
Unload Me
'-Make frmMain visible
frmMain.Visible = True
End Sub

Private Sub cmdLReset_Click()
'-Declare varibles
Dim X As Integer
'-Set up loop to make the command buttons
'-visible when the reset button is clicked
For X = 0 To 3 '-For 0 because thats the first index of my
               '-labels and 0 for the last
    lblArray(X).Caption = ""
Next X
End Sub

Private Sub Label1_Click()
'-Setup messagebox
MsgBox ("Arrays are objects with the same name, but a diffrent Index number try out the sample of arrays here.")
End Sub

Private Sub lblArray_Click(Index As Integer)
'-Dim Varibles
Dim Word As String
Dim X As Integer
'-Setup Input Box to gather data from the user
Word = InputBox("Enter a word to be displayed as an array", "Enter Word Array", "Hello")
'-Setup loop to use the input from the Inputbox Word /\
For X = 0 To 3 '-0 is the start of my label array and
               '-3 is the last
    '-Label named lblArray(Idex of label) and its caption property
    '-will say what was gathered from the Input Box Word
    lblArray(X).Caption = Word
'-Continue the loop
Next X
End Sub
