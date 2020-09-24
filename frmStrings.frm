VERSION 5.00
Begin VB.Form frmStrings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Strings"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmStrings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done Learning About Strings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   0
      Picture         =   "frmStrings.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   6735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Joining Strings"
      Height          =   3135
      Left            =   0
      TabIndex        =   11
      Top             =   2520
      Width           =   6735
      Begin VB.CommandButton cmdJClear 
         Caption         =   "&Clear Joining Strings Data"
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
         Height          =   975
         Left            =   3600
         Picture         =   "frmStrings.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CommandButton cmdJoin 
         Caption         =   "&Join The Strings"
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
         Height          =   975
         Left            =   120
         Picture         =   "frmStrings.frx":1ECE
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtJ2 
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
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Text            =   "Enter you second part of the string to join"
         Top             =   840
         Width           =   6375
      End
      Begin VB.TextBox txtJ1 
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
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Text            =   "Enter you first part of the string to join"
         Top             =   360
         Width           =   6375
      End
      Begin VB.Label Label2 
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
         Left            =   2520
         TabIndex        =   8
         Top             =   2640
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "String Testing"
      Height          =   2415
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton cmdTClear 
         Caption         =   "Clear &String Test Data"
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
         Height          =   975
         Left            =   3600
         Picture         =   "frmStrings.frx":3680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   2775
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "&Test String"
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
         Height          =   975
         Left            =   120
         Picture         =   "frmStrings.frx":437A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtString 
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
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Text            =   "Type in the string you want to test here!"
         Top             =   240
         Width           =   6495
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
         Left            =   2640
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
'-Unload frmStrings
Unload Me
'-Make frmMain visible
frmMain.Visible = True
End Sub

Private Sub cmdJClear_Click()
'-Clear the text boxes
txtJ1.Text = ""
txtJ2.Text = ""
End Sub

Private Sub cmdJoin_Click()
'-Declare varible
Dim Join1 As String
Dim Join2 As String
'-Get text from the boxes and format it
Join1 = Trim$(txtJ1.Text)
Join2 = Trim$(txtJ2.Text)
'-Setup messagebox
MsgBox ("You strings join together is:" & Chr(13) & Join1 & " " & Join2)
End Sub

Private Sub cmdTClear_Click()
txtString.Text = ""
End Sub

Private Sub cmdTest_Click()
'-Declare varibles
Dim Word As String
'-If nothing is in the text box it will give you a
'-messagebox
'- "" = blank
If txtString.Text = "" Then
    MsgBox ("Please enter text to test.")
End If
'-If the txtString is not blank it will test it
If txtString.Text <> "" Then
    Word = txtString.Text '-Varible Word = whats in the txtString textbox
    '-Just Len(String)
    MsgBox ("The length of your string with the begining and ending spaces is " & Len(Word))
    '-Just Trim$(Len(String))
    MsgBox ("The length of your string with the begining and ending spaces removed is " & Len(Trim$(Word)))
    '-Just UCase(String)
    MsgBox ("The upper case of your string is" & Chr(13) & UCase(Word))
    '-Just LCase(String)
    MsgBox ("The lower case of you string is" & Chr(13) & LCase(Word))
End If
End Sub

Private Sub Label1_Click()
'-Setup messagebox
MsgBox ("Type the string you want to test in the text box, It will test:" & Chr(13) & Chr(13) & "The length of the string in two diffrent ways" & Chr(13) & "Upper case" & Chr(13) & "Lower Case" & Chr(13) & Chr(13) & "So have fun!")
'-Make the rest of the program work
txtString.Enabled = True
cmdTest.Enabled = True
cmdTClear.Enabled = True
'-Clear the txtString box
txtString.Text = ""
End Sub

Private Sub Label2_Click()
'-Setup messagebox
MsgBox ("Joining strings together is called Concatination, but have fun and play around with it.")
'-Make the rest of the program work
txtJ1.Enabled = True
txtJ2.Enabled = True
cmdJoin.Enabled = True
cmdJClear.Enabled = True
'-Clear the text boxes
txtJ1.Text = ""
txtJ2.Text = ""
End Sub
