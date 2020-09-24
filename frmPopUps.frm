VERSION 5.00
Begin VB.Form frmPopUps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pop-Up Menus"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "frmPopUps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done Learning About PopUp Menus"
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
      Picture         =   "frmPopUps.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   6375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Popup Menus"
      Height          =   3015
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdClick 
         Caption         =   "Use &Right Click"
         Enabled         =   0   'False
         Height          =   735
         Index           =   1
         Left            =   3240
         Picture         =   "frmPopUps.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CommandButton cmdClick 
         Caption         =   "Use &Left Click"
         Enabled         =   0   'False
         Height          =   735
         Index           =   0
         Left            =   120
         Picture         =   "frmPopUps.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1680
         Width           =   2775
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
         Left            =   2400
         TabIndex        =   3
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblGood 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Good"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblBad 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bad"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label lblHAY 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "How Are You?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   2205
      End
   End
   Begin VB.Menu mnuHowAreYou 
      Caption         =   "HowAreYou"
      Visible         =   0   'False
      Begin VB.Menu mnuGood 
         Caption         =   "Good"
      End
      Begin VB.Menu mnuBad 
         Caption         =   "Bad"
      End
   End
End
Attribute VB_Name = "frmPopUps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-To make a Popup Menu Go to Tools in the menu bar
'-of Visual Basic, Select Menu Editor. From there make
'-a new menu by filing in the Caption and The name
'-For example, Make the caption say Popup and the name be
'-mnuPopup. Uncheck the box that says visible so it will
'-not be on the form, Leave all the sub menus visible
'-hit the Next button then make a new caption call
'-it Yes and name it mnuYes, but hit the arrow showing
'-to the right it will move it over and have 3 dots in
'-front of it, this makes a sub menu. Hit the Next button
'-make a new caption call it No and name it mnuNo Hit the
'-button to move it to the right and thats it for making a
'-popup menu
'-No to make a popup menu double click your form and go to
'-the proc list, default is Click but change it to MouseDown
'-Under Mouse Down enter this code
'-If Button = vbRightButton Then
'-  PopupMenu mnuPopUp
'-End If
'-Its that simple, and to use the menu go unfer the list to
'-the left of the proc list and select your menu, if you
'-using my example then choose mnuYes and put the code
'-under there for what ever you want to do
'-You can always refer to my code if you need help
'-below               \    /
'-                     \  /
'-                      \/

Private Sub cmdClick_Click(Index As Integer)
'-Select index of command button
Select Case Index
    Case "0"    '-Left button
        L = True    '-Use Left button = True
        R = False   '-Use Right button = False
        '-Command Button cmdClick(0) is disabled
        cmdClick(0).Enabled = False
        '-This says if the button that says use left click
        '-is disabled make the button that says use right
        '-click enabled--------\/-------------
        If cmdClick(0).Enabled = False Then
            cmdClick(1).Enabled = True
        End If
        '----------------------/\-----------
    Case "1"    '-Right button
        R = True    '-Use Right button = True
        L = False   '-Use left button = False
        '-Command Button cmdClick(1) is disabled
        cmdClick(1).Enabled = False
        '-This says if the button that says use right click
        '-is disabled make the button that says use left
        '-click enabled--------\/-------------
        If cmdClick(1).Enabled = False Then
            cmdClick(0).Enabled = True
        End If
        '----------------------/\-------------
'-End the case select
End Select
End Sub

Private Sub cmdDone_Click()
'-Unload frmPopUps
Unload Me
'-make frmMain visible
frmMain.Visible = True
End Sub

Private Sub Label1_Click()
'-Setup message box
MsgBox ("You need to choose which button to use before the Popup Menu will work. Click one the How Are You Question with the button you chose to use")
'-Make both buttons enabled
cmdClick(0).Enabled = True
cmdClick(1).Enabled = True
End Sub

Private Sub lblHAY_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-This says if the case 1 was selected from the sub
'-Private Sub cmdClick_Click(Index As Integer) /\
'-was true then it will make you use the      /  \
'-right button to make the popup menu appear
If R = True Then
'-Says if you use the right mouse button then...
'Right mouse button   \/   code also can use 2
'-If Button = 2 Then ...
    If Button = vbRightButton Then
        PopupMenu mnuHowAreYou
    End If
End If
'-Same as comment above where it say right it would say left here
If L = True Then
'-Left mouse button is the same as 2
'-if Button = 2 Then...
    If Button = vbLeftButton Then
        PopupMenu mnuHowAreYou
    End If
End If
End Sub

Private Sub mnuBad_Click()
'-If you select Bad from the menu it makes a sign appear
lblBad.Visible = True
'-This makes so that only good or bad can be up
'-if bgood is visible and they choose bad
'-it makes it so that good will disapear
'-----------\/-------------------------
If lblGood.Visible = True Then
    lblGood.Visible = False
End If
'-----------/\-------------------------
End Sub

Private Sub mnuGood_Click()
'-If you select Good from the menu it makes a sign appear
lblGood.Visible = True
'-This makes so that only good or bad can be up
'-if bad is visible and they choose good
'-it makes it so that bad will disapear
'-----------\/-------------------------
If lblBad.Visible = True Then
    lblBad.Visible = False
End If
'-----------/\-------------------------
End Sub
