VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Learn Visual Basic - Joker - |P|h|r|o|z|e|n| Entertainment"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDates 
      Caption         =   " Dates&\Times"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      Picture         =   "frmMain.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdInputBoxes 
      Caption         =   "Inpu&t Boxes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      Picture         =   "frmMain.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdStrings 
      Caption         =   "Strin&gs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      Picture         =   "frmMain.frx":18CE
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Arra&ys"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "frmMain.frx":2198
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdToolBars 
      Caption         =   "&Tool Bars"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      Picture         =   "frmMain.frx":2A62
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdPopUps 
      Caption         =   "&Pop-Up Menus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      Picture         =   "frmMain.frx":2D6C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdKill 
      Caption         =   "D&elete Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      Picture         =   "frmMain.frx":3636
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdLoops 
      Caption         =   "&Loops"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "frmMain.frx":3940
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit Learn Visual Basic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "frmMain.frx":3C4A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5880
      Width           =   7455
   End
   Begin VB.CommandButton cmdIfThen 
      Caption         =   "&If...Then..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      Picture         =   "frmMain.frx":4514
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdCommD 
      Caption         =   "Common Dial&og"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      Picture         =   "frmMain.frx":4DDE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdHotSpots 
      Caption         =   "&Hot Spots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      Picture         =   "frmMain.frx":56A8
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdRandomNumbers 
      Caption         =   "&Randoms Numbers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "frmMain.frx":5AEA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdSounds 
      Caption         =   "&Sounds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      Picture         =   "frmMain.frx":63B4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdMessageBox 
      Caption         =   "&Message Boxes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      Picture         =   "frmMain.frx":6C7E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdTimer 
      Caption         =   "&Timer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      Picture         =   "frmMain.frx":7548
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton cmdAPI 
      Caption         =   "&API"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "frmMain.frx":810A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton cmdListBoxes 
      Caption         =   "&List Boxes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      Picture         =   "frmMain.frx":8414
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdColors 
      Caption         =   "&Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      Picture         =   "frmMain.frx":871E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdPictures 
      Caption         =   "&Pictures"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      Picture         =   "frmMain.frx":95E8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton CmdFontText 
      Caption         =   "&Fonts\Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "frmMain.frx":98F2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Whenever you see a box that says ""Click Me"" while going thru Learn Visual Basic, click it for more information about that area"
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
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   6720
      Width           =   7455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "What Would You Like To Learn???"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAPI_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmAPI Visible
frmAPI.Visible = True
End Sub

Private Sub cmdArray_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmArraysVisible
frmArrays.Visible = True
End Sub

Private Sub cmdColors_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmColors Visible
frmColors.Visible = True
End Sub

Private Sub cmdCommD_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmCommD Visible
frmCommD.Visible = True
End Sub

Private Sub cmdDates_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmDatesTimes Visible
frmDatesTimes.Visible = True
End Sub

Private Sub cmdExit_Click()
'-Unload frmMain
Unload Me
'-Ends the program
End
End Sub

Private Sub CmdFontText_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmFontText Visible
FrmFontText.Visible = True
End Sub

Private Sub cmdHotSpots_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmHotSpotsVisible
frmHotSpots.Visible = True
End Sub

Private Sub cmdIfThen_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmIfThen Visible
frmIfThen.Visible = True
End Sub

Private Sub cmdInputBoxes_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmInputBoxes Visible
frmInputBoxes.Visible = True
End Sub

Private Sub cmdKill_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmKill Visible
frmKill.Visible = True
End Sub

Private Sub cmdListBoxes_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmListBoxes Visible
frmListBoxes.Visible = True
End Sub

Private Sub cmdLoops_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmLoops Visible
frmLoops.Visible = True
End Sub

Private Sub cmdMessageBox_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmMessageBoxes Visible
frmMessageBoxes.Visible = True
End Sub

Private Sub cmdPictures_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmPictures Visible
frmPictures.Visible = True
End Sub

Private Sub cmdPopUps_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmPopUps Visible
frmPopUps.Visible = True
End Sub

Private Sub cmdRandomNumbers_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmRandomNumbers Visible
frmRandomNumbers.Visible = True
End Sub

Private Sub cmdSounds_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmSounds Visible
frmSounds.Visible = True
End Sub

Private Sub cmdStrings_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmStringsVisible
frmStrings.Visible = True
End Sub

Private Sub cmdTimer_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmTimer Visible
frmTimer.Visible = True
End Sub

Private Sub cmdToolBars_Click()
'-Make frmMain invisible
frmMain.Visible = False
'-Make frmToolBar Visible
frmToolBar.Visible = True
End Sub
