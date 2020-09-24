VERSION 5.00
Begin VB.Form frmPictures 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pictures"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmPictures.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done Learning About Pictures"
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
      Picture         =   "frmPictures.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Saving Pictures"
      Height          =   1935
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   4455
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Picture"
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
         Left            =   2400
         Picture         =   "frmPictures.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.PictureBox picSave 
         AutoRedraw      =   -1  'True
         FillStyle       =   0  'Solid
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1395
         ScaleWidth      =   1995
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label2 
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
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Load Picture"
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load Picture"
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
         Left            =   2400
         Picture         =   "frmPictures.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.PictureBox picLoad 
         AutoRedraw      =   -1  'True
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1515
         ScaleWidth      =   1995
         TabIndex        =   1
         Top             =   360
         Width           =   2055
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
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
'-Unload frmPictures
Unload Me
'-Make frmMain visible
frmMain.Visible = True
End Sub

Private Sub cmdLoad_Click()
'-Loads pic into picLoad, looks for an icon where my prog
'-is and the icon names der
Randomize   '-Randomize the random number
'-Declare varibles

Dim X As Integer
X = Int(Rnd() * 4 + 1) '-Setup a random integer between 1 and 4

'-Selects then icon to be loaded based on the random number
'------------------------------------------\/---------------
If X = 1 Then
    picLoad.Picture = LoadPicture(App.Path & "\Icons\der.ico")
ElseIf X = 2 Then
    picLoad.Picture = LoadPicture(App.Path & "\Icons\folder_gradient.ico")
ElseIf X = 3 Then
    picLoad.Picture = LoadPicture(App.Path & "\Icons\emark.ico")
ElseIf X = 4 Then
    picLoad.Picture = LoadPicture(App.Path & "\Icons\folder.ico")
End If
'-done with the random number selection----/\----------------
End Sub

Private Sub cmdSave_Click()
'-If it messes up
On Error Resume Next
'-Save the pic to my prog's app.path and the folder
'-Saves, and the pic is saved as picture
SavePicture picSave.Image, (App.Path & "\Saves\Picture.bmp")
'-setup msgbox
MsgBox ("Your picture was saved as " & App.Path & "\Saves\Picture.bmp")
End Sub

Private Sub Label1_Click()
'-Setup Messagebox to be displayed
MsgBox ("When you click the load picture button diffent pictures will be loaded each time.")
End Sub

Private Sub Label2_Click()
'-Setup message box
MsgBox ("Click on the picture box to change colors, then save the picture.")
End Sub

Private Sub picSave_Click()
'-Declare varibles
Dim C As Integer

Randomize   '-Ranomize the random number
C = Int(Rnd() * 15 + 1)     '-Setup a random integer between 1 and 16
'-Fills the picSave box with color
picSave.BackColor = QBColor(C)
End Sub
