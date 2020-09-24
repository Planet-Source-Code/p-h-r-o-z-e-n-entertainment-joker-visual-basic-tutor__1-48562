VERSION 5.00
Begin VB.Form FrmFontText 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fonts and Text"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "FrmFontText.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done Learning About Fonts and Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      Picture         =   "FrmFontText.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3600
      Width           =   8055
   End
   Begin VB.Frame Frame5 
      Caption         =   "Font Color"
      Height          =   3615
      Left            =   6000
      TabIndex        =   22
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton cmdColor 
         Caption         =   "Yellow"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "Red"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "Green"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "Blue"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "Black"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click Me!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   3000
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Passwords"
      Height          =   1815
      Left            =   0
      TabIndex        =   21
      Top             =   1800
      Width           =   2655
      Begin VB.CommandButton cmdPassword 
         Caption         =   "&Password"
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
         Picture         =   "FrmFontText.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtPassword 
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
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Load Text"
      Height          =   1815
      Left            =   2640
      TabIndex        =   20
      Top             =   1800
      Width           =   3375
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load Text"
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
         Left            =   360
         Picture         =   "FrmFontText.frx":0EDE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtLoad 
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
         TabIndex        =   9
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Saving Text"
      Height          =   1815
      Left            =   2640
      TabIndex        =   19
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Text"
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
         Left            =   360
         Picture         =   "FrmFontText.frx":17A8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Text Options"
      Height          =   1815
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   2655
      Begin VB.CheckBox chkItalic 
         Caption         =   "Italic"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox chkStrike 
         Caption         =   "Strike "
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox chkBold 
         Caption         =   "Bold"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkUnline 
         Caption         =   "Under Line"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtText 
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
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FrmFontText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkBold_Click()
'-If the check box is not checked
If chkBold.Value = 0 Then
    '-Then the txtText.Text is not bold
    txtText.FontBold = False
'-If the check box is checked
ElseIf chkBold.Value = 1 Then
    '-Then the txtText.Text is bold
    txtText.FontBold = True
End If
End Sub

Private Sub chkItalic_Click()
'-If the check box is not checked
If chkItalic.Value = 0 Then
    '-Then the txtText.Text is not italic
    txtText.FontItalic = False
'-If the check box is checked
ElseIf chkItalic.Value = 1 Then
    '-Then the txtText.Text is italic
    txtText.FontItalic = True
End If
End Sub

Private Sub chkStrike_Click()
'-If the check box is not checked
If chkStrike.Value = 0 Then
    '-Then the txtText.Text is not striked thru
    txtText.FontStrikethru = False
'-If the check box is checked
ElseIf chkStrike.Value = 1 Then
    '-Then the txtText.Text is striked thru
    txtText.FontStrikethru = True
End If
End Sub

Private Sub chkUnline_Click()
'-If the check box is not checked
If chkUnline.Value = 0 Then
    '-Then the txtText.Text is not underlined
    txtText.FontUnderline = False
'-If the check box is checked
ElseIf chkUnline.Value = 1 Then
    '-Then the txtText.Text is underlined
    txtText.FontUnderline = True
End If
End Sub

Private Sub cmdColor_Click(Index As Integer)
'-Select the index of the command button choosen
Select Case Index
    Case "0"    '-Black
        txtText.ForeColor = &H0&    '-Changes the forecolor
        txtSave.ForeColor = &H0&    '-of txtText, txtSave
        txtLoad.ForeColor = &H0&    '-txtLoad, txtPassword
        txtPassword.ForeColor = &H0&
    Case "1"    '-Blue
        txtText.ForeColor = &HFF0000    '-Ditto ^
        txtSave.ForeColor = &HFF0000    '-     / \
        txtLoad.ForeColor = &HFF0000
        txtPassword.ForeColor = &HFF0000
    Case "2"    '-Green
        txtText.ForeColor = &HFF00&     '-Ditto ^
        txtSave.ForeColor = &HFF00&     '-     / \
        txtLoad.ForeColor = &HFF00&
        txtPassword.ForeColor = &HFF00&
    Case "3"    '-Red
        txtText.ForeColor = &HFF&       '-Ditto ^
        txtSave.ForeColor = &HFF&       '-     / \
        txtLoad.ForeColor = &HFF&
        txtPassword.ForeColor = &HFF&
    Case "4"    '-Yellow
        txtText.ForeColor = &HFFFF&     '-Ditto ^
        txtSave.ForeColor = &HFFFF&     '-     / \
        txtLoad.ForeColor = &HFFFF&
        txtPassword.ForeColor = &HFFFF&
End Select
End Sub

Private Sub cmdDone_Click()
'-Unload frmFontText
Unload Me
'-make frmMain visible
frmMain.Visible = True
End Sub

Private Sub cmdLoad_Click()
'-If it messes up
On Error GoTo errhandle
'-Declare Varibles
Dim directory$
Dim MyString$
'-Loads "Text.txt" from my prog's path and the Saves Folder
directory$ = App.Path & "\Saves\Text.txt"
'-Opens the files and reads it
Open directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
    Wend
    '-Puts it in the text box
    txtLoad.Text = MyString$
    '-Closes the file
    Close #1
'-If it does mess up it will not do anything
errhandle:
End Sub

Private Sub cmdPassword_Click()
'-Declare Varibles
Dim PassChar As String
'-PassChar is a Inputbox made to gather information
PassChar = InputBox("Enter one character for the Password Character", "Password Character", "*")
'-Changes the Password Character to what the user enters
txtPassword.PasswordChar = PassChar
End Sub

Private Sub cmdSave_Click()
'-Declare Varibles
Dim directory$
'-On error it will not save
On Error Resume Next
    '-Save to where my program is and in the folder "Saves"
    directory$ = App.Path & "\Saves\Text.txt"
    '-Writing the file
    Open directory$ For Output As #1
    Print #1, (txtSave)
    '-Close the file
    Close #1
'-Setup msgbox
MsgBox ("Your text was saved as " & App.Path & "\Saves\Text.txt")
End Sub

Private Sub Label1_Click()
'-Message Box pops up when you click the label
MsgBox ("There are many more colors, but I did not want to go through them all!")
End Sub

