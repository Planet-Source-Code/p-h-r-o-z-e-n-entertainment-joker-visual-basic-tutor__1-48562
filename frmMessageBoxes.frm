VERSION 5.00
Begin VB.Form frmMessageBoxes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message Boxes"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "frmMessageBoxes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done Learning About Message Boxes"
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
      Picture         =   "frmMessageBoxes.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4440
      Width           =   5535
   End
   Begin VB.Frame Frame5 
      Caption         =   "Genarate Message Boxes"
      Height          =   2535
      Left            =   0
      TabIndex        =   23
      Top             =   1800
      Width           =   5535
      Begin VB.CommandButton cmdMake 
         Caption         =   "Make Message Box Using &Style Options"
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
         Left            =   120
         Picture         =   "frmMessageBoxes.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   5295
      End
      Begin VB.CommandButton cmdMakes 
         Caption         =   "Make Message Box Using &Button Options"
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
         Picture         =   "frmMessageBoxes.frx":0EDE
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1440
         Width           =   5295
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Buttons"
      Height          =   2535
      Left            =   5520
      TabIndex        =   21
      Top             =   2640
      Width           =   2535
      Begin VB.CheckBox chkButton 
         Caption         =   "Yes No Cancel"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "Retry Cancel"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "Ok Cancel"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "Abort Retry Ignore"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "Ok Only"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkButton 
         Caption         =   "Yes No"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Prompt"
      Height          =   1815
      Left            =   2640
      TabIndex        =   18
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdPrompt 
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
         Height          =   855
         Left            =   120
         Picture         =   "frmMessageBoxes.frx":17A8
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   2535
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
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Title "
      Height          =   1815
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   2655
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
         Height          =   855
         Left            =   120
         Picture         =   "frmMessageBoxes.frx":2072
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   2295
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
         TabIndex        =   0
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Style"
      Height          =   1815
      Left            =   5520
      TabIndex        =   19
      Top             =   0
      Width           =   2535
      Begin VB.CheckBox chkBox 
         Caption         =   "Information"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox chkBox 
         Caption         =   "Critical"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CheckBox chkBox 
         Caption         =   "Question"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox chkBox 
         Caption         =   "Exclamation"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Height          =   495
         Left            =   1320
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Label Label2 
      Caption         =   "                       /\                            Only Select One At A Time                              \/"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   5640
      TabIndex        =   22
      Top             =   1920
      Width           =   2295
   End
End
Attribute VB_Name = "frmMessageBoxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkBox_Click(Index As Integer)
'-select index of type choosen
Select Case Index
    Case "0"    '-Exclamation
        If chkBox(0).Value = 0 Then '-Not checked
            Exclam = False
        ElseIf chkBox(0).Value = 1 Then '-Checked
            Exclam = True
        End If
    Case "1"    '-Question
        If chkBox(1).Value = 0 Then '-Not checked
            Ques = False
        ElseIf chkBox(1).Value = 1 Then '-Checked
            Ques = True
        End If
    Case "2"    '-Critical
        If chkBox(2).Value = 0 Then '-Not checked
            Crit = False
        ElseIf chkBox(2).Value = 1 Then '-Checked
            Crit = True
        End If
    Case "3"    '-Information
        If chkBox(3).Value = 0 Then '-Not checked
            Info = False
        ElseIf chkBox(3).Value = 1 Then '-Checked
            Info = True
        End If
End Select
End Sub

Private Sub chkButton_Click(Index As Integer)
'-select index of type choosen
Select Case Index
    Case "0"    '-Yes No
        If chkButton(0).Value = 0 Then '-Not checked
            YN = False
        ElseIf chkButton(0).Value = 1 Then '-Checked
            YN = True
        End If
    Case "1"    '-Ok Only
        If chkButton(1).Value = 0 Then '-Not checked
            OkOnly = False
        ElseIf chkButton(1).Value = 1 Then '-Checked
            OkOnly = True
        End If
    Case "2"    '-Abort Retry Ignore
        If chkButton(2).Value = 0 Then '-Not checked
            ARI = False
        ElseIf chkButton(2).Value = 1 Then '-Checked
            ARI = True
        End If
    Case "3"    '-Ok Cancel
        If chkButton(3).Value = 0 Then '-Not checked
            OC = False
        ElseIf chkButton(3).Value = 1 Then '-Checked
            OC = True
        End If
    Case "4"    '-Retry Cancel
        If chkButton(4).Value = 0 Then '-Not checked
            RC = False
        ElseIf chkButton(4).Value = 1 Then '-Checked
            RC = True
        End If
    Case "5"    '-Yes No Cancel
        If chkButton(5).Value = 0 Then '-Not checked
            YNC = False
        ElseIf chkButton(5).Value = 1 Then '-Checked
            YNC = True
        End If
End Select
End Sub

Private Sub cmdDone_Click()
'-Unload frmMessageboxes
Unload Me
'-Make frmMain visible
frmMain.Visible = True
End Sub

Private Sub cmdMake_Click()
'-Generate the Message Box
'-Pro = the prompt set for the messagebox
'-Til = title set for the messagebox
Dim Message
If Exclam Then
    Message = MsgBox(Pro, vbExclamation, Til)
End If
If Crit Then
    Message = MsgBox(Pro, vbCritical, Til)
End If
If Ques Then
    Message = MsgBox(Pro, vbQuestion, Til)
End If
If Info Then
    Message = MsgBox(Pro, vbInformation, Til)
End If
End Sub

Private Sub cmdMakes_Click()
'-Generate the Message Box
'-Pro = the prompt set for the messagebox
'-Til = title set for the messagebox
Dim Message
If YN Then
    Message = MsgBox(Pro, vbYesNo, Til)
End If
If OkOnly Then
    Message = MsgBox(Pro, vbOKOnly, Til)
End If
If ARI Then
    Message = MsgBox(Pro, vbAbortRetryIgnore, Til)
End If
If OC Then
    Message = MsgBox(Pro, vbOKCancel, Til)
End If
If RC Then
    Message = MsgBox(Pro, vbRetryCancel, Til)
End If
If YNC Then
    Message = MsgBox(Pro, vbYesNoCancel, Til)
End If
End Sub

Private Sub cmdPrompt_Click()
'-Sets the text to the prompt
Pro = txtPrompt.Text
End Sub

Private Sub cmdSetTitle_Click()
'-Sets the text to the Title
Til = txtTitle.Text
End Sub




