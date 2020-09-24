VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmToolBar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tool Bars"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   Icon            =   "frmToolBar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   3465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done Learning About Toolbars"
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
      Picture         =   "frmToolBar.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "ZEO Toolbar"
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.PictureBox picStart 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   720
         Picture         =   "frmToolBar.frx":0BD4
         ScaleHeight     =   3705
         ScaleWidth      =   1065
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start"
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
         Left            =   720
         Picture         =   "frmToolBar.frx":22AB
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4080
         Width           =   2175
      End
      Begin MSComctlLib.Toolbar tbStart 
         Height          =   810
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1429
         ButtonWidth     =   1984
         ButtonHeight    =   1376
         Appearance      =   1
         Style           =   1
         ImageList       =   "ilPics"
         DisabledImageList=   "ilPics"
         HotImageList    =   "ilPics"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Programs"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Control Panels"
               ImageIndex      =   3
               Object.Width           =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Run"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Shut Down..."
               ImageIndex      =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin MSComctlLib.ImageList ilPics 
         Left            =   2760
         Top             =   4320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmToolBar.frx":25B5
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmToolBar.frx":28D1
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmToolBar.frx":2BED
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmToolBar.frx":34C9
               Key             =   ""
            EndProperty
         EndProperty
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
         Height          =   435
         Left            =   1200
         TabIndex        =   4
         Top             =   5040
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmToolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
'-Unload frmToolbars
Unload Me
'-Make frmMain visible
frmMain.Visible = True
End Sub

Private Sub cmdStart_Click()
'-Make the toolbar visible
tbStart.Visible = True
'-Make the launch picture visible
picStart.Visible = True
End Sub

Private Sub Label1_Click()
'-Setup messagebox
MsgBox ("To see how to set up a toolbar look at the Properties of the toolbar and go under custom to see how its set up.")
End Sub

Private Sub picStart_Click()
'-Setup messagebox
MsgBox ("Don't click me! Click the toolbar")
End Sub

Private Sub tbStart_ButtonClick(ByVal Button As MSComctlLib.Button)
'-select which button they hit
Select Case Button.Caption
    Case "Shut Down..."    '-Shut Down
        '-Setup messagebox
        MsgBox ("This is to show you how to make a tool bar! You think I would shut the computer down?")
        '-make the toolbar invisible and the picture
        tbStart.Visible = False
        picStart.Visible = False
    Case "Run"    '-Run
        '-Setup messagebox
        MsgBox ("Where do you think you are running to?")
        '-make the toolbar invisible and the picture
        tbStart.Visible = False
        picStart.Visible = False
    Case "Control Panels"    '-Control panels
        '-Setup messagebox
        MsgBox ("This is the wrong toolbar to get to the control panels!")
        '-make the toolbar invisible and the picture
        tbStart.Visible = False
        picStart.Visible = False
    '-Case "4" is a space so i skiped it
    Case "Programs"    '-Programs
        '-Setup messagebox
        MsgBox ("You are already running the program!")
        '-make the toolbar invisible and the picture
        tbStart.Visible = False
        picStart.Visible = False
End Select
End Sub
