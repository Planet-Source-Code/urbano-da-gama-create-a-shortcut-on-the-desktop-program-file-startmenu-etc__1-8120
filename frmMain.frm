VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Create Shortcut/Icons"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   0
      TabIndex        =   10
      Top             =   -90
      Width           =   5355
      Begin VB.TextBox txtArgs 
         Height          =   375
         Left            =   2190
         TabIndex        =   2
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtTitle 
         Height          =   375
         Left            =   2190
         TabIndex        =   0
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtPath 
         Height          =   375
         Left            =   2190
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
      Begin VB.Frame Frame3 
         Caption         =   "Icon Destination"
         Height          =   1635
         Left            =   90
         TabIndex        =   11
         Top             =   1770
         Width           =   5145
         Begin VB.OptionButton optStartUp 
            Caption         =   "StartMenu"
            Height          =   285
            Left            =   3570
            TabIndex        =   5
            Top             =   330
            Width           =   1275
         End
         Begin VB.OptionButton optDesktop 
            Caption         =   "Desktop"
            Height          =   405
            Left            =   2160
            TabIndex        =   4
            Top             =   300
            Width           =   1185
         End
         Begin VB.OptionButton optProgFiles 
            Caption         =   "Program Group"
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   300
            Width           =   1695
         End
         Begin VB.TextBox txtGroup 
            Height          =   375
            Left            =   2280
            TabIndex        =   6
            Top             =   1140
            Width           =   2775
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Program Group Name *"
            Height          =   345
            Left            =   60
            TabIndex        =   15
            Top             =   1200
            Width           =   2115
         End
         Begin VB.Line Line1 
            X1              =   960
            X2              =   960
            Y1              =   500
            Y2              =   990
         End
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Command Line Args"
         Height          =   345
         Left            =   120
         TabIndex        =   14
         Top             =   1380
         Width           =   1965
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Icon Title *"
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   450
         Width           =   1965
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Program Path *"
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   900
         Width           =   1965
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   0
      TabIndex        =   9
      Top             =   3420
      Width           =   5355
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   400
         Left            =   3120
         TabIndex        =   8
         Top             =   330
         Width           =   1500
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create Icon"
         Height          =   400
         Left            =   840
         TabIndex        =   7
         Top             =   300
         Width           =   1500
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub cmdCancel_Click()
   Unload Me
End Sub

' Urbano DaGama (udgama@rocketmail.com)

' Drop me a line in case you need any help on this program or
' if you liked the code. That will encourage me to create more such
' programs.

Private Sub optDesktop_Click()
   txtGroup.Text = "..\..\Desktop"
   txtGroup.Locked = True
End Sub

Private Sub optProgFiles_Click()
   txtGroup.Text = ""
   txtGroup.Locked = False
End Sub

Private Sub optStartup_Click()
   txtGroup.Text = ".."
   txtGroup.Locked = True
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Create a shortcut of an application to the desktop or a program group or
' in the start menu.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCreate_Click()
On Error GoTo EH
Dim strProgramPath   As String   ' The path of the executable file
Dim strGroup         As String
Dim strProgramIconTitle As String
Dim strProgramArgs   As String
Dim sParent          As String

   If txtPath.Text = "" Then
      MsgBox "Please enter the location of the exe file.", vbInformation, "udgama's shortcut"
      Exit Sub
   End If
   If txtGroup.Text = "" Then
      MsgBox "Please enter the destination of the shortcut.", vbInformation, "udgama's shortcut"
      Exit Sub
   End If
   If txtTitle.Text = "" Then
      MsgBox "Please enter a title for the shortcut.", vbInformation, "udgama's shortcut"
      Exit Sub
   End If
   
   strProgramPath = txtPath.Text
   strGroup = txtGroup.Text
   strProgramIconTitle = txtTitle.Text
   strProgramArgs = txtArgs.Text
   
   sParent = "$(Programs)"
   
   CreateShellLink strProgramPath, strGroup, strProgramArgs, strProgramIconTitle, True, sParent
   
   Exit Sub
EH:
   MsgBox Err.Description
   Exit Sub
End Sub

