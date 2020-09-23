VERSION 5.00
Begin VB.Form frmShell 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shell Execute"
   ClientHeight    =   4695
   ClientLeft      =   3615
   ClientTop       =   2430
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdWinHelp 
      BackColor       =   &H00C0FFFF&
      Caption         =   "WinHelp"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Exit"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdExecute 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Execute"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1455
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4695
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C0FFFF&
      Height          =   1890
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0FFFF&
      Height          =   2040
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   4695
   End
End
Attribute VB_Name = "frmShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdExecute_Click()
Dim sPath As String

If Right(File1.Path, 1) <> "\" Then
    sPath = File1.Path & "\"
Else
    sPath = File1.Path
End If

Call ShellExecute(0, "Open", sPath & File1.List(File1.ListIndex), vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub cmdExit_Click()
Unload Me
End
End Sub

Private Sub cmdWinHelp_Click()
frmHelp.Show
frmShell.Hide
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub


Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub





