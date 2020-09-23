VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHelp 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finding And Displaying Help Files"
   ClientHeight    =   2595
   ClientLeft      =   2400
   ClientTop       =   2085
   ClientWidth     =   4725
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label lblBack 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label lblIndex 
      BackStyle       =   0  'Transparent
      Caption         =   "Display Index"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   2445
   End
   Begin VB.Label lblContents 
      BackStyle       =   0  'Transparent
      Caption         =   "Display Contents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblDisplay 
      BackStyle       =   0  'Transparent
      Caption         =   "Display Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblPick 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oWinHelp As CWinHelp
Private Sub Form_Load()
    ' Basic object initialization
    Set oWinHelp = New CWinHelp
End Sub
Private Sub lblBack_Click()
frmShell.Show
frmHelp.Hide
End Sub
Private Sub lblContents_Click()
' Displays the help file content
    If Len(Text1.Text) Then
        oWinHelp.ShowContents
    Else
        MsgBox "Please select a help file first.", vbInformation
    End If
End Sub
Private Sub lblDisplay_Click()
 ' Displays the help-on-help
    If Len(Text1.Text) Then
        oWinHelp.ShowHelpOnHelp
    Else
        MsgBox "Please select a help file first.", vbInformation

    End If
End Sub
Private Sub lblIndex_Click()
' Displays the help file index
    If Len(Text1.Text) Then
        oWinHelp.ShowFinder
    Else
        MsgBox "Please select a help file first.", vbInformation

    End If
End Sub
Private Sub lblPick_Click()
 ' Gives you the possibility to select a Windows help file
    With CommonDialog1
        .DialogTitle = "Select a help file"
        .Filter = "Help files (*.hlp)|*.hlp"
        .ShowOpen
        If Len(.FileName) Then
            Text1.Text = .FileName
            oWinHelp.FileName = Text1.Text
        End If
    End With
End Sub
