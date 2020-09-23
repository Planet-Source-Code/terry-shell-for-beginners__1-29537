VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWinHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'*****************************************************************************************
'* Class       : CWinHelp
'* Description : Class for working with Windows Help.
'* Notes       : Use this class to implement a Help system in your application. With it
'*               you can display help file contents, index, specific topics etc.
'*****************************************************************************************

Option Explicit

' Error handling definitions
Private Const E_ERR_BASE = 18460 + vbObjectError
Public Enum EErrWinHelp
    eErrWinHelp_NotInitialized = E_ERR_BASE + 1
    eErrWinHelp_InvalidFileName
    eErrWinHelp_FileNotFound
    eErrWinHelp_InvalidWindowHandle
    eErrWinHelp_CallToWinHelpFailed
    eErrWinHelp_ComponentFailure
End Enum
Private Const S_ERR_NotIntialized = "Object not properly initialized"
Private Const S_ERR_InvalidFileName = "Invalid file name"
Private Const S_ERR_FileNotFound = "File not found"
Private Const S_ERR_InvalidWindowHandle = "Invalid window handle"
Private Const S_ERR_CallToWinHelpFailed = "Call to WinHelp failed"
Private Const S_ERR_ComponentFailure = "CWinHelp component failure"

' Private class constants
Private Const HELP_COMMAND = &H102&
Private Const HELP_CONTENTS = &H3&
Private Const HELP_CONTEXT = &H1
Private Const HELP_CONTEXTPOPUP = &H8&
Private Const HELP_FINDER = &HB&
Private Const HELP_FORCEFILE = &H9&
Private Const HELP_HELPONHELP = &H4
Private Const HELP_INDEX = &H3
Private Const HELP_KEY = &H101
Private Const HELP_MULTIKEY = &H201&
Private Const HELP_PARTIALKEY = &H105&
Private Const HELP_QUIT = &H2
Private Const HELP_SETCONTENTS = &H5&
Private Const HELP_SETINDEX = &H5
Private Const HELP_SETWINPOS = &H203&
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_LENGTH = 512

' Private class type definitions
Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_LENGTH
    cAlternate          As String * 14
End Type

' Private class API function declarations
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpFileName As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

' Private variables to hold property values
Private m_sFileName     As String
Private m_lHWndOwner    As Long


'*****************************************************************************************
'* Sub         : Init
'* Notes       : Use this routine for basic object initialization.
'*****************************************************************************************
Public Sub Init(FileName As String, Optional HWndOwner As Long = 0)
    On Error GoTo hComponentFailure

    If Len(Trim$(FileName)) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_InvalidFileName, App.EXEName & ".CWinHelp", S_ERR_InvalidFileName
    End If
    
    If Not FileExists(FileName) Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_FileNotFound, App.EXEName & ".CWinHelp", S_ERR_FileNotFound
    End If

    m_sFileName = FileName
    
    If HWndOwner > 0 Then
        If IsWindow(HWndOwner) = 0 Then
            On Error GoTo 0
            Err.Raise eErrWinHelp_InvalidWindowHandle, App.EXEName & ".CWinHelp", S_ERR_InvalidWindowHandle
        End If
    End If
    
    m_lHWndOwner = HWndOwner

    Exit Sub

hComponentFailure:
    Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Property    : FileName
'* Notes       : Returns or sets the path, if necessary, and the name of the help file
'*               (*.hlp) to display.
'*****************************************************************************************
Public Property Get FileName() As String
    On Error GoTo hComponentFailure
    
    FileName = m_sFileName

    Exit Property

hComponentFailure:
    Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
End Property

Public Property Let FileName(Value As String)
    On Error GoTo hComponentFailure
    
    If Len(Trim$(Value)) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_InvalidFileName, App.EXEName & ".CWinHelp", S_ERR_InvalidFileName
    End If
    
    If Not FileExists(Value) Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_FileNotFound, App.EXEName & ".CWinHelp", S_ERR_FileNotFound
    End If

    m_sFileName = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : HWndOwner
'* Notes       : Returns or sets the handle identifying the window requesting Help.
'*****************************************************************************************
Public Property Get HWndOwner() As Long
    On Error GoTo hComponentFailure
    
    HWndOwner = m_lHWndOwner

    Exit Property

hComponentFailure:
    Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
End Property

Public Property Let HWndOwner(hwnd As Long)
    On Error GoTo hComponentFailure
    
    If IsWindow(hwnd) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_InvalidWindowHandle, App.EXEName & ".CWinHelp", S_ERR_InvalidWindowHandle
    End If
        
    m_lHWndOwner = hwnd

    Exit Property

hComponentFailure:
    Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Sub         : CloseFile
'* Notes       : Informs Windows Help that it is no longer needed.
'*****************************************************************************************
Public Sub CloseFile()
    On Error GoTo hComponentFailure
    
    If Len(m_sFileName) > 0 Then
    
        If WinHelp(m_lHWndOwner, m_sFileName, HELP_QUIT, "") = 0 Then
            On Error GoTo 0
            Err.Raise eErrWinHelp_CallToWinHelpFailed, App.EXEName & ".CWinHelp", S_ERR_CallToWinHelpFailed
        End If
    End If
    
    Exit Sub

hComponentFailure:
    Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Sub         : RunMacro
'* Notes       : Executes a Help macro or macro string in the current help file.
'*****************************************************************************************
Public Sub RunMacro(Macro As String)
    On Error GoTo hComponentFailure
    
    Dim m_APIRet As Long
    
    If Len(m_sFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_NotInitialized, App.EXEName & ".CWinHelp", S_ERR_NotIntialized
    End If
    
    If WinHelp(m_lHWndOwner, m_sFileName, HELP_COMMAND, Macro) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_CallToWinHelpFailed, App.EXEName & ".CWinHelp", S_ERR_CallToWinHelpFailed
    End If

    Exit Sub

hComponentFailure:
    Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Sub         : Search
'* Notes       : Displays the topic in the keyword table that matches the specified
'*               keyword, if there is an exact match, in the current help file. If there
'*               is more than one match, displays the Topics Found dialog box.
'*****************************************************************************************
Public Sub Search(Keyword As String)
    On Error GoTo hComponentFailure
    
    Dim m_APIRet As Long
    
    If Len(m_sFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_NotInitialized, App.EXEName & ".CWinHelp", S_ERR_NotIntialized
    End If
    
    If WinHelp(m_lHWndOwner, m_sFileName, HELP_PARTIALKEY, Keyword) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_CallToWinHelpFailed, App.EXEName & ".CWinHelp", S_ERR_CallToWinHelpFailed
    End If

    Exit Sub

hComponentFailure:
    Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Sub         : ShowContents
'* Notes       : Displays the Contents page of the current help file.
'*****************************************************************************************
Public Sub ShowContents()
    On Error GoTo hComponentFailure
    
    Dim m_APIRet As Long
    
    If Len(m_sFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_NotInitialized, App.EXEName & ".CWinHelp", S_ERR_NotIntialized
    End If
    
    If WinHelp(m_lHWndOwner, m_sFileName, HELP_CONTENTS, 0&) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_CallToWinHelpFailed, App.EXEName & ".CWinHelp", S_ERR_CallToWinHelpFailed
    End If

    Exit Sub

hComponentFailure:
    Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Sub         : ShowContext
'* Notes       : Displays the topic identified by the specified context identifier in the
'*               current help file.
'*****************************************************************************************
Public Sub ShowContext(ContextID As Long)
    On Error GoTo hComponentFailure
    
    Dim m_APIRet As Long
    
    If Len(m_sFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_NotInitialized, App.EXEName & ".CWinHelp", S_ERR_NotIntialized
    End If
    
    If WinHelp(m_lHWndOwner, m_sFileName, HELP_CONTEXT, ContextID) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_CallToWinHelpFailed, App.EXEName & ".CWinHelp", S_ERR_CallToWinHelpFailed
    End If

    Exit Sub

hComponentFailure:
    Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Sub         : ShowFinder
'* Notes       : Displays the Windows 9X Help Topics dialog box for the current help file.
'*               This method should be used instead of the ShowContents method.
'*****************************************************************************************
Public Sub ShowFinder()
    On Error GoTo hComponentFailure
    
    Dim m_APIRet As Long
    
    If Len(m_sFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_NotInitialized, App.EXEName & ".CWinHelp", S_ERR_NotIntialized
    End If
    
    If WinHelp(m_lHWndOwner, m_sFileName, HELP_FINDER, 0&) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_CallToWinHelpFailed, App.EXEName & ".CWinHelp", S_ERR_CallToWinHelpFailed
    End If

    Exit Sub

hComponentFailure:
    Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Sub         : ShowHelpOnHelp
'* Notes       : Displays Help on how to use Windows Help, if the WINHLP32.HLP file is
'*               available.
'*****************************************************************************************
Public Sub ShowHelpOnHelp()
    On Error GoTo hComponentFailure
    
    Dim m_APIRet As Long
    
    If Len(m_sFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_NotInitialized, App.EXEName & ".CWinHelp", S_ERR_NotIntialized
    End If
    
    If WinHelp(m_lHWndOwner, m_sFileName, HELP_HELPONHELP, 0&) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_CallToWinHelpFailed, App.EXEName & ".CWinHelp", S_ERR_CallToWinHelpFailed
    End If

    Exit Sub

hComponentFailure:
    Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Sub         : ShowSearch
'* Notes       : Displays the current help file with the Search dialog active.
'*****************************************************************************************
Public Sub ShowSearch()
    On Error GoTo hComponentFailure
    
    Dim m_APIRet As Long
    
    If Len(m_sFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_NotInitialized, App.EXEName & ".CWinHelp", S_ERR_NotIntialized
    End If
    
    If WinHelp(m_lHWndOwner, m_sFileName, HELP_PARTIALKEY, "") = 0 Then
        On Error GoTo 0
        Err.Raise eErrWinHelp_CallToWinHelpFailed, App.EXEName & ".CWinHelp", S_ERR_CallToWinHelpFailed
    End If

    Exit Sub

hComponentFailure:
    Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Function    : FileExists
'* Notes       : API-based routine to check for a file's existence.
'*               Returns True if the specified file exists, False otherwise.
'*****************************************************************************************
Private Function FileExists(FileName As String) As Boolean
    On Error GoTo hComponentFailure
    
    Dim wfd    As WIN32_FIND_DATA
    Dim hFile  As Long
    
    FileExists = False
    
    hFile = FindFirstFile(FileName, wfd)
    If hFile <> INVALID_HANDLE_VALUE Then
        FileExists = True
        
        If FindClose(hFile) = 0 Then
            On Error GoTo 0
            Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
        End If
    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Sub         : Class_Initialize
'* Notes       : Class data space initialization.
'*****************************************************************************************
Private Sub Class_Initialize()
    On Error GoTo hComponentFailure
    
    m_sFileName = ""
    m_lHWndOwner = 0

    Exit Sub

hComponentFailure:
    Err.Raise eErrWinHelp_ComponentFailure, App.EXEName & ".CWinHelp", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Sub         : Class_Terminate
'* Notes       : Perform object cleanup.
'*****************************************************************************************
Private Sub Class_Terminate()
    On Error Resume Next
    
    If Len(Trim$(m_sFileName)) > 0 Then WinHelp m_lHWndOwner, m_sFileName, HELP_QUIT, 0&
End Sub
