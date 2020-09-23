Attribute VB_Name = "Module1"
Option Explicit
Public dext As String, fname As String, fname2 As String, ftitle As String
Public FileLoc As String
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Enum STGM
    STGM_DIRECT = &H0&
    STGM_TRANSACTED = &H10000
    STGM_SIMPLE = &H8000000
    STGM_READ = &H0&
    STGM_WRITE = &H1&
    STGM_READWRITE = &H2&
    STGM_SHARE_DENY_NONE = &H40&
    STGM_SHARE_DENY_READ = &H30&
    STGM_SHARE_DENY_WRITE = &H20&
    STGM_SHARE_EXCLUSIVE = &H10&
    STGM_PRIORITY = &H40000
    STGM_DELETEONRELEASE = &H4000000
    STGM_CREATE = &H1000&
    STGM_CONVERT = &H20000
    STGM_FAILIFTHERE = &H0&
    STGM_NOSCRATCH = &H100000
End Enum
'
' Shell Folder Path Constants...
'
' on NT:
'   ..\WinNT\profiles\username
'
' on Windows 9x:
'   ..\Windows
Public Enum SHELLFOLDERS
    CSIDL_DESKTOP = &H0&            ' \Desktop
    CSIDL_PROGRAMS = &H2&           ' \Start Menu\Programs
    CSIDL_CONTROLS = &H3&           ' No Path
    CSIDL_PRINTERS = &H4&           ' No Path
    CSIDL_PERSONAL = &H5&           ' \Personal
    CSIDL_FAVORITES = &H6&          ' \Favorites
    CSIDL_STARTUP = &H7&            ' \Start Menu\Programs\Startup
    CSIDL_RECENT = &H8&             ' \Recent
    CSIDL_SENDTO = &H9&             ' \SendTo
    CSIDL_BITBUCKET = &HA&          ' No Path
    CSIDL_STARTMENU = &HB&          ' \Start Menu
    CSIDL_DESKTOPDIRECTORY = &H10&  ' \Desktop
    CSIDL_DRIVES = &H11&            ' No Path
    CSIDL_NETWORK = &H12&           ' No Path
    CSIDL_NETHOOD = &H13&           ' \NetHood
    CSIDL_FONTS = &H14&             ' \fonts
    CSIDL_TEMPLATES = &H15&         ' \ShellNew
    CSIDL_COMMON_STARTMENU = &H16&  ' ..\WinNT\profiles\All Users\Start Menu
    CSIDL_COMMON_PROGRAMS = &H17&   ' ..\WinNT\profiles\All Users\Start Menu\Programs
    CSIDL_COMMON_STARTUP = &H18&    ' ..\WinNT\profiles\All Users\Start Menu\Programs\Startup
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19& '..\WinNT\profiles\All Users\Desktop
    CSIDL_APPDATA = &H1A&           ' ..\WinNT\profiles\username\Application Data
    CSIDL_PRINTHOOD = &H1B&         ' ..\WinNT\profiles\username\PrintHood
End Enum

Public Enum SHOWCMDFLAGS
    SHOWNORMAL = 5
    SHOWMAXIMIZE = 3
    SHOWMINIMIZE = 7
End Enum

Public Const MAX_PATH = 255

Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hwndOwner As Long, ByVal nFolder As Integer, ppidl As Long) As Long
Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal szPath As String) As Long

Public Function fCreateShellLink(sLnkFile As String, sExeFile As String, sWorkDir As String, _
       sExeArgs As String, sIconFile As String, lIconIdx As Long, ShowCmd As SHOWCMDFLAGS) As Long

    Dim cShellLink   As ShellLinkA   ' An explorer IShellLinkA(Win 9x/Win NT) instance
    Dim cPersistFile As IPersistFile ' An explorer IPersistFile instance
    
    If (sLnkFile = "") Or (sExeFile = "") Then
        Exit Function
    End If

    On Error GoTo fCreateShellLinkError
    Set cShellLink = New ShellLinkA   'Create new IShellLink interface
    Set cPersistFile = cShellLink     'Implement cShellLink's IPersistFile interface
    
    With cShellLink
        'Debug.Print "sLnkFile " & sLnkFile
        'Set command line exe name & path to new ShortCut.
        .SetPath sExeFile
        'Debug.Print "sExeFile " & sExeFile
        'Set working directory in shortcut
        If sWorkDir <> "" Then .SetWorkingDirectory sWorkDir
        'Debug.Print "sWorkDir " & sWorkDir
        'Add arguments to command line
        If sExeArgs <> "" Then .SetArguments sExeArgs
        'Debug.Print "sExeArgs " & sExeArgs
        'Set shortcut description
        .SetDescription "Access Shortcut" & vbNullChar
    '   If (LnkDesc <> "") Then .SetDescription pszName
        'Set shortcut icon location & index
        If sIconFile <> "" Then .SetIconLocation sIconFile, lIconIdx
        'Debug.Print "sIconFile " & sIconFile
        'Set shortcut's startup mode (min,max,normal)
        .SetShowCmd ShowCmd
    End With

    cShellLink.Resolve 0, SLR_UPDATE
    cPersistFile.Save StrConv(sLnkFile, vbUnicode), 0 'Unicode conversion that must be done!
    fCreateShellLink = True 'Return Success
fCreateShellLinkError:
    Set cPersistFile = Nothing
    Set cShellLink = Nothing
End Function

Public Function fGetSystemFolderPath(ByVal hWnd As Long, ByVal Id As Integer, sfPath As String) As Long

    Dim lReturn As Long
    Dim lPidl   As Long
    Dim lPath   As Long
    Dim sPath   As String
    
    sPath = Space$(MAX_PATH)
    lReturn = SHGetSpecialFolderLocation(hWnd, Id, lPidl)  ' Get lPidl for Id...
    If lReturn = 0 Then                                    ' If success is 0
        lReturn = SHGetPathFromIDList(lPidl, sPath)        '   Get Path from Item Id List
        If lReturn = 1 Then                                '   If success is 1
            sPath = Trim$(sPath)                           '     Fix path string
            lPath = Len(sPath)                             '     Get length of path
            If Asc(Right$(sPath, 1)) = 0 Then lPath = lPath - 1 'Adjust path length
            If lPath > 0 Then sfPath = Left$(sPath, lPath) '     Adjust path string variable
            fGetSystemFolderPath = True                    '     Return success
        End If
    End If
End Function

Public Function fGetShellLinkInfo(sLnkFile As String, sExeFile As String, sWorkDir As String, _
       sExeArgs As String, sIconFile As String, lIconIdx As Long, lShowCmd As Long) As Long

    Dim lPidl        As Long              ' Item id list
    Dim lHotKey      As Long              ' Hotkey to shortcut...
    Dim lBuffLen     As Long
    Dim sTemp        As String
    Dim sDescription As String
    Dim cShellLink   As ShellLinkA        ' An explorer IShellLink instance
    Dim cPersistFile As IPersistFile      ' An explorer IPersistFile instance
    Dim fd           As WIN32_FIND_DATA

    If sLnkFile = "" Then
        Exit Function
    End If

    Set cShellLink = New ShellLinkA       ' Create new IShellLink interface
    Set cPersistFile = cShellLink         ' Implement cShellLink's IPersistFile interface

    'Load Shortcut file...(must do this UNICODE hack!)
    On Error GoTo fGetShellLinkInfoError
    cPersistFile.Load StrConv(sLnkFile, vbUnicode), STGM_DIRECT

    With cShellLink
        'Get command line exe name & path of shortcut
        sExeFile = Space$(MAX_PATH)
        lBuffLen = Len(sExeFile)
        .GetPath sExeFile, lBuffLen, fd, SLGP_UNCPRIORITY
        sTemp = fd.cFileName  ' Not returned to calling function
        
        'Get working directory of shortcut
        sWorkDir = Space$(MAX_PATH)
        lBuffLen = Len(sWorkDir)
        .GetWorkingDirectory sWorkDir, lBuffLen
        
        'Get command line arguments of shortcut
        sExeArgs = Space$(MAX_PATH)
        lBuffLen = Len(sExeArgs)
        .GetArguments sExeArgs, lBuffLen
        
        'Get description of shortcut
        sDescription = Space$(MAX_PATH)
        lBuffLen = Len(sDescription)
        .GetDescription sDescription, lBuffLen ' Not returned to calling function
        
        'Get the HotKey for shortcut
        .GetHotkey lHotKey  ' Not returned to calling function
       
        'Get shortcut icon location & index
        sIconFile = Space$(MAX_PATH)
        lBuffLen = Len(sIconFile)
        .GetIconLocation sIconFile, lBuffLen, lIconIdx
        
        'Get Item ID List...
        .GetIDList lPidl ' Not returned to calling function
                
        'Set shortcut's startup mode (min,max,normal)
        .GetShowCmd lShowCmd
    End With
    fGetShellLinkInfo = True

fGetShellLinkInfoError:
    Set cPersistFile = Nothing
    Set cShellLink = Nothing
End Function



