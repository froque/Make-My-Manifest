VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppEx"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'AppEx
'=====
'
'This is an internal Class with a predeclared global instance.
'It can be used throughout your programs without creating an
'explicit instance, much like using the intrinsic App object
'in VB.
'
'AppEx returns OS Version information, special folder paths,
'and other information.  It can be useful in writing programs
'that conform to the Windows standards that Vista now enforces.
'
'For Windows 2000 and later, some functions require Windows XP
'or later.
'

Private Const CSIDL_APPDATA = &H1A                 '...\<user name>\Application Data
Private Const CSIDL_COMMON_APPDATA = &H23          '...\All Users\Application Data
Private Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19 '...\All Users\Desktop
Private Const CSIDL_COMMON_DOCUMENTS = &H2E        '...\All Users\Documents
Private Const CSIDL_COMMON_PROGRAMS = &H17         '...\All Users\Start Menu\Programs
Private Const CSIDL_COMMON_STARTMENU = &H16        '...\All Users\Start Menu
Private Const CSIDL_COMMON_STARTUP = &H18          '...\All Users\Start Menu\Programs\Startup
Private Const CSIDL_DESKTOPDIRECTORY = &H10        '...\<user name>\Desktop
Private Const CSIDL_LOCAL_APPDATA = &H1C           '...\<user name>\Local Settings\Application Data (non roaming)
Private Const CSIDL_MYPICTURES = &H27              '...\<user name>\My Documents\My Pictures
Private Const CSIDL_PERSONAL = &H5                 '...\<user name>\My Documents
Private Const CSIDL_PROFILE = &H28                 '...\<user name>
Private Const CSIDL_PROGRAMS = &H2                 '...\<user name>\Start Menu\Programs
Private Const CSIDL_PROGRAM_FILES = &H26           '...\Program Files
Private Const CSIDL_PROGRAM_FILES_COMMON = &H2B    '...\Program Files\Common
Private Const CSIDL_STARTMENU = &HB                '...\<user name>\Start Menu
Private Const CSIDL_STARTUP = &H7                  '...\<user name>\Start Menu\Programs\Startup
Private Const CSIDL_SYSTEM = &H25                  'GetSystemDirectory()
Private Const CSIDL_WINDOWS = &H24                 'GetWindowsDirectory()

Private Const CSIDL_FLAG_CREATE = &H8000&          'combine with other CSIDL value to force folder creation

Private Const S_OK = 0                             'Good result.
Private Const S_FALSE = 1                          'Valid CSIDL but folder doesn't exist.
Private Const E_INVALIDARG = &H80070057            'Bad CSIDL for requested function.

Private Const MAX_COMPUTERNAME_LENGTH = 31
Private Const MAX_PATH = 260

Public Enum AppExPaths
    aipAppData = CSIDL_APPDATA
    aipAppDataCommon = CSIDL_COMMON_APPDATA
    aipAppDataLocal = CSIDL_LOCAL_APPDATA
    aipDesktop = CSIDL_DESKTOPDIRECTORY
    aipDesktopCommon = CSIDL_COMMON_DESKTOPDIRECTORY
    aipMyDocuments = CSIDL_PERSONAL
    aipMyPictures = CSIDL_MYPICTURES
    aipMyProfile = CSIDL_PROFILE
    aipProgramFiles = CSIDL_PROGRAM_FILES
    aipProgramFilesCommon = CSIDL_PROGRAM_FILES_COMMON
    aipStartMenu = CSIDL_STARTMENU
    aipStartMenuCommon = CSIDL_COMMON_STARTMENU
    aipStartPrograms = CSIDL_PROGRAMS
    aipStartProgramsCommon = CSIDL_COMMON_PROGRAMS
    aipStartStartup = CSIDL_STARTUP
    aipStartStartupCommon = CSIDL_COMMON_STARTUP
    aipSystem = CSIDL_SYSTEM
    aipWindows = CSIDL_WINDOWS
End Enum

Public Enum AppExPathTypes
    aitCurrent = 0 'Current path after possible user renaming!
    aitDefault = 1 'As originally defined by Windows installation!
End Enum

Public Enum AppExOSVersion
    aioBuildNumber
    aioHomeEdition
    aioHomeServer
    aioPlatformId
    aioProductType
    aioServicePackString
    aioServicePackVersion
    aioVersion
    aioWorkstation
End Enum

Public Enum AppExPlatformIds
    aipidWIN32s = 0
    aipidWIN32_WINDOWS = 1
    aipisWIN32_NT = 2
End Enum

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
        'Operating System     Value
        'Windows 3.1            3
        'Windows 95             4
        'Windows 98             4
        'Windows Me             4
        'Windows NT 3.51        3
        'Windows NT 4.0         4
        'Windows 2000           5
        'Windows XP             5
        'Windows .Net Server    5
        'Windows 2003 Server    5
        'Windows 2003 R2 Server 5
        'Windows Vista          6
        'Windows 2008 Server    6
    dwMinorVersion As Long
        'Operating System     Value
        'Windows 3.1            1
        'Windows 95             0
        'Windows 98             10
        'Windows Me             90
        'Windows NT 3.51        51
        'Windows NT 4.0         0
        'Windows 2000           0
        'Windows XP             1
        'Windows .Net Server    1
        'Windows 2003 Server    2
        'Windows 2003 R2 Server 2
        'Windows Vista          0
        'Windows 2008 Server    0
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    
    'Extended information (optional), i.e. OSVERSIONINFOEX:
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
        'Operating System     Value
        'NT Workstation         1
        'NT Domain Controller   2
        'NT Server              3
    wReserved As Byte
End Type

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" ( _
    ByVal lpBuffer As String, _
    nSize As Long) As Long

Private Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" ( _
    ByVal lpBuffer As String, _
    nSize As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" _
    Alias "GetVersionExA" ( _
    lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function InitCommonControlsAPI Lib "comctl32" _
    Alias "InitCommonControls" ( _
    ) As Boolean

Private Declare Function SHGetFolderPath Lib "shell32" _
    Alias "SHGetFolderPathA" ( _
    ByVal hWndOwner As Long, _
    ByVal nFolder As Long, _
    ByVal hToken As Long, _
    ByVal dwFlags As Long, _
    ByVal lpszPath As String) As Long

Private Declare Function SHGetFolderPathAndSubDir Lib "shell32" _
    Alias "SHGetFolderPathAndSubDirA" ( _
    ByVal hWnd As Long, _
    ByVal csidl As Long, _
    ByVal hToken As Long, _
    ByVal dwFlags As Long, _
    ByVal pszSubDir As String, _
    ByVal pszPath As String) As Long
    'Only Windows XP and later (Shell32.dll 5.60 or later).

Private lngGetVersionExError As Long
Private osvResult As OSVERSIONINFO
Private strDecimalChar As String

Public Sub InitCommonControls()
Attribute InitCommonControls.VB_Description = "Initialize Common Controls library to permit XP Styles manifest to be applied"
    InitCommonControlsAPI
End Sub

Public Function ComputerName() As String
Attribute ComputerName.VB_Description = "NetBIOS name of this computer"
    Dim lngLen As Long
    Dim lngResult As Long
    Dim strComputerName As String
    
    strComputerName = String$(MAX_COMPUTERNAME_LENGTH, 0)
    lngLen = 255
    lngResult = GetComputerName(strComputerName, lngLen)
    If lngResult <> 0 Then
        ComputerName = Left$(strComputerName, lngLen)
    Else
        ComputerName = ""
    End If
End Function

Public Function OSVersion(ByVal VersionData As AppExOSVersion) As Variant
Attribute OSVersion.VB_Description = "Retrieve OS version data item requested"
    Const VER_NT_WORKSTATION = &H1&
    Const VER_SUITE_PERSONAL = &H200&
    Const VER_SUITE_WH_SERVER = &H8000&

    If lngGetVersionExError <> 0 Then
        Err.Raise vbObjectError + &HB106&, _
                  "AppEx.OSVersion", _
                  "System error retrieving information " & CStr(lngGetVersionExError)
    Else
        With osvResult
            Select Case VersionData
                Case aioBuildNumber
                    OSVersion = .dwBuildNumber
                
                Case aioHomeEdition
                    OSVersion = CBool((.dwPlatformId And VER_SUITE_PERSONAL) <> 0)
                
                Case aioHomeServer
                    OSVersion = CBool((.dwPlatformId And VER_SUITE_WH_SERVER) <> 0)
                
                Case aioPlatformId
                    OSVersion = .dwPlatformId
                
                Case aioProductType
                    OSVersion = .wProductType
                
                Case aioServicePackString
                    OSVersion = Left$(.szCSDVersion, InStr(.szCSDVersion, vbNullChar) - 1)
                
                Case aioServicePackVersion
                    OSVersion = CSng(CStr(.wServicePackMajor) & strDecimalChar & CStr(.wServicePackMinor))
                
                Case aioVersion
                    OSVersion = CSng(CStr(.dwMajorVersion) & strDecimalChar & CStr(.dwMinorVersion))
                
                Case aioWorkstation
                    OSVersion = CBool((.wProductType And VER_NT_WORKSTATION) <> 0)
            End Select
        End With
    End If
End Function

Public Function Path( _
    ByVal PathId As AppExPaths, _
    Optional ByVal ForceCreation As Boolean = False, _
    Optional ByVal PathType As AppExPathTypes = aitCurrent) As String
Attribute Path.VB_Description = "Retrieves special folder path or creates it"
    
    Dim strPath As String
    Dim lngResult As Long
    
    strPath = Space$(MAX_PATH)
    If ForceCreation Then PathId = PathId Or CSIDL_FLAG_CREATE
    lngResult = SHGetFolderPath(0&, PathId, 0&, PathType, strPath)
    Select Case lngResult
        Case S_OK
            Path = Left$(strPath, InStr(strPath, vbNullChar) - 1)
        Case S_FALSE
            Err.Raise vbObjectError + &HB112&, _
                      "AppEx.Path", _
                      "Valid PathId but folder doesn't exist"
        Case E_INVALIDARG
            Err.Raise vbObjectError + &HB114&, _
                      "AppEx.Path", _
                      "Invalid PathId for requested function"
        Case Else
            Err.Raise vbObjectError + &HB116&, _
                      "AppEx.Path", _
                      "System error " & CStr(lngResult)
    End Select
End Function

Function PathAndSubDir( _
    ByVal PathId As AppExPaths, _
    ByVal SubDir As String, _
    Optional ByVal ForceCreation As Boolean = False, _
    Optional ByVal PathType As AppExPathTypes = aitCurrent) As String
Attribute PathAndSubDir.VB_Description = "Retrieves special folder path and subdir or creates it"

    Dim strPath As String
    Dim lngResult As Long

    If OSVersion(aioVersion) < 5.1 Then
        Err.Raise vbObjectError + &HB120&, _
                  "AppEx.PathAndSubDir", _
                  "Requires Windows XP or later"
    Else
        SubDir = SubDir & String$(MAX_PATH - Len(SubDir), vbNullChar)
        strPath = Space$(MAX_PATH)
        If ForceCreation Then PathId = PathId Or CSIDL_FLAG_CREATE
        lngResult = SHGetFolderPathAndSubDir(0&, PathId, 0&, PathType, SubDir, strPath)
        Select Case lngResult
            Case S_OK
                PathAndSubDir = Left$(strPath, InStr(strPath, vbNullChar) - 1)
            Case S_FALSE
                Err.Raise vbObjectError + &HB122&, _
                          "AppEx.PathAndSubDir", _
                          "Valid PathId but folder doesn't exist"
            Case E_INVALIDARG
                Err.Raise vbObjectError + &HB124&, _
                          "AppEx.PathAndSubDir", _
                          "Invalid PathId for requested function"
            Case Else
                Err.Raise vbObjectError + &HB126&, _
                          "AppEx.PathAndSubDir", _
                          "System error " & CStr(lngResult)
        End Select
    End If
End Function

Public Function UserName() As String
Attribute UserName.VB_Description = "Logon user name of this session"
    Dim lngLen As Long
    Dim lngResult As Long
    Dim strUserName As String
    
    strUserName = String$(MAX_PATH, 0)
    lngLen = 255
    lngResult = GetUserName(strUserName, lngLen)
    If lngResult <> 0 Then
        UserName = Left$(strUserName, lngLen - 1)
    Else
        UserName = ""
    End If
End Function

Private Sub Class_Initialize()
    strDecimalChar = Mid$(CStr(1 / 2), 2, 1)
    With osvResult
        .dwOSVersionInfoSize = Len(osvResult)
        .szCSDVersion = ""
        If GetVersionEx(osvResult) = 0 Then
            lngGetVersionExError = Err.LastDllError
        End If
    End With
End Sub
