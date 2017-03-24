Attribute VB_Name = "modUtils"
Option Explicit

'----- Private Data -----

'Misc. API constants.
Private Const MAX_PATH = 260

'----- Private Declares -----

Private Declare Function ExpandEnvironmentStrings Lib "kernel32" _
    Alias "ExpandEnvironmentStringsA" ( _
    ByVal lpSrc As String, _
    ByVal lpDst As String, _
    ByVal nSize As Long) As Long
'@@@@ old old old
'Private Declare Function GetFullPathName Lib "kernel32" _
'    Alias "GetFullPathNameW" ( _
'    ByVal lpFileName As Long, _
'    ByVal nBufferLength As Long, _
'    ByVal lpBuffer As Long, _
'    ByVal lpFilePart As Long) As Long

Private Declare Function GetFullPathName Lib "kernel32" _
    Alias "GetFullPathNameA" ( _
    ByVal lpFileName As String, _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As String, _
    ByVal lpFilePart As String) As Long

Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
    
'----- Public Data -----

Public Enum GetLCIDSource
    glsFirst = 0
    glsNone = glsFirst
    glsUserDefault
    glsSysDefault
    glsLast = glsSysDefault
End Enum

'----- Public Methods -----

Public Function DepFileName(ByVal FQFileName As String) As String
    DepFileName = Left$(FQFileName, InStrRev(FQFileName, ".")) & "DEP"
End Function

Public Function DLLSearch( _
    ByVal DepPath As String, _
    ByVal ProjPath As String, _
    ByVal LibSimpleFileName As String) As String
    'Try to locate the specified DLL or OCX by approaching Windows'
    'runtime DLL search process, but biased toward development.
    '
    'Returns full path and filename or an empty String.
    
    DLLSearch = DepPath & LibSimpleFileName
    If FilePresent(DLLSearch) Then Exit Function
    
    DLLSearch = ProjPath & LibSimpleFileName
    If FilePresent(DLLSearch) Then Exit Function
    
    DLLSearch = GetSystemPath() & LibSimpleFileName
    If FilePresent(DLLSearch) Then Exit Function
    
    DLLSearch = GetWindowsPath() & LibSimpleFileName
    If FilePresent(DLLSearch) Then Exit Function
    
    DLLSearch = ""
End Function

Public Function DQ(ByVal QuotedText As String) As String
    DQ = Trim$(Mid$(QuotedText, 2, Len(QuotedText) - 2))
End Function

Public Function ExcludedLib(ByVal LibFile As String) As Boolean
    Dim strExt As String
    
    strExt = ExtOfFileName(LibFile)
    ExcludedLib = ExcludedListScan(LibFile) _
               Or strExt = "TLB" _
               Or strExt = "OLB"
End Function

Public Function ExpandEnv(ByVal Path As String) As String
    Dim lngNeeded As Long
    
    lngNeeded = ExpandEnvironmentStrings(Path, ExpandEnv, 0)
    If lngNeeded = 0 Then
        Err.Raise vbObjectError Or &H8900&, "ExpandEnv", _
                  "Kernel32 error &H" & Hex$(Err.LastDllError) & " on Path:" _
                & vbNewLine & """" & Path & """"
    End If
    ExpandEnv = Space$(lngNeeded)
    ExpandEnvironmentStrings Path, ExpandEnv, lngNeeded
    ExpandEnv = Left$(ExpandEnv, InStr(ExpandEnv, vbNullChar) - 1)
End Function

Public Function ExtOfFileName(ByVal FileName As String) As String
    ExtOfFileName = UCase$(Mid$(FileName, InStrRev(FileName, ".") + 1))
End Function

Public Function FilePresent(ByVal FileName As String) As Boolean
    FilePresent = Len(Dir$(FileName, vbNormal Or vbHidden Or vbSystem)) > 0
End Function

Public Function FormatXML(ByVal Value As String) As String
    Dim lngCharX As Long
    Dim lngChar As Long
    
    lngCharX = InStr(Value, vbNullChar)
    If lngCharX > 0 Then Value = Left$(Value, lngCharX - 1)
    FormatXML = Replace$(Value, "&", "&amp;")
    FormatXML = Replace$(FormatXML, """", "&quot;")
    FormatXML = Replace$(FormatXML, "'", "&apos;")
    FormatXML = Replace$(FormatXML, "<", "&lt;")
    FormatXML = Replace$(FormatXML, ">", "&gt;")
    
    For lngCharX = Len(FormatXML) To 1 Step -1
        lngChar = AscW(Mid$(FormatXML, lngCharX, 1))
        If &H20& > lngChar Or lngChar > &H7E& Then
            Mid$(FormatXML, lngCharX, 1) = "_"
        End If
    Next
End Function

Public Function GetDepLCID(ByVal Which As GetLCIDSource) As String
    Dim lngLCID As Long
    
    Select Case Which
        Case glsNone
            Exit Function
        
        Case glsUserDefault
            lngLCID = GetUserDefaultLCID()

        Case glsSysDefault
            lngLCID = GetSystemDefaultLCID()
    End Select
    
    GetDepLCID = " <" & Right$("000" & Hex$(lngLCID And &H3F&), 4) & ">"
End Function
'@@@@ old old old
'Public Function GetFullPath(ByVal Path As String)
'    Dim lngLen As Long
'
'    lngLen = GetFullPathName(StrPtr(Path), 0, StrPtr(GetFullPath), ByVal 0&)
'    GetFullPath = Space$(lngLen - 1)
'    GetFullPathName StrPtr(Path), lngLen, StrPtr(GetFullPath), ByVal 0&
'End Function

Public Function GetFullPath(ByVal Path As String) As String
    Dim lngNeeded As Long
    
    lngNeeded = GetFullPathName(Path, 0, GetFullPath, ByVal 0&)
    If lngNeeded = 0 Then
        Err.Raise vbObjectError Or &H8902&, "GetFullPath", _
                  "Kernel32 error &H" & Hex$(Err.LastDllError) & " on Path:" _
                & vbNewLine & """" & Path & """"
    End If
    GetFullPath = Space$(lngNeeded)
    GetFullPathName Path, lngNeeded, GetFullPath, ByVal 0&
    GetFullPath = Left$(GetFullPath, InStr(GetFullPath, vbNullChar) - 1)
End Function

Public Function GetProductSettingsPath() As String
    GetProductSettingsPath = App.Path & "\"
End Function

Public Function GetSystemPath() As String
    GetSystemPath = AppEx.Path(aipSystem) & "\"
End Function

Public Function GetWindowsPath() As String
    GetWindowsPath = AppEx.Path(aipWindows) & "\"
End Function

Public Sub IniDOMCloneSection( _
    ByVal SectionName As String, _
    ByRef FromDOM As IniDOM, _
    ByRef ToDOM As IniDOM)
    Dim ikKey As IniKey
    
    With FromDOM(SectionName)
        ToDOM.Sections.Add SectionName, .Comment, .Unrecognized
        For Each ikKey In .Keys
            With ikKey
                ToDOM(SectionName).Keys.Add .Name, _
                                            .Value, _
                                            .QuoteName, _
                                            .QuoteValue, _
                                            .Comment, _
                                            .Unrecognized
            End With
        Next
    End With
End Sub

Public Function IniDOMFromFile( _
    ByVal FileName As String, _
    Optional ByVal CreateNotPresent As Boolean = False) As IniDOM
    
    If Not FilePresent(FileName) Then
        If CreateNotPresent Then
            Set IniDOMFromFile = New IniDOM
        End If
        'Else exit returning Nothing.
    Else
        Dim stmIni As New ADODB.Stream
        
        With stmIni
            .Open
            .Type = adTypeText
            .Charset = "ascii"
            .LoadFromFile FileName
            Set IniDOMFromFile = New IniDOM
            IniDOMFromFile.Load stmIni
            .Close
        End With
    End If
End Function

Public Sub IniDOMToFile(ByRef DOM As IniDOM, ByVal FileName As String)
    Dim stmIni As New ADODB.Stream
    
    With stmIni
        .Open
        .Type = adTypeText
        .Charset = "ascii"
        DOM.Save stmIni
        .SaveToFile FileName, adSaveCreateOverWrite
        .Close
    End With
End Sub

Public Function PathOfFQFileName(ByVal FQFileName As String) As String
    PathOfFQFileName = Left$(FQFileName, InStrRev(FQFileName, "\"))
End Function

Public Function SimpleFileName(ByVal FQFileName As String) As String
    SimpleFileName = Mid$(FQFileName, InStrRev(FQFileName, "\") + 1)
End Function

Public Function TrimSlash(ByVal Path As String) As String
    If Right$(Path, 1) = "\" Then
        TrimSlash = Left$(Path, Len(Path) - 1)
    Else
        TrimSlash = Path
    End If
End Function

'----- Private Methods -----

Private Function ExcludedListScan(ByVal LibFile As String) As Boolean
    Dim iksExclusions As IniKeys
    Dim ikLib As IniKey
    
    LibFile = UCase$(SimpleFileName(LibFile))
    Set iksExclusions = G_Product_IniDOM!Exclusions.Keys
    For Each ikLib In iksExclusions
        'We use UCase$(.Name) here, as in VB6DEP.INI!
        If Len(ikLib.Name) > 0 Then
            'Valid Key, not a comment, etc.
            If LibFile Like UCase$(ikLib.Name) Then
                ExcludedListScan = True
                Exit For
            End If
        End If
    Next
End Function
