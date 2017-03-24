VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'LibData
'=======
'
'This class is used to parse, process, and store Reference
'and Object strings from a VB 6.0 project file.
'
'Usage:
'
'   Create and instance and call the LoadRefByProjFileObject() or
'   LoadRefByProjFileRef() method.  Then the properties of this
'   class can be interrogated to retrieve the attributes of
'   the subject ActiveX library.
'
'Dependencies:
'
'   o Uses the VB Type Library Information DLL TLBINF32.DLL
'
'   o The internal class ClsData.
'

'----- Public Data -----

'----- Private Data -----

Private m_ClassData() As ClsData
Private m_ClassLast As Long
Private m_Description As String
Private m_FileLocation As String
Private m_Flags As String
Private m_Included As Boolean
Private m_IncludedNever As Boolean
Private m_IsCOM As Boolean
Private m_LCID As Long
Private m_LIBID As String
Private m_MajorVersion As Integer
Private m_MinorVersion As Integer
Private m_Name As String
Private m_Source As String

'----- Private Methods -----

Private Function LookupRegByCLSIDVersLCID() As Boolean
    'Return True if TypeLibInfo can't be obtained.
    Dim tliLib As TLI.TypeLibInfo
    
    On Error Resume Next
    Set tliLib = TypeLibInfoFromRegistry(m_LIBID, -1, -1, -1)
    With tliLib
        m_MajorVersion = .MajorVersion
        m_MinorVersion = .MinorVersion
        m_LCID = .LCID
    End With
    If Err.Number = 0 Then
        LookupRegByCLSIDVersLCID = LookupRegByTLibInfo(tliLib)
        Set tliLib = Nothing
    Else
        LookupRegByCLSIDVersLCID = True
    End If
End Function

Private Function LookupRegByFileName(ByVal LibFileName As String) As Boolean
    'Returns True if this is a Standard or unregistered DLL/OCX.
    Dim strTLibFileName As String
    Dim tliLib As TLI.TypeLibInfo
    Dim blnFoundTypeLib As Boolean
    
    strTLibFileName = ExpandEnv(LibFileName)
    Do
        On Error Resume Next
        Set tliLib = TypeLibInfoFromFile(strTLibFileName)
        If Err.Number = 0 Then
            On Error GoTo 0
            blnFoundTypeLib = True
        Else
            On Error GoTo 0
            Select Case UCase$(Right$(strTLibFileName, 3))
                Case "DLL", "OCX"
                    strTLibFileName = Left$(strTLibFileName, Len(strTLibFileName) - 3) & "TLB"
                
                Case "TLB"
                    strTLibFileName = Left$(strTLibFileName, Len(strTLibFileName) - 3) & "OLB"
                
                Case Else
                    'Exhausted possibilities.
                    Exit Do
            End Select
        End If
    Loop Until blnFoundTypeLib
    
    If blnFoundTypeLib Then
        'COM DLL/OCX.
        m_IsCOM = True
        On Error GoTo 0
        With tliLib
            m_LIBID = .GUID
            m_MajorVersion = .MajorVersion
            m_MinorVersion = .MinorVersion
            m_LCID = .LCID
        End With
        
        If LookupRegByTLibInfo(tliLib) Then
            'Failed.
            m_Included = False
        Else
            'Recheck using actual file location.
            m_IncludedNever = ExcludedLib(LibFileName)
            m_Included = Not m_IncludedNever
        End If
        
        Set tliLib = Nothing
    Else
        'Standard or unregistered DLL/OCX.  Build what we can.
        On Error GoTo 0
        m_Description = "<Standard or unregistered DLL/OCX>"
        m_FileLocation = ExpandEnv(LibFileName)
        m_LIBID = "n/a"
        m_Name = SimpleFileName(LibFileName)
        m_Name = Left$(m_Name, InStr(m_Name, ".") - 1)
        m_IncludedNever = ExcludedLib(m_FileLocation)
        m_Included = Not m_IncludedNever
        m_ClassLast = -1
        
        LookupRegByFileName = True
    End If
End Function

Private Function LookupRegByTLibInfo(ByVal LibInfo As TLI.TypeLibInfo) As Boolean
    Dim blnBadFileLocation As Boolean
    Dim tiClass As TLI.TypeInfo
    Dim cdClass As ClsData
    Dim intAttrStrs As Integer
    Dim strAttrStrs() As String
    Dim strMiscStatusAttribs As String
    
    m_IsCOM = True
    With LibInfo
        m_Name = .Name
        intAttrStrs = .AttributeStrings(strAttrStrs)
        m_Flags = Join(strAttrStrs, ",") 'TLI CHM suggests multiples though MSDN says only one.
        m_Description = .HelpString
        
        '.ContainingFile can be unreliable if we scan ourself Reg-Free!
        m_FileLocation = .ContainingFile
        If Len(m_FileLocation) = 0 Then
            blnBadFileLocation = True
        Else
            m_FileLocation = ExpandEnv(m_FileLocation)
            blnBadFileLocation = Not FilePresent(m_FileLocation)
        End If
        If blnBadFileLocation Then
            m_FileLocation = GetRegistryValue(HKEY_CLASSES_ROOT, _
                                              "Typelib\" _
                                            & .GUID _
                                            & "\" _
                                            & VersionHex() _
                                            & "\" _
                                            & Hex$(.LCID) _
                                            & "\win32", _
                                              "")
            If StrPtr(m_FileLocation) = 0 Then
                LookupRegByTLibInfo = True 'Failed.
                frmLog.Log "*** Cannot resolve library file name for this dependency ***"
            Else
                m_FileLocation = ExpandEnv(m_FileLocation)
            End If
        End If
        
        m_IncludedNever = ExcludedLib(m_FileLocation)
        m_Included = Not m_IncludedNever
        m_ClassLast = -1
        For Each tiClass In .CoClasses
            If (tiClass.AttributeMask And TYPEFLAG_FHIDDEN) = 0 Then
                Set cdClass = New ClsData
                With cdClass
                    .CLSID = tiClass.GUID
                    .Description = tiClass.HelpString
                    .Name = tiClass.Name
                    .ProgID = GetRegistryValue(HKEY_CLASSES_ROOT, _
                                               "CLSID\" _
                                             & .CLSID _
                                             & "\ProgID", _
                                               "")
                    .ThreadingModel = GetRegistryValue(HKEY_CLASSES_ROOT, _
                                                      "CLSID\" _
                                                    & .CLSID _
                                                    & "\InprocServer32", _
                                                      "ThreadingModel")
                    If Len(.ThreadingModel) > 0 Then
                        'This class both exists in the Registry and has a non-null
                        'ThreadingModel.
                        .MiscStatusError = GetMiscStatusAttribs(.CLSID, strMiscStatusAttribs)
                        .MiscStatusAttributes = strMiscStatusAttribs
                    End If
                End With
                
                m_ClassLast = m_ClassLast + 1
                ReDim Preserve m_ClassData(m_ClassLast)
                Set m_ClassData(m_ClassLast) = cdClass
            End If
        Next
    End With
End Function

'----- Properties -----

Public Property Get Class(ByVal Index As Long) As ClsData
    If 0 <= Index And Index <= UBound(m_ClassData) Then
        Set Class = m_ClassData(Index)
    Else
        Err.Raise &H80045A01, "Class index out of range", "LibData"
    End If
End Property

Public Property Get Count() As Long
    Count = m_ClassLast + 1
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Get FileLocation() As String
    FileLocation = m_FileLocation
End Property

Public Property Get Flags() As String
    Flags = m_Flags
End Property

Public Property Get Included() As Boolean
    Included = m_Included
End Property

Friend Property Let Included(ByVal Included As Boolean)
    m_Included = Included
End Property

Public Property Get IncludedNever() As Boolean
    IncludedNever = m_IncludedNever
End Property

Public Property Get IsCOM() As Boolean
    IsCOM = m_IsCOM
End Property

Public Property Get LIBID() As String
    LIBID = m_LIBID
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Source() As String
    Source = m_Source
End Property

Public Property Get Version() As String
    Version = CStr(m_MajorVersion) & "." & CStr(m_MinorVersion)
End Property

Private Property Get VersionHex() As String
    VersionHex = LCase$(Hex$(m_MajorVersion) & "." & Hex$(m_MinorVersion))
End Property

'----- Public Methods -----

Public Function LoadManRefByFileName(ByVal LibFileName As String) As Boolean
    'Returns True if this is a Standard or unregistered DLL/OCX.
    m_Source = "MAN"
    LoadManRefByFileName = LookupRegByFileName(LibFileName)
End Function

Public Function LoadDepRefByFileName(ByVal LibFileName As String) As Boolean
    'Returns True if this is a Standard or unregistered DLL/OCX.
    m_Source = "DEP"
    LoadDepRefByFileName = LookupRegByFileName(LibFileName)
End Function

Public Function LoadRefByProjFileObject(ByVal ObjectString As String) As Boolean
    'String format:
    '
    '   <LIBID>#<Maj>.<Min>#<LCID>; <OCXName>
    '
    '   We now disregard this source for MajorVersion, MinorVersion, and LCID.
    
    Dim strParts() As String
    Dim strVersionParts() As String
    
    strVersionParts = Split(ObjectString, ";", 2)
    strParts = Split(strVersionParts(0), "#", 3)
    m_LIBID = strParts(0)
    m_Source = "VBP"
    
    LoadRefByProjFileObject = LookupRegByCLSIDVersLCID()
End Function

Public Function LoadRefByProjFileRef(ByVal ReferenceString As String) As Boolean
    'String format:
    '
    '   *\G<LIBID>#<Maj>.<Min>#<LCID>#<RelFileLoc>#<Desc>
    '
    '   We now disregard this source for MajorVersion, MinorVersion, and LCID.
    
    Dim strParts() As String
    Dim strVersionParts() As String
    
    strParts = Split(Mid$(ReferenceString, 4), "#", 5)
    m_LIBID = strParts(0)
    m_Source = "VBP"
    
    If UCase$(Right$(strParts(3), 4)) = ".TLB" Then
        m_IncludedNever = True
        m_Included = Not m_IncludedNever
        m_Name = SimpleFileName(strParts(3))
        m_Name = Left$(m_Name, InStr(m_Name, ".") - 1)
        m_Description = "<type library>"
    Else
        LoadRefByProjFileRef = LookupRegByCLSIDVersLCID()
    End If
End Function

