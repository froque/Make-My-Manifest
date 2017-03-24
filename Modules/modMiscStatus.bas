Attribute VB_Name = "modMiscStatus"
Option Explicit
'
'Per Microsoft KB 828629 (rev 1.1):
'   FIX: Windows side-by-side execution is not supported for Visual
'        Basic 6.0 ActiveX controls.
'
'   SYMPTOMS
'        With Microsoft Windows XP and later versions, you can run
'        Microsoft Component Object Model (COM) DLL modules in
'        side-by-side (SxS) mode. With SxS, different versions of a
'        COM DLL to co-exist in the same computer environment without
'        conflict. This behavior occurs by using .manifest files that
'        specify how a program may bind to a particular COM DLL.
'        Microsoft Visual Basic 6.0 ActiveX controls are essentially
'        COM DLL modules with .ocx file name extensions. If you try to
'        configure these modules for SxS operation in Windows XP, you
'        receive the following error message:
'
'        Runtime Error '336' Component not correctly registered.
'
'   CAUSE
'        This problem occurs because both the Windows and Visual Basic
'        runtimes do not support configuring SxS execution of Visual
'        Basic 6.0 ActiveX controls.
'
'   RESOLUTION
'        To fully resolve this problem you must have both of the
'        following:
'
'           • The version of the Visual Basic runtime that is included
'             with Visual Basic 6 Service Pack 6 (SP6).
'
'           • Windows XP Service Pack 2 (SP2).
'
'Thus programs applying this remedy may only run on Windows XP SP2 (or
'later?) with the SP6 VB6 runtime.
'

'----- Private Data -----

'API error codes.
Private Const S_OK = 0
Private Const REGDB_E_CLASSNOTREG = &H80040154
Private Const REGDB_E_READREGDB = &H80040150
Private Const REGDB_E_KEYMISSING = &H80040152
Private Const OLE_E_REGDB_KEY = REGDB_E_KEYMISSING
Private Const E_OUTOFMEMORY = &H8007000E

'OLE_Misc attribute strings for CoClass nodes of VB 6.0 ActiveX controls in manifests.
Private Const OLEMISC_ATTRIBS As String = _
    "miscStatus%miscStatusIcon%miscStatusThumbnail%miscStatusDocPrint%" _
  & "miscStatusContent"
Private Const OLEMISC_FLAGS As String = _
    "recomposeonresize%onlyiconic%insertnotreplace%static%cantlinkinside%" _
  & "canlinkbyole1%islinkobject%insideout%activatewhenvisible%" _
  & "renderingisdeviceindependent%invisibleatruntime%alwaysrun%actslikebutton%" _
  & "actslikelabel%nouiactivate%alignable%simpleframe%setclientsitefirst%" _
  & "imemode%ignoreactivatewhenvisible%wantstomenumerge%supportsmultilevelundo"

Private Enum tagDVASPECT
    DVASPECT_DEFAULT = 0
    DVASPECT_CONTENT = 1
    DVASPECT_THUMBNAIL = 2
    DVASPECT_ICON = 4
    DVASPECT_DOCPRINT = 8
End Enum

Private Enum tagOLEMISC 'Bitwise.
    OLEMISC_RECOMPOSEONRESIZE = 1
    OLEMISC_ONLYICONIC = 2
    OLEMISC_INSERTNOTREPLACE = 4
    OLEMISC_STATIC = 8
    OLEMISC_CANTLINKINSIDE = 16
    OLEMISC_CANLINKBYOLE1 = 32
    OLEMISC_ISLINKOBJECT = 64
    OLEMISC_INSIDEOUT = 128
    OLEMISC_ACTIVATEWHENVISIBLE = 256
    OLEMISC_RENDERINGISDEVICEINDEPENDENT = 512
    OLEMISC_INVISIBLEATRUNTIME = 1024
    OLEMISC_ALWAYSRUN = 2048
    OLEMISC_ACTSLIKEBUTTON = 4096
    OLEMISC_ACTSLIKELABEL = 8192
    OLEMISC_NOUIACTIVATE = 16384
    OLEMISC_ALIGNABLE = 32768
    OLEMISC_SIMPLEFRAME = 65536
    OLEMISC_SETCLIENTSITEFIRST = 131072
    OLEMISC_IMEMODE = 262144
    OLEMISC_IGNOREACTIVATEWHENVISIBLE = 524288
    OLEMISC_WANTSTOMENUMERGE = 1048576
    OLEMISC_SUPPORTSMULTILEVELUNDO = 2097152
End Enum

'----- Private Declares -----

Private Declare Function CLSIDFromString Lib "ole32" ( _
    ByVal lpsz As Long, _
    ByRef pclsid As Any) As Long

Private Declare Function OleRegGetMiscStatus Lib "ole32" ( _
    ByRef CLSID As Any, _
    ByVal dwAspect As tagDVASPECT, _
    ByRef pdwStatus As Long) As Long

'----- Public Methods -----

Public Function GetMiscStatusAttribs( _
    ByVal CLSID As String, _
    ByRef AttribsString As String) As Long
    'Get CLSID's miscStatus settings as a manifest attribute string.
    '
    'Call GetMiscStatusAttrib() for each DVASPECT of CLSID obtaining
    'OLE_Misc XML attribute strings (or until an error), accumlating
    'the values in AttribsString (or returning an error description).
    Dim GUID(15) As Byte
    Dim intAspect As Integer
    Dim strAttrib As String
    
    AttribsString = ""
    CLSIDFromString StrPtr(CLSID), GUID(0)
    For intAspect = 1 To 5
        GetMiscStatusAttribs = _
            GetMiscStatusAttrib(GUID, Choose(intAspect, _
                                             DVASPECT_DEFAULT, _
                                             DVASPECT_CONTENT, _
                                             DVASPECT_THUMBNAIL, _
                                             DVASPECT_ICON, _
                                             DVASPECT_DOCPRINT), strAttrib)
        If GetMiscStatusAttribs = S_OK Then
            If Len(strAttrib) > 0 Then
                'Accumulate AttribsString value.
                AttribsString = AttribsString & " " & strAttrib
            End If
        Else
            'Build error description string and exit immediately.
            AttribsString = strAttrib 'Capture attribute name.
            Select Case GetMiscStatusAttribs
                Case REGDB_E_CLASSNOTREG
                    strAttrib = "REGDB_E_CLASSNOTREG"
                Case REGDB_E_READREGDB
                    strAttrib = "REGDB_E_READREGDB"
                Case OLE_E_REGDB_KEY
                    strAttrib = "OLE_E_REGDB_KEY"
                Case E_OUTOFMEMORY
                    strAttrib = "E_OUTOFMEMORY"
            End Select
            AttribsString = AttribsString & " error " & strAttrib
            
            Exit Function
        End If
    Next
End Function

'----- Private Methods -----

Private Function GetMiscStatusAttrib( _
    ByRef GUID() As Byte, _
    ByVal DVASPECT As tagDVASPECT, _
    ByRef AttribString As String) As Long
    'Call OleRegGetMiscStatus() for GUID and DVASPECT and return
    'an error indication or a constructed manifest AttrbiString
    'value (or attribute name on an error).
    '
    'Returns:
    '   Good, no flags:
    '       Return value 0 and empty AttribString.
    '   Good, flags:
    '       Return value 0 and AttribString value of:
    '           <OLEMISC_ATTRIB name>="<comma-separated list of OLEMISC_FLAG values"
    '   Error:
    '       Return value of error code and AttribString value of:
    '           <OLEMISC_ATTRIB name>
    Dim lngDWord As Long
    Dim intBit As Integer
    Dim lngStatus As Long
    Dim strFlagList As String
    
    lngDWord = DVASPECT
    intBit = -1
    NextBitNum lngDWord, intBit
    AttribString = Split(OLEMISC_ATTRIBS, "%")(intBit + 1)
    GetMiscStatusAttrib = OleRegGetMiscStatus(GUID(0), DVASPECT, lngStatus)
    If GetMiscStatusAttrib = S_OK Then
        intBit = -1
        NextBitNum lngStatus, intBit
        Do Until intBit < 0
            If intBit >= 0 Then
                If Len(strFlagList) > 0 Then
                    strFlagList = strFlagList & ","
                Else
                    strFlagList = strFlagList & "="""
                End If
                strFlagList = strFlagList & Split(OLEMISC_FLAGS, "%")(intBit)
            End If
            NextBitNum lngStatus, intBit
        Loop
        If Len(strFlagList) > 0 Then
            AttribString = AttribString & strFlagList & """"
        Else
            AttribString = ""
        End If
    End If
End Function

Private Sub NextBitNum(ByRef DWord As Long, _
                       ByRef BitNum As Integer)
    'Returns with BitNum equal to number of next bit set in DWord.
    '
    'Alters both DWord and BitNum.  Initial call must have original
    'DWord image (or copy) and BitNum set to -1.
    '
    'Returns -1 in BitNum if no more bits set in DWord.
    Dim blnBitSet As Boolean
    
    Do While DWord <> 0
        BitNum = BitNum + 1
        blnBitSet = CBool(DWord And 1)
        DWord = ((DWord And &HFFFFFFFE) \ 2) And &H7FFFFFFF
        If blnBitSet Then
            Exit Sub
        End If
    Loop
    BitNum = -1
End Sub
