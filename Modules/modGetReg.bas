Attribute VB_Name = "modGetReg"
Option Explicit

'----- Private Data -----

Private Const KEY_READ = &H20019  '(    (READ_CONTROL Or KEY_QUERY_VALUE
                                  '      Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY)
                                  ' And (Not SYNCHRONIZE)
                                  ')

Private Const ERROR_SUCCESS = 0
Private Const ERROR_FILE_NOT_FOUND = 2
Private Const ERROR_MORE_DATA = 234

'----- Private Declarations -----

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long
    
Private Declare Function RegCloseKey Lib "advapi32" ( _
    ByVal hKey As Long) As Long
    
Private Declare Function RegQueryValueEx Lib "advapi32" _
    Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As String, _
    lpcbData As Long) As Long

'----- Public Data -----

Public Const HKEY_CLASSES_ROOT = &H80000000

'----- Public Methods -----

Public Function GetRegistryValue( _
    ByVal hKey As Long, _
    ByVal KeyName As String, _
    ByVal ValueName As String) As String
    'Read a REG_SZ value.  If ValueName is empty ("") retrieve
    'the Key's default value.
    '
    'Returns:
    '   Value if present.
    '   Empty ("") if not present (or empty).
    '   vbNullString if Key doesn't exist in the Registry.
    
    Dim hKeyName As Long
    Dim lngRes As Long
    Dim lngType As Long
    Dim lngValueLen As Long
    Dim strValue As String
    
    'Open the key, exit with vbNullString if not found.
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, hKeyName) Then
        GetRegistryValue = vbNullString
    Else
        'Query ValueName.
        lngValueLen = 1024
        strValue = String$(lngValueLen, 0)
        lngRes = RegQueryValueEx(hKeyName, ValueName, 0, lngType, ByVal strValue, lngValueLen)
        
        'Check for bytValue() too short, try again if required.
        If lngRes = ERROR_MORE_DATA Then
            strValue = String$(lngValueLen, 0)
            lngRes = RegQueryValueEx(hKeyName, ValueName, 0, lngType, ByVal strValue, lngValueLen)
        End If
        
        Select Case lngRes
            Case ERROR_SUCCESS
                GetRegistryValue = Left$(strValue, lngValueLen - 1)
                
            Case ERROR_FILE_NOT_FOUND
                GetRegistryValue = ""
                
            Case Else
                Err.Raise &H80047401, _
                          "GetRegistryValue", _
                          "RegQueryValueEx returned error " & CStr(lngRes)
        End Select
    End If
    
    'Close the key.
    RegCloseKey hKeyName
End Function


