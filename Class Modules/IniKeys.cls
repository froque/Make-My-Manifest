VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Keys As Collection
Private m_intKeyComment As Integer

Public Enum IniKeysExceptions
    iksNoKey = &H80047D10
    iksDupKey = &H80047D12
End Enum

Public Function Add( _
    Optional ByVal Name As String = "", _
    Optional ByVal Value As String = "", _
    Optional ByVal QuoteName As Boolean = False, _
    Optional ByVal QuoteValue As Boolean = False, _
    Optional ByVal Comment As String = "", _
    Optional ByVal Unrecognized As Boolean = False) As IniKey
    Dim strKeyKey As String
    
    If Len(Name) = 0 Then
        'This is a Key-level comment in the INI file.
        strKeyKey = ";" & CStr(m_intKeyComment)
        m_intKeyComment = m_intKeyComment + 1
    Else
        strKeyKey = UCase$(Name)
    End If
    
    Set Add = New IniKey
    With Add
        'Assign these first, since assigning Name and Value may
        'Need to override the settings passed to us here.
        .QuoteName = QuoteName
        .QuoteValue = QuoteValue
        
        .Name = Name
        .Value = Value
        .Comment = Comment
        .Unrecognized = Unrecognized
    End With
    
    On Error Resume Next
    m_Keys.Add Add, strKeyKey
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise iksDupKey, "IniDOM", "Duplicate Key"
    End If
End Function

Public Property Get Count() As Integer
Attribute Count.VB_Description = "Number of Keys in the collection"
    Count = m_Keys.Count
End Property

Public Property Get Item(ByVal Index As Variant) As IniKey
Attribute Item.VB_UserMemId = 0
    'Default property.
    If VarType(Index) = vbString Then Index = UCase$(Index)
    On Error Resume Next
    Set Item = m_Keys.Item(Index)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise iksNoKey, "IniDOM", "Key not present"
    End If
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Iterator"
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'Hidden procedure, Procedure ID = -4
    Set NewEnum = m_Keys.[_NewEnum]
End Function

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Remove specificed Key from collection"
    If VarType(Index) = vbString Then Index = UCase$(Index)
    m_Keys.Remove Index
End Sub

Private Sub Class_Initialize()
    Set m_Keys = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Keys = Nothing
End Sub
