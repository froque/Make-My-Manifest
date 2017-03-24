VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniSections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Sections As Collection
Private m_intSectionComment As Integer

Public Enum IniSectionsExceptions
    issNoSection = &H80047D00
    issDupSection = &H80047D02
End Enum

Public Function Add( _
    Optional ByVal Name As String = "", _
    Optional ByVal Comment As String = "", _
    Optional ByVal Unrecognized As Boolean = False) As IniSection
    Dim strSectionKey As String
    
    If Len(Name) = 0 Then
        'This is a Section-level comment in the INI file.
        strSectionKey = ";" & CStr(m_intSectionComment)
        m_intSectionComment = m_intSectionComment + 1
    Else
        strSectionKey = UCase$(Name)
    End If
    
    Set Add = New IniSection
    With Add
        .Name = Name
        .Comment = Comment
        .Unrecognized = Unrecognized
    End With
    
    On Error Resume Next
    m_Sections.Add Add, strSectionKey
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise issDupSection, "IniDOM", "Duplicate Section"
    End If
End Function

Public Property Get Count() As Integer
Attribute Count.VB_Description = "Number of Sections in the collection"
    Count = m_Sections.Count
End Property

Public Property Get Item(ByVal Index As Variant) As IniSection
Attribute Item.VB_UserMemId = 0
    'Default property.
    If VarType(Index) = vbString Then Index = UCase$(Index)
    On Error Resume Next
    Set Item = m_Sections.Item(Index)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise issNoSection, "IniDOM", "Section not present"
    End If
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Iterator"
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'Hidden procedure, Procedure ID = -4
    Set NewEnum = m_Sections.[_NewEnum]
End Function

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Remove specified Section from collection"
    If VarType(Index) = vbString Then Index = UCase$(Index)
    m_Sections.Remove Index
End Sub

Private Sub Class_Initialize()
    Set m_Sections = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Sections = Nothing
End Sub
