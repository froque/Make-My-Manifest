VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Comment As String
Attribute Comment.VB_VarDescription = "Key comment"
Public Unrecognized As Boolean
Attribute Unrecognized.VB_VarDescription = "If True this is an unregonized INI file line, saved literally"

Private m_Name As String
Private m_QuoteName As Boolean
Private m_QuoteValue As Boolean
Private m_Value As String
Attribute m_Value.VB_VarUserMemId = 0
Attribute m_Value.VB_VarDescription = "Key value"

Public Enum IniKeyExceptions
    ikeQuoteName = &H80047D20
    ikeQuoteValue = &H80047D22
End Enum

Public Property Get Name() As String
Attribute Name.VB_Description = "Key name"
    Name = m_Name
End Property

Public Property Let Name(ByVal Name As String)
    If InStr(Name, """") > 0 And QuoteName Then
        Err.Raise ikeQuoteName, "IniKey", "Quoted quotes not allowed in Key Names"
    End If
    
    If InStr(Name, "=") > 0 Then m_QuoteName = True
    m_Name = Name
End Property

Public Property Get QuoteName() As Boolean
Attribute QuoteName.VB_Description = "Name is to be quoted on Save"
    QuoteName = m_QuoteName
End Property

Public Property Let QuoteName(ByVal QuoteName As Boolean)
    If (InStr(m_Name, """") > 0) And QuoteName Then
        Err.Raise ikeQuoteName, "IniKey", "Quoted quotes not allowed in Key Names"
    End If

    m_QuoteName = QuoteName
End Property

Public Property Get QuoteValue() As Boolean
Attribute QuoteValue.VB_Description = "Value is to be quoted on Save"
    QuoteValue = m_QuoteValue
End Property

Public Property Let QuoteValue(ByVal QuoteValue As Boolean)
    If InStr(m_Value, """") > 0 And QuoteValue Then
        Err.Raise ikeQuoteValue, "IniKey", "Quoted quotes not allowed in Key Values"
    End If

    m_QuoteValue = QuoteValue
End Property

Public Property Get Value() As String
Attribute Value.VB_Description = "Key value"
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property

Public Property Let Value(ByVal Value As String)
    If InStr(Value, """") > 0 And QuoteName Then
        Err.Raise ikeQuoteValue, "IniKey", "Quoted quotes not allowed in Key Values"
    End If
    
    If (InStr(Value, ";") > 0) Or (Right$(Value, 1) = "\") Then m_QuoteValue = True
    m_Value = Value
End Property
