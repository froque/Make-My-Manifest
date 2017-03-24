VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Keys As IniKeys

Public Comment As String
Attribute Comment.VB_VarDescription = "Section comment"
Public Name As String
Attribute Name.VB_VarDescription = "Section name"
Public Unrecognized As Boolean
Attribute Unrecognized.VB_VarDescription = "If True this is an unregonized INI file line, saved literally"

Public Property Get Keys() As IniKeys
Attribute Keys.VB_Description = "Returns the Keys collection for this Section"
Attribute Keys.VB_UserMemId = 0
    Set Keys = m_Keys
End Property

Private Sub Class_Initialize()
    Set m_Keys = New IniKeys
End Sub

Private Sub Class_Terminate()
    Set m_Keys = Nothing
End Sub
