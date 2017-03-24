VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'ClsData
'=======
'
'This class is used by LibData to store class attributes
'for the createable classes found in ActiveX libraries
'referenced in the subject VB 6.0 project file.
'
'Usage:
'
'   Instances should only be created by LibData.  Members
'   of this class are meant to be read-only by programs
'   using LibData and ClsData.
'
'Dependencies:
'
'   none
'

'==================== Private data =======================

Private m_CLSID As String
Private m_Description As String
Private m_MiscStatusError As Boolean
Private m_MiscStatusAttributes As String 'May contain error description.
Private m_Name As String
Private m_ProgID As String
Private m_ThreadingModel As String

'===================== Properties ========================

Public Property Get CLSID() As String
    CLSID = m_CLSID
End Property

Friend Property Let CLSID(ByVal CLSID As String)
    m_CLSID = CLSID
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Friend Property Let Description(ByVal Description As String)
    m_Description = Description
End Property

Public Property Get MiscStatusError() As Boolean
    MiscStatusError = m_MiscStatusError
End Property

Friend Property Let MiscStatusError(ByVal MiscStatusError As Boolean)
    m_MiscStatusError = MiscStatusError
End Property

Public Property Get MiscStatusAttributes() As String
    MiscStatusAttributes = m_MiscStatusAttributes
End Property

Friend Property Let MiscStatusAttributes(ByVal MiscStatusAttributes As String)
    m_MiscStatusAttributes = MiscStatusAttributes
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Friend Property Let Name(ByVal Name As String)
    m_Name = Name
End Property

Public Property Get ProgID() As String
    ProgID = m_ProgID
End Property

Friend Property Let ProgID(ByVal ProgID As String)
    m_ProgID = ProgID
End Property

Public Property Get ThreadingModel() As String
    'Can be "" (empty String), have a value, or be vbNullString which
    'could mean "not a registered class" but as of version 0.12 we'll
    'assume these are non-creatable classes.
    ThreadingModel = m_ThreadingModel
End Property

Friend Property Let ThreadingModel(ByVal ThreadingModel As String)
    m_ThreadingModel = ThreadingModel
End Property

