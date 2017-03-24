VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FontWiz"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'=======
'FontWiz
'=======
'
'Adding the FontWiz class to a VB6 project produces a predeclared global
'object named FontWiz to your program.  You can call methods on this
'object at runtime to adapt the font used in Form and UserControl
'controls to the system-standard font.
'
'This can be seen as a second step in modernizing your VB6 program's UI,
'after taking the steps necessary to invoke the Common Controls 6.0
'themeing.  It is often neglected, leaving your programs still looking
'crude and a bit "off."
'
'In early versions of Windows the standard UI font is 8pt MS Sans Serif,
'and VB6 is hard-coded to enforce this unless you override the default
'manually or through new IDE module templates.
'
'Beginning with Windows 2000 the default was changed to 8pt Tahoma for a
'cleaner and more legible look.  Starting with Vista the system default
'is now 9pt Segoe UI to further improve clarity, especially for non-
'English language characters.
'
'But "dumb old Vb6" was never updated to accomodate the changes.  Thus
'FontWiz was born.
'
'
'OPERATION
'
'FontWiz interrogates the system for the font Name and Size used by the
'system MessageBox dialog as its baseline.  It creates a StdFont object
'to obtain the size of VB6's default font.  From these sizes it
'calculates the font ratio between the system font and the VB font.
'
'From here your Forms and UserControls must call the AdjustControl() or
'AdjustControls() methods to correct the VB6 font to the system font.
'On these calls you pass the control (or Me.Controls or
'UserControl.Controls) to the method and it updates the font settings
'for any font-enabled controls involved.
'
'FontName, FontNameDefault, FontSizeDefault, and FontSizeRatio
'properties are also exposed for runtime use as desired.
'
'
'USAGE
'
'Add FontWiz.cls to your VB6 Project and then insert the appropriate
'calls into your Forms and userControls.
'
'
'NOTES
'
'   o ONLY those controls where the font Name is equal to VB6's default
'     font Name will be changed.
'
'   o These font changes may result in text that is too large to
'     properly fit the size of your controls.  Make sure you leave some
'     leeway, or resize such controls appropriately at runtime after
'     adjusting the font.
'
'
'CLASS MEMBERS
'
'   METHODS
'
'   AdjustControl(ByVal Control As Control)
'
'       Adjust the Font.Name property of Control to the system UI font
'       and scale the Font.Size property of Control by the ratio
'       between the system UI font Size and VB6's default font Size.
'
'       Only changes font settings when Control's Font.Name is the
'       VB6 default font Name.
'
'   AdjustControls(ByVal Controls As Object)
'
'       Iterate through the Controls collection and call AdjustControl()
'       on each one ignoring any exceptions (due to a Control not having
'       a Font property).
'
'       In a Form this would usually be Me.Controls, and would be called
'       in the From's Load event handler.
'
'       In a UserControl this would usually be UserControl.Controls, and
'       would be called in the UserControl's ReadProperties event handler
'       if UserControl.Ambient.UserMode = True.
'
'   PROPERTIES
'
'   FontName As String (R/O)
'
'       Returns a String value containing the Name of the font-face used
'       as the current system UI font.  Might be used when creating
'       controls at runtime.
'
'   FontNameDefault As String (R/O)
'
'       Returns a String value containing the Name of the font-face used
'       as the VB6 default font.
'
'   FontSizeDefault As Currency (R/O)
'
'       Returns a Currency value containing the VB6 default font Size in
'       Points.  Note that setting a value for a Font.Size property to
'       something like 8 actually results in a value more like 8.25 in
'       most cases.  Using FontSizeDefault can help you avoid slightly
'       undersized fonts.
'
'   FontSizeRatio As Currency (R/O)
'
'       Returns a Currency value containing the ratio between the current
'       system MessageBox text font Size and the default VB6 font Size.
'       Might be used when creating controls at runtime, by multiplying
'       your "VB6 size" by this value.
'

Private Const LF_FACESIZE = 32
Private Const SPI_GETNONCLIENTMETRICS = 41

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Type NONCLIENTMETRICS
        cbSize As Long
        iBorderWidth As Long
        iScrollWidth As Long
        iScrollHeight As Long
        iCaptionWidth As Long
        iCaptionHeight As Long
        lfCaptionFont As LOGFONT
        iSMCaptionWidth As Long
        iSMCaptionHeight As Long
        lfSMCaptionFont As LOGFONT
        iMenuWidth As Long
        iMenuHeight As Long
        lfMenuFont As LOGFONT
        lfStatusFont As LOGFONT
        lfMessageFont As LOGFONT
End Type

Private Declare Function SystemParametersInfo Lib "user32" _
    Alias "SystemParametersInfoA" ( _
    ByVal uAction As Long, _
    ByVal uParam As Long, _
    ByRef lpvParam As Any, _
    ByVal fuWinIni As Long) As Long

Private m_FontName As String
Private m_FontNameDefault As String
Private m_FontSizeDefault As Currency
Private m_FontSizeRatio As Currency

Public Property Get FontName() As String
Attribute FontName.VB_Description = "Returns the system UI font name"
    FontName = m_FontName
End Property

Public Property Get FontNameDefault() As String
Attribute FontNameDefault.VB_Description = "Returns the VB6 default font Name"
    FontNameDefault = m_FontNameDefault
End Property

Public Property Get FontSizeDefault() As Currency
Attribute FontSizeDefault.VB_Description = "Returns the VB6 default font Size in Points"
    FontSizeDefault = m_FontSizeDefault
End Property

Public Property Get FontSizeRatio() As Currency
Attribute FontSizeRatio.VB_Description = "Returns the ratio between the system UI font Size and the VB6 default font Size"
    FontSizeRatio = m_FontSizeRatio
End Property

Public Sub AdjustControl(ByVal Control As Control)
Attribute AdjustControl.VB_Description = "Set the system UI font for the Control passed in"
    With Control.Font
        If .Name = m_FontNameDefault Then
            .Name = m_FontName
            .Size = .Size * m_FontSizeRatio
        End If
    End With
End Sub

Public Sub AdjustControls(ByVal Controls As Object)
Attribute AdjustControls.VB_Description = "Set the system UI font for all controls in the Controls collection passed in"
    Dim Control As Control
    
    On Error Resume Next
    For Each Control In Controls
        AdjustControl Control
    Next
    On Error GoTo 0
End Sub

Private Sub Class_Initialize()
    Dim fonDefault As StdFont
    Dim ncm As NONCLIENTMETRICS
    
    ncm.cbSize = LenB(ncm)
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0, ncm, 0
    Set fonDefault = New StdFont
    m_FontNameDefault = fonDefault.Name
    m_FontSizeDefault = fonDefault.Size
    m_FontSizeRatio = _
        (-ncm.lfMessageFont.lfHeight * Screen.TwipsPerPixelY / 20) / m_FontSizeDefault
    Set fonDefault = Nothing
    m_FontName = StrConv(ncm.lfMessageFont.lfFaceName, vbUnicode)
    m_FontName = Left$(m_FontName, InStr(m_FontName, vbNullChar) - 1)
End Sub
