Attribute VB_Name = "modTBEx"
Option Explicit
'
'Extend the ability to store and retrieve text in a TextBox beyond 64K.
'

'Window message consts.
Private Const WM_SETTEXT = &HC&
Private Const WM_GETTEXT = &HD&
Private Const WM_GETTEXTLENGTH = &HE&
Private Const WM_USER = &H400&
Private Const EM_SCROLLCARET = WM_USER + 49&
Private Const EM_GETSEL = &HB0&
Private Const EM_SETSEL = &HB1&
Private Const EM_REPLACESEL = &HC2&
Private Const EM_SETMARGINS = &HD3&
Private Const EC_LEFTMARGIN = &H1&
Private Const EC_RIGHTMARGIN = &H2&

Private Declare Function SendMessageWLng Lib "user32" Alias "SendMessageW" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Function LoHi(ByVal Low As Long, ByVal High As Long) As Long
    'Pack two UINT values (passed as Long) into a ULONG value (as Long).
    Low = Low And &HFFFF&
    High = High And &HFFFF&
    LoHi = Low Or ((High And &H7FFF) * &H10000) Or IIf(CBool(High And &H8000&), &H80000000, 0)
End Function

Public Sub TBGetSel(ByVal TB As TextBox, ByRef SelStart As Long, ByRef SelLen As Long)
    'Get current SelStart and SelLen of TB.  SelStart is 0-based.
    SendMessageWLng TB.hWnd, EM_GETSEL, VarPtr(SelStart), VarPtr(SelLen)
    SelLen = SelLen - SelStart
End Sub

Public Function TBLen(ByVal TB As TextBox) As Long
    'Get length of Text contents of TB.
    TBLen = SendMessageWLng(TB.hWnd, WM_GETTEXTLENGTH, 0, 0)
End Function

Public Function TBRead(ByVal TB As TextBox, Optional ByVal MaxChars As Long = 65536) As String
    'Read Text contents of TB.
    TBRead = String$(MaxChars, 0)
    TBRead = Left$(TBRead, SendMessageWLng(TB.hWnd, WM_GETTEXT, MaxChars, StrPtr(TBRead)))
End Function

Public Sub TBScrollCaret(ByVal TB As TextBox)
    'Scrolls TB so that the insertion point (caret) is in view.
    SendMessageWLng TB.hWnd, EM_SCROLLCARET, 0, 0
End Sub

Public Sub TBSetMargins(ByVal TB As TextBox, ByVal LeftPixels As Long, ByVal RightPixels As Long)
    'Sets the left/right margins of TB in pixels.
    SendMessageWLng TB.hWnd, _
                    EM_SETMARGINS, EC_LEFTMARGIN Or EC_RIGHTMARGIN, _
                    LoHi(LeftPixels, RightPixels)
End Sub

Public Sub TBSetSel(ByVal TB As TextBox, ByVal SelStart As Long, Optional ByVal SelLen As Long = 0)
    'Sets the selection for TB.  SelStart is 0-based.
    '   o Use a very high value like &H7FFFFFFF to set the
    '     selection/insertion point at end.
    '   o Pass SelStart=0 and SelLen=0 to select all.
    SendMessageWLng TB.hWnd, EM_SETSEL, SelStart, SelStart + SelLen - 1
End Sub

Public Sub TBWrite(ByVal TB As TextBox, ByRef Text As String)
    'Replaces all Text in TB.
    If Not SendMessageWLng(TB.hWnd, WM_SETTEXT, 0, StrPtr(Text)) Then
        Err.Raise &H80043300, "TBWrite", "Insufficient space in " & TB.Name
    End If
End Sub

Public Sub TBWriteSel(ByVal TB As TextBox, ByRef Text As String)
    'Replaces selected Text (or inserts at insertion point) in TB w/o undo.
    SendMessageWLng TB.hWnd, EM_REPLACESEL, 0, StrPtr(Text)
End Sub

