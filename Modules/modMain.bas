Attribute VB_Name = "modMain"
Option Explicit

'----- Public Data -----

'----- Public Constants -----

'----- Private Methods -----

Private Sub Main()
    With App
        G_Product_Identity = _
            .ProductName & " " & CStr(.Major) & "." & CStr(.Minor) & "." & CStr(.Revision)
    End With
    G_Product_Settings_Path = GetProductSettingsPath()
    
    If Len(Command$()) > 0 Then
        'Command line initiation.
        MsgBox "Not implemented" '<<<<<<<<< @@@@@@@@@@@@@@@@@@@@@@@@@@@
    Else
        'GUI mode.
        AppEx.InitCommonControls
        
        frmMain.Show vbModeless
    End If
End Sub
