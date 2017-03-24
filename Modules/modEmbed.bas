Attribute VB_Name = "modEmbed"
Option Explicit

'---- Private Data ----

Private Enum BOOLWIN32
    b32false = 0
    b32True = 1
End Enum

'---- Private Declares ----

Private Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" ( _
    ByVal pFileName As String, _
    ByVal bDeleteExistingResources As BOOLWIN32) As Long
    
Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" ( _
    ByVal hUpdate As Long, _
    ByVal lpType As Long, _
    ByVal lpName As Long, _
    ByVal wLanguage As Integer, _
    ByRef lpData As Byte, _
    ByVal cbData As Integer) As BOOLWIN32

Private Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" ( _
    ByVal hUpdate As Long, _
    ByVal fDiscard As BOOLWIN32) As BOOLWIN32

'---- Public Methods ----

Public Sub EmbedManifest(ByVal EXEFileName As String, ByVal Manifest As ADODB.Stream)
    Dim hEXE As Long
    Dim bytManifest() As Byte
    Dim lngUpdateError As Long
    
    hEXE = BeginUpdateResource(EXEFileName, b32false)
    If hEXE = 0 Then
        Err.Raise vbObjectError Or &H3802&, _
                  "EmbedManifest", _
                  "Failed to get handle to EXE, system error " & CStr(Err.LastDllError)
    Else
        With Manifest
            .Position = 0
            .Type = adTypeBinary
            bytManifest = .Read(adReadAll)
            If UpdateResource(hEXE, 24&, 1&, 0, bytManifest(0), .Size) = b32false Then
                lngUpdateError = Err.LastDllError
                If EndUpdateResource(hEXE, b32True) = b32false Then
                    Err.Raise vbObjectError Or &H3804&, _
                              "EmbedManifest", _
                              "Failed to flush resource update, system error " _
                            & CStr(Err.LastDllError)
                End If
                Err.Raise vbObjectError Or &H3806&, _
                          "EmbedManifest", _
                          "Failed to update manifest resource, system error " _
                        & CStr(lngUpdateError)
            Else
                If EndUpdateResource(hEXE, b32false) = b32false Then
                    Err.Raise vbObjectError Or &H3808&, _
                              "EmbedManifest", _
                              "Failed to commit resource update, system error " _
                            & CStr(Err.LastDllError)
                End If
            End If
        End With
    End If
End Sub


