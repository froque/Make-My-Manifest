VERSION 5.00
Begin VB.Form frmMake 
   Caption         =   "Make package"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4755
   ControlBox      =   0   'False
   LinkTopic       =   "frmMake"
   MDIChild        =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   4755
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   4500
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   4500
      Width           =   975
   End
   Begin VB.Timer tmrMake 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   180
      Top             =   4440
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish!"
      Height          =   375
      Left            =   2580
      TabIndex        =   1
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label lblCheck 
      BackStyle       =   0  'Transparent
      Caption         =   "¨"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   5
      Left            =   660
      TabIndex        =   19
      Top             =   3060
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblLegend 
      BackStyle       =   0  'Transparent
      Caption         =   "Package settings saved"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   5
      Left            =   1080
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "You can still make changes to many of them."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Index           =   2
      Left            =   660
      TabIndex        =   17
      Top             =   3180
      Width           =   3495
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Take a moment to review your option selections by going back through the other panels."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1395
      Index           =   1
      Left            =   660
      TabIndex        =   16
      Top             =   1380
      Width           =   3495
   End
   Begin VB.Label lblPrompt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Almost there!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C09000&
      Height          =   435
      Index           =   0
      Left            =   600
      TabIndex        =   15
      Top             =   660
      Width           =   3495
   End
   Begin VB.Label lblLegend 
      BackStyle       =   0  'Transparent
      Caption         =   "Manifest embedded"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   2
      Left            =   1080
      TabIndex        =   14
      Top             =   1500
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label lblCheck 
      BackStyle       =   0  'Transparent
      Caption         =   "¨"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   6
      Left            =   660
      TabIndex        =   13
      Top             =   3540
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblLegend 
      BackStyle       =   0  'Transparent
      Caption         =   "Activity log written"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   6
      Left            =   1080
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label lblCheck 
      BackStyle       =   0  'Transparent
      Caption         =   "¨"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   2
      Left            =   660
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblLegend 
      BackStyle       =   0  'Transparent
      Caption         =   "Added files copied"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   4
      Left            =   1080
      TabIndex        =   10
      Top             =   2580
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label lblCheck 
      BackStyle       =   0  'Transparent
      Caption         =   "¨"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   4
      Left            =   660
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblLegend 
      BackStyle       =   0  'Transparent
      Caption         =   "Dependencies copied"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   1080
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label lblCheck 
      BackStyle       =   0  'Transparent
      Caption         =   "¨"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   660
      TabIndex        =   7
      Top             =   1980
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblLegend 
      BackStyle       =   0  'Transparent
      Caption         =   "Manifest written"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label lblCheck 
      BackStyle       =   0  'Transparent
      Caption         =   "¨"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   660
      TabIndex        =   5
      Top             =   900
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblCheck 
      BackStyle       =   0  'Transparent
      Caption         =   "¨"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   660
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblLegend 
      BackStyle       =   0  'Transparent
      Caption         =   "Package folder prepared"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   420
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Image imgSplash 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   4230
      Left            =   60
      Top             =   60
      Width           =   4650
   End
End
Attribute VB_Name = "frmMake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----- Private Data -----

Private blnHaveAdditions As Boolean
Private stmManifest As ADODB.Stream

'----- Public Methods -----

'----- Private Methods -----

Private Sub CopyAdditions()
    Dim kysAdditions As IniKeys
    Dim kyAddition As IniKey
    Dim strFQTargetFolder As String
    Dim strLogCopyTo As String
    
    frmLog.Log "Copying additional files:"
    
    Set kysAdditions = G_Package_IniDOM!Additions.Keys
    For Each kyAddition In kysAdditions
        With G_Package_IniDOM.Sections("A:" & kyAddition.Name).Keys
            If CBool(!Included) Then
                strFQTargetFolder = TrimSlash(G_PackageFQFolder & !TargetFolder)
                If Len(Dir$(strFQTargetFolder, vbDirectory)) = 0 Then
                    MkDir strFQTargetFolder
                End If
                
                FileCopy !SourceLocation, strFQTargetFolder & "\" & !FileName
                
                strLogCopyTo = TrimSlash(G_PackageFolder & !TargetFolder)
                frmLog.Log "  " & !SourceLocation & " to " & strLogCopyTo
            End If
            'Else not Included.
        End With
    Next
    frmLog.Log
    lblCheck(4).Caption = "þ"
End Sub

Private Sub FinalDepInclusions()
    'Get final dependency inclusions.
    Dim issPackage As IniSections
    Dim itmLib As ListItem
    Dim ldLib As LibData
    
    Set issPackage = G_Package_IniDOM.Sections
    With frmDeps
        For Each itmLib In .lvwDeps.ListItems
            With itmLib
                Set ldLib = G_ProjLibData.Item(.Text)
                ldLib.Included = .Checked
                issPackage("D:" & .Text)!Included = CStr(.Checked)
            End With
        Next
    End With
End Sub

Private Sub FinalSettings()
    'Capture and log final settings.
    Dim iksPackage As IniKeys
    
    Set iksPackage = G_Package_IniDOM!MMMPackage.Keys
    With frmSettings
        G_PackageFolder = Trim$(.txtXCopy.Text)
        iksPackage!Folder = G_PackageFolder
        G_DepsFolder = Trim$(.txtDeps.Text)
        iksPackage!DepsFolder = G_DepsFolder
        iksPackage!EmbedManifest = CStr(.chkEmbedMan.Value = vbChecked)
        iksPackage!SaveLog = CStr(.chkSaveLog.Value = vbChecked)
        iksPackage!CC6 = CStr(.chkCC6.Value = vbChecked)
        iksPackage!KB828629 = CStr(.chkKB828629.Value = vbChecked)
        iksPackage!DPIAware = CStr(.chkDPIAware.Value = vbChecked)
        iksPackage!Compatibility = CStr(.chkCompatibility.Value = vbChecked)
        iksPackage!TrustInfo = CStr(.chkTrustInfo.Value = vbChecked)
        iksPackage!ExecutionLevel = .ExecutionLevel()
        iksPackage!UIAccess = CStr(.optUIAccess(1).Value)
    
        With frmLog
            .Log "MMM project settings:"
            .Log "  Embed manifest in EXE = " & iksPackage!EmbedManifest
            .Log "  Save MMM log to disk = " & iksPackage!SaveLog
            .Log "  Common Controls 6.0 manifest node = " & iksPackage!CC6
            .Log "  VB 6.0 ActiveX control KB 828629 remediation = " & iksPackage!KB828629
            .Log "  DPI Aware manifest node = " & iksPackage!DPIAware
            .Log "  Compatibility manifest node = " & iksPackage!Compatibility
            .Log "  TrustInfo manifest node = " & iksPackage!TrustInfo
            If frmSettings.chkTrustInfo.Value = vbChecked Then
                .Log "    Execution level = " & iksPackage!ExecutionLevel
                .Log "    UI Access = " & iksPackage!UIAccess
            End If
            .Log
        End With
    End With
End Sub

Private Sub ManifestApp()
    Dim bytAppMan() As Byte
    Dim strAppMan As String
    Dim strDescTag As String
    Dim lngManPad As Long
    
    'Use ADO Stream object to create manifest in UTF-8 format.
    Set stmManifest = New ADODB.Stream
    With stmManifest
        .Open
        .Charset = "UTF-8"
        .Type = adTypeText
    End With

    'Write application manifest head.
    bytAppMan = LoadResData("APPHEAD", "TEXT")
    strAppMan = StrConv(bytAppMan, vbUnicode)
    Erase bytAppMan
    strAppMan = Replace$(strAppMan, _
                         "[MMMID]", _
                         FormatXML(G_Product_Identity))
    If Len(G_ProjCompany) = 0 Then
        strAppMan = Replace$(strAppMan, _
                             "[APPNAME]", _
                             FormatXML(G_ProjName))
    Else
        strAppMan = Replace$(strAppMan, _
                             "[APPNAME]", _
                             FormatXML(Replace$(G_ProjCompany, " ", ".") & "." & G_ProjName))
    End If
    strAppMan = Replace$(strAppMan, _
                         "[VERSION]", _
                         G_ProjVersion)
    strDescTag = FormatXML(G_ProjDescription)
    If Len(strDescTag) > 0 Then
        strDescTag = "<description>" & strDescTag & "</description>"
    End If
    strAppMan = Replace$(strAppMan, _
                         "[APPDESC]", _
                         strDescTag)
    stmManifest.WriteText strAppMan, adWriteChar
    
    'Write Windows 7 Compatibility node if requested.
    If frmSettings.chkCompatibility.Value = vbChecked Then
        bytAppMan = LoadResData("COMPAT", "TEXT")
        strAppMan = StrConv(bytAppMan, vbUnicode)
        Erase bytAppMan
        stmManifest.WriteText strAppMan, adWriteChar
    End If
    
    'Write Common Controls 6 node if requested.
    If frmSettings.chkCC6.Value = vbChecked Then
        bytAppMan = LoadResData("APPCC6", "TEXT")
        strAppMan = StrConv(bytAppMan, vbUnicode)
        Erase bytAppMan
        stmManifest.WriteText strAppMan, adWriteChar
    End If

    'Copy component files, write component file entries in manifest.
    frmLog.Log "Copying dependencies:"
    ManifestComponents
    frmLog.Log
    
    'Write dpiAware node if requested.
    If frmSettings.chkDPIAware.Value = vbChecked Then
        bytAppMan = LoadResData("DPIAWARE", "TEXT")
        strAppMan = StrConv(bytAppMan, vbUnicode)
        Erase bytAppMan
        stmManifest.WriteText strAppMan, adWriteChar
    End If
    
    'Write trustInfo node if requested.
    If frmSettings.chkTrustInfo.Value = vbChecked Then
        bytAppMan = LoadResData("TRUSTINFO", "TEXT")
        strAppMan = StrConv(bytAppMan, vbUnicode)
        Erase bytAppMan
        strAppMan = Replace$(strAppMan, _
                             "[EXLEVEL]", _
                             frmSettings.ExecutionLevel())
        strAppMan = Replace$(strAppMan, _
                             "[UIACCESS]", _
                             IIf(frmSettings.optUIAccess(1).Value, "true", "false"))
        stmManifest.WriteText strAppMan, adWriteChar
    End If
    
    'Write application manifest foot.
    bytAppMan = LoadResData("APPFOOT", "TEXT")
    strAppMan = StrConv(bytAppMan, vbUnicode)
    Erase bytAppMan
    stmManifest.WriteText strAppMan, adWriteChar
    
    'Copy project EXE file.
    frmLog.Log "Copying EXE:"
    FileCopy G_ProjEXE_FQFileName, G_PackageFQFolder & G_ProjEXE
    With frmLog
        .Log "  " & G_ProjEXE_FQFileName & " to " & TrimSlash(G_PackageFolder)
        .Log
    End With

    With stmManifest
        'Pad to DWord boundary for possible embedding.
        lngManPad = .Size Mod 4
        If lngManPad > 0 Then
            .Position = .Size
            .WriteText Space$(4 - lngManPad)
        End If
        lblCheck(1).Caption = "þ"
        
        If frmSettings.chkEmbedMan.Value = vbChecked Then
            EmbedManifest G_PackageFQFolder & G_ProjEXE, stmManifest
            With frmLog
                .Log "Embedding application manifest in:"
                .Log "  " & G_ProjEXE
            End With
            lblCheck(2).Caption = "þ"
        Else
            .SaveToFile G_PackageFQFolder & G_ProjEXE & ".manifest"
            With frmLog
                .Log "Writing application manifest:"
                .Log "  " & G_ProjEXE & ".manifest to " & G_PackageFolder
            End With
        End If
        .Close
        frmLog.Log
    End With
    
    lblCheck(3).Caption = "þ"
End Sub

Private Sub ManifestComponents()
    Dim bytAssemManTempl() As Byte
    Dim strAssemMan As String
    Dim ldLib As LibData
    Dim lngClass As Long
    Dim strReportCopyTo As String
    Dim strProgIDAttr As String
    Dim strTModelAttr As String
    
    strReportCopyTo = TrimSlash(G_PackageFolder & "\" & G_DepsFolder)
    For Each ldLib In G_ProjLibData
        'Write manifest component file head if not excluded and a COM library.
        With ldLib
            If .Included Then
                FileCopy .FileLocation, _
                         G_PackageFQFolder _
                       & G_DepsFolder _
                       & SimpleFileName(.FileLocation)
                bytAssemManTempl = LoadResData("FILEHEAD", "TEXT")
                strAssemMan = StrConv(bytAssemManTempl, vbUnicode)
                Erase bytAssemManTempl
                
                If .IsCOM Then
                    strAssemMan = Replace$(strAssemMan, _
                                           "[LIBFILE]", _
                                           G_DepsFolder & SimpleFileName(.FileLocation))
                    strAssemMan = Replace$(strAssemMan, _
                                           "[LOADFROM]", _
                                           "")
                    stmManifest.WriteText strAssemMan, adWriteChar
                    
                    bytAssemManTempl = LoadResData("TYPELIB", "TEXT")
                    strAssemMan = StrConv(bytAssemManTempl, vbUnicode)
                    Erase bytAssemManTempl
                    strAssemMan = Replace$(strAssemMan, _
                                           "[LIBID]", _
                                           .LIBID)
                    strAssemMan = Replace$(strAssemMan, _
                                           "[VERSION]", _
                                           .Version)
                    strAssemMan = Replace$(strAssemMan, _
                                           "[FLAGS]", _
                                           .Flags)
                    stmManifest.WriteText strAssemMan, adWriteChar
                    
                    frmLog.Log "  " & .FileLocation & " to " & strReportCopyTo
                    
                    For lngClass = 0 To ldLib.Count - 1
                        With .Class(lngClass)
'''This is no longer conditional as of version 0.12, since there may well be
'''legitimate non-creatable classes in an OCX that won't have a ThreadingModel.
'''                            'Registered class?  Then write component file class.
'''                            If StrPtr(.ThreadingModel) <> 0 Then
                                'Write manifest component file class tag.
                            bytAssemManTempl = LoadResData("CLASS", "TEXT")
                            strAssemMan = StrConv(bytAssemManTempl, vbUnicode)
                            Erase bytAssemManTempl
                            strAssemMan = Replace$(strAssemMan, _
                                                   "[CLSID]", _
                                                   .CLSID)
                            strTModelAttr = .ThreadingModel
                            If Len(strTModelAttr) > 0 Then
                                strTModelAttr = " threadingModel=""" & .ThreadingModel & """"
                            End If
                            strAssemMan = Replace$(strAssemMan, _
                                                   "[TMODEL]", _
                                                   strTModelAttr)
                            strAssemMan = Replace$(strAssemMan, _
                                                   "[LIBID]", _
                                                   ldLib.LIBID)
                            strAssemMan = Replace$(strAssemMan, _
                                                   "[DESC]", _
                                                   FormatXML(.Description))
                            strProgIDAttr = .ProgID
                            If Len(strProgIDAttr) > 0 Then
                                strProgIDAttr = " progid=""" & FormatXML(strProgIDAttr) & """"
                            End If
                            strAssemMan = Replace$(strAssemMan, _
                                                   "[PROGID]", _
                                                   strProgIDAttr)
                            strAssemMan = Replace$(strAssemMan, _
                                                   "[MISCSTATUS]", _
                                                   IIf(frmSettings.chkKB828629.Value = vbChecked, _
                                                       .MiscStatusAttributes, _
                                                       ""))
                            stmManifest.WriteText strAssemMan, adWriteChar
'''                            End If
                        End With
                    Next
                Else
                    'Non-COM, slightly different <FILE> tag.
                    If Len(G_DepsFolder) > 0 Then
                        strAssemMan = Replace$(strAssemMan, _
                                               "[LIBFILE]", _
                                               SimpleFileName(.FileLocation))
                        strAssemMan = Replace$(strAssemMan, _
                                               "[LOADFROM]", _
                                               " loadFrom=""" _
                                             & G_DepsFolder & SimpleFileName(.FileLocation) _
                                             & """")
                        stmManifest.WriteText strAssemMan, adWriteChar
                    End If
                End If
                    
                'Write manifest component file foot.
                bytAssemManTempl = LoadResData("FILEFOOT", "TEXT")
                strAssemMan = StrConv(bytAssemManTempl, vbUnicode)
                Erase bytAssemManTempl
                stmManifest.WriteText strAssemMan, adWriteChar
            End If
        End With
    Next
End Sub

Private Sub SetUIFailed()
    Dim intLbl As Integer
    
    'Hide "Make" prompt, checklist, and splash background.
    For intLbl = 0 To 2
        lblPrompt(intLbl).Visible = False
    Next
    For intLbl = 0 To 6
        lblCheck(intLbl).Visible = False
        lblLegend(intLbl).Visible = False
    Next
    imgSplash.Visible = False
    
    UICancelCleanup
End Sub

Private Sub UICancelCleanup()
    frmSettings.UICancelCleanup
    cmdCancel.Caption = "E&xit"
    cmdFinish.Enabled = False
End Sub

Private Function VerifyPackageFolder() As Boolean
    'Returns True if user cancels the Make operation here.
    Dim intResult As Integer
    Dim strDir As String
    
    With frmLog
        'Check for package folder.
        G_PackageFQFolder = G_VBP_Path & G_PackageFolder 'Add trailing "\" ASAP though!
        If Len(Dir$(G_PackageFQFolder, vbDirectory)) > 0 Then
            .Log G_PackageFolder & " package folder exists:"
            .Log "    " & G_PackageFQFolder
            intResult = _
                MsgBox(G_PackageFolder & " package folder exists." _
                     & vbNewLine & vbNewLine _
                     & "Replace with new package?", _
                       vbYesNo Or vbExclamation, _
                       "MMM - Package folder exists")
            If intResult = vbNo Then
                .Log "    Stopping with no changes made."
                SetUIFailed
                frmMain.sbMain.SimpleText = "Stopping with no changes made.  See log."
                VerifyPackageFolder = True
            Else
                G_PackageFQFolder = G_PackageFQFolder & "\"
                .Log "        Replacing package folder contents."
                On Error Resume Next
                'Might be empty!
                Kill G_PackageFQFolder & "*.*"
                On Error GoTo 0
                strDir = Dir$(G_PackageFQFolder & "*.*", vbDirectory)
                Do While Len(strDir) > 0
                    If Left$(strDir, 1) <> "." Then
                        'Clean out the directory and remove it.
                        On Error Resume Next
                        Kill G_PackageFQFolder & strDir & "\*.*"
                        On Error GoTo 0
                        RmDir G_PackageFQFolder & strDir
                    End If
                    
                    strDir = Dir$()
                Loop
                If Len(G_DepsFolder) > 0 Then
                    MkDir G_PackageFQFolder & G_DepsFolder
                    G_DepsFolder = G_DepsFolder & "\"
                End If
    
                lblCheck(0).Caption = "þ"
            End If
        Else
            .Log "Creating " & G_PackageFolder & " package folder:"
            .Log "    " & G_PackageFQFolder
            MkDir G_PackageFQFolder
            G_PackageFQFolder = G_PackageFQFolder & "\"
            If Len(G_DepsFolder) > 0 Then
                MkDir G_PackageFQFolder & G_DepsFolder
                G_DepsFolder = G_DepsFolder & "\"
            End If
            lblCheck(0).Caption = "þ"
        End If
        .Log
    End With
End Function

'----- Event Handlers -----

Private Sub cmdBack_Click()
    frmMain.mnuViewSettings_Click
End Sub

Private Sub cmdCancel_Click()
    If G_ExitMode Then
        Unload frmMain
    Else
        G_ExitMode = True
        SetUIFailed
        frmMain.sbMain.SimpleText = "Project processing canceled."
    End If
End Sub

Private Sub cmdFinish_Click()
    Dim intLbl As Integer
    Dim intAdditions As Integer
    Dim intButton As Integer
    Dim intButtonMenu As Integer
    
    'Hide "Make" prompt, show checklist.
    For intLbl = 0 To 2
        lblPrompt(intLbl).Visible = False
    Next
    For intLbl = 0 To 6
        lblCheck(intLbl).Visible = True
        lblLegend(intLbl).Visible = True
    Next
    On Error Resume Next
    intAdditions = G_Package_IniDOM!Additions.Keys.Count
    If Err.Number = 0 Then
        'We have some?  Section exists at least.
        On Error GoTo 0
        If intAdditions > 0 Then
            'We have some!
            blnHaveAdditions = True
            lblCheck(4).ForeColor = vbBlack
            lblLegend(4).ForeColor = vbBlack
        End If
    End If
    'Else no Section.
    On Error GoTo 0
    With frmSettings
        If .chkEmbedMan.Value = vbChecked Then
            lblCheck(2).ForeColor = vbBlack
            lblLegend(2).ForeColor = vbBlack
        End If
        If .chkSaveLog.Value = vbChecked Then
            lblCheck(6).ForeColor = vbBlack
            lblLegend(6).ForeColor = vbBlack
        End If
    End With
    
    With frmMain
        'Clear menu.
        .mnuViewProject.Enabled = False
        .mnuViewScan.Enabled = False
        .mnuViewDependencies.Enabled = False
        .mnuViewAddedFiles.Enabled = False
        .mnuViewSettings.Enabled = False
        
        'Toolbar setup.
        With .tbrMain
            For intButton = GC_TBBTN_OPEN To GC_TBBTN_SETTINGS
                If intButton <> GC_TBBTN_LOG Then .Buttons(intButton).Enabled = False
            Next
            With .Buttons(GC_TBBTN_OPEN)
                For intButtonMenu = 1 To 2
                    .ButtonMenus(intButtonMenu).Enabled = False
                Next
            End With
        End With
        .sbMain.SimpleText = "Creating deployment package..."
    End With
    
    UICancelCleanup
    cmdFinish.Visible = False
    cmdBack.Visible = False
    cmdCancel.Caption = "E&xit"
    G_ExitMode = True 'Allow Exit action.
    cmdCancel.Enabled = False
    
    MousePointer = vbHourglass
    tmrMake.Enabled = True
End Sub

Private Sub Form_GotFocus()
    If cmdFinish.Enabled Then cmdFinish.SetFocus
End Sub

Private Sub Form_Load()
    FontWiz.AdjustControls Me.Controls

    imgSplash.Picture = frmSplash.Picture

    'View frmMake now.
    With frmMain
        .mnuViewMake.Enabled = True
        With .tbrMain.Buttons(GC_TBBTN_MAKE)
            .Enabled = True
            .Value = tbrPressed
        End With
        .sbMain.SimpleText = "Ready to build the deployment package."
    End With
    Show
    ZOrder vbBringToFront
    cmdFinish.SetFocus
End Sub

Private Sub Form_Resize()
    Dim sngVal As Single
    
    If WindowState <> vbMinimized Then
        sngVal = ScaleHeight - GC_BUTTONSTOPDELTA
        If ScaleWidth - GC_CANCELLEFTDELTA > 0 And sngVal > 0 Then
            cmdBack.Top = sngVal
            cmdBack.Left = ScaleWidth - GC_BACKLEFTDELTA
            cmdFinish.Top = sngVal
            cmdFinish.Left = ScaleWidth - GC_NEXTLEFTDELTA
            cmdCancel.Top = sngVal
            cmdCancel.Left = ScaleWidth - GC_CANCELLEFTDELTA
        End If
    End If
End Sub

Private Sub tmrMake_Timer()
    Dim intFSaveLog As Integer
    
    tmrMake.Enabled = False
    
    FinalDepInclusions
    FinalSettings
    If Not VerifyPackageFolder() Then
        ManifestApp
        
        If blnHaveAdditions Then CopyAdditions
        
        With frmLog
            .Log "Writing MMM package project file:"
            .Log "  " & G_VBP_Path & G_PackageFolder & ".mmmp"
            .Log
            IniDOMToFile G_Package_IniDOM, G_VBP_Path & G_PackageFolder & ".mmmp"
            lblCheck(5).Caption = "þ"
            
            .Log "Complete!"
        End With
        
        If frmSettings.chkSaveLog.Value = vbChecked Then
            intFSaveLog = FreeFile()
            Open G_VBP_Path & G_PackageFolder & ".log" For Output As #intFSaveLog
            Print #intFSaveLog, frmLog.txtLog.Text
            Close #intFSaveLog
            lblCheck(6).Caption = "þ"
        End If
            
        frmMain.sbMain.SimpleText = "Deployment package complete."
    End If

    cmdCancel.Enabled = True 'Allow Exit.
    MousePointer = vbDefault
End Sub
