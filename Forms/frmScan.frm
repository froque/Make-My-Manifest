VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScan 
   Caption         =   "Scan report"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   ControlBox      =   0   'False
   LinkTopic       =   "frmScan"
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   6915
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   4740
      TabIndex        =   5
      Top             =   5280
      Width           =   975
   End
   Begin VB.Timer tmrStartProcessing 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   120
      Top             =   5100
   End
   Begin VB.TextBox txtScanEvent 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   6885
   End
   Begin MSComctlLib.ImageList imlScanEvents 
      Left            =   720
      Top             =   5100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScan.frx":0000
            Key             =   "Green"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScan.frx":015A
            Key             =   "Yellow"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScan.frx":02B4
            Key             =   "Red"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwScanEvents 
      CausesValidation=   0   'False
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   3625
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlScanEvents"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Severity"
         Text            =   "Severity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Event"
         Text            =   "Event"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Events noted during scan:"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   2820
      Width           =   2715
   End
   Begin VB.Label Label1 
      Caption         =   "Event detail:"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   2715
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----- Private Data -----

'lvwScanEvents Subitems.
Private Const SCANEVT_SUBITEM_SEVERITY = 0
Private Const SCANEVT_SUBITEM_EVENT = 1

Public Enum ScanEventSeverities
    sesComment = 1
    sesNote
    sesFatal
End Enum

'Form resize events.
Private Const FORM_SCALEHEIGHT As Single = 5700# 'UPDATE THESE IF FORM DESIGN LAYOUT CHANGES

Private sngScanEventsHeightDelta As Single

'Misc.
Private blnPhaseComplete As Boolean 'Already completed this phase.

'----- Public Methods -----

Public Function ManualRefByFile(ByVal LibFileLocation As String) As Boolean
    'Returns True if new Unreg Warning.
    Dim ldLib As LibData
    Dim blnEntryUnregWarning As Boolean 'Cache prior warning.

    blnEntryUnregWarning = G_UnregWarning
    G_UnregWarning = False
    Set ldLib = New LibData
    With ldLib
        If .LoadManRefByFileName(LibFileLocation) Then
            'Warn unless we're not going to use it anyway.
            G_UnregWarning = Not .IncludedNever
        End If
        
        'Test for Lib in collection already by adding it.
        On Error Resume Next
        G_ProjLibData.Add ldLib, .Name
        If Err.Number = 0 Then
            On Error GoTo 0
        
            'New dependency found.  Record settings.
            With G_Package_IniDOM.Sections
                'Add library to list.
                !Dependencies.Keys.Add ldLib.Name
                
                'Add and populate dependency's Section.
                .Add "D:" & ldLib.Name
                With .Item("D:" & ldLib.Name).Keys
                    .Add "Source", ldLib.Source
                    .Add "Included", CStr(ldLib.Included)
                    .Add "IncludedNever", CStr(ldLib.IncludedNever)
                    .Add "FileLocation", ldLib.FileLocation, QuoteValue:=True
                End With
                .Add 'Blank line, Section level.
            End With
            
            'Log it.
            frmLog.Log "Manually added:"
            LogLibrary ldLib
            frmProject.UpdateProp GPP_PERMEXCLUDED, G_PermExcludedCount
            
            'Process any DEP file (recurses).
            ProjectProcessDepFile .FileLocation
        Else
            On Error GoTo 0
        End If
    End With
    
    ManualRefByFile = G_UnregWarning
    G_UnregWarning = blnEntryUnregWarning Or G_UnregWarning 'Restore and update.
End Function

Public Sub UICancelCleanup()
    frmProject.UICancelCleanup
    With frmMain
        .mnuFileOpenVBP.Enabled = False
        .mnuFileOpenMMMP.Enabled = False
        .tbrMain.Buttons(GC_TBBTN_OPEN).Enabled = False
    End With
    cmdCancel.Caption = "E&xit"
End Sub

'----- Private Methods -----

Private Sub AddEvent(ByVal Severity As ScanEventSeverities, ByVal NewEvent As String)
    Dim itmEvent As ListItem
    
    Set itmEvent = lvwScanEvents.ListItems.Add()
    With itmEvent
        .SmallIcon = Severity
        .Text = Choose(Severity, "Comment", "Note", "Fatal")
        .Bold = True
        .SubItems(SCANEVT_SUBITEM_EVENT) = NewEvent
        .Selected = False
    End With
End Sub

Private Function LogLibrary(ByRef LibData As LibData) As Boolean
    Dim lngClass As Long
    Dim sesSeverity As ScanEventSeverities
    Dim strProgID As String
    Dim strExcludeType As String
    Dim strTModelAttr As String
    Dim strMiscStatusAttr As String
    
    With frmLog
        .Log "  " & LibData.Name
        .Log "    Description = " & LibData.Description
        .Log "    File location = " & LibData.FileLocation
        .Log "    Typlib ID = " & LibData.LIBID
        .Log "    Version = " & LibData.Version
        .Log "    Flags = " & IIf(Len(LibData.Flags) > 0, LibData.Flags, "*none*")
        If LibData.Included Then
            .Log
            
            .Log "    CoClasses: " & CStr(LibData.Count)
            .Log
            For lngClass = 0 To LibData.Count - 1
                .Log "      " & LibData.Class(lngClass).Name
                .Log "        Description = " & LibData.Class(lngClass).Description
                strProgID = LibData.Class(lngClass).ProgID
                If Len(strProgID) > 0 Then
                    .Log "        ProgID = " & strProgID
                End If
                
                strTModelAttr = LibData.Class(lngClass).ThreadingModel
'''We no longer treat a vbNullString value as meaning "not registered" but
'''instead as of version 0.12 assume we have a legitimate non-creatable class
'''instead.
'''                If StrPtr(strTModelAttr) = 0 Then
'''                    strTModelAttr = "Not registered, Class will be skipped."
'''                ElseIf Len(strTModelAttr) = 0 Then
                If Len(strTModelAttr) = 0 Then
                    strTModelAttr = "*null*, no ThreadingModel entry will be made in manifest."
                End If
                .Log "        Threading model = " & strTModelAttr
                .Log "        CLSID = " & LibData.Class(lngClass).CLSID
                
                strTModelAttr = LibData.Class(lngClass).ThreadingModel
                strMiscStatusAttr = LibData.Class(lngClass).MiscStatusAttributes
                If Len(strMiscStatusAttr) > 0 And StrPtr(strTModelAttr) <> 0 Then
                    If LibData.Class(lngClass).MiscStatusError Then
                        AddEvent sesFatal, _
                                 "VB6 KB 828629 remediation error" & "·" _
                               & "·" _
                               & "  Library = " & LibData.Name & "·" _
                               & "  File location = " & LibData.FileLocation & "·" _
                               & "·" _
                               & "  " & LibData.Class(lngClass).MiscStatusAttributes
                        .Log "        VB6 KB 828629 remediation error: " _
                           & LibData.Class(lngClass).MiscStatusAttributes
                        .Log "          Processing will be aborted!"
                        LogLibrary = True
                    Else
                        .Log "        VB6 KB 828629 remediation:"
                        strMiscStatusAttr = Replace$(strMiscStatusAttr, _
                                                     " ", _
                                                     vbNewLine & Space$(10))
                        strMiscStatusAttr = Replace$(strMiscStatusAttr, _
                                                     "=", _
                                                     " = " & vbNewLine & Space$(18))
                        strMiscStatusAttr = Replace$(strMiscStatusAttr, _
                                                     ",", _
                                                     "," & vbNewLine & Space$(18))
                        .Log Mid$(strMiscStatusAttr, 3)
                    End If
                End If
            
                .Log
            Next
        Else
            If LibData.IncludedNever Then
                G_PermExcludedCount = G_PermExcludedCount + 1
                sesSeverity = sesComment
                strExcludeType = "hard"
            Else
                sesSeverity = sesNote
                strExcludeType = "soft"
            End If
            With LibData
                AddEvent sesSeverity, _
                         "Library " & strExcludeType & "-excluded by MMM:" & "·" _
                       & "·" _
                       & "  Library = " & .Name & "·" _
                       & "  Description = " & .Description & "·" _
                       & "  File location = " & .FileLocation & "·" _
                       & "  Typelib ID = " & .LIBID & "·" _
                       & "  Version = " & .Version & "·" _
                       & "  Source = """ & .Source & """"
            End With
            .Log "    *Excluded*"
            .Log
        End If
    End With
End Function

Private Function ProjectProcess() As Boolean
    'Returns True if an error occurred which must abort processing.
    ProjectProcess = ProjectProcessFileScan()
    If Not ProjectProcess Then
        ProjectProcess = ProjectProcessLog()
    End If
End Function

Private Function ProjectProcessFileScan() As Boolean
    'Returns True if an error occurred which must abort processing.
    Dim intFile As Integer
    Dim ldLib As LibData
    Dim strLibKey As String
    Dim strSource As String
    Dim strSParts() As String
    Dim strMsg As String
    
    G_Package_IniDOM.Sections.Add "Dependencies"
    
    With frmLog
        intFile = FreeFile()
        Open G_VBP_FQFileName For Input As #intFile
        Do While Not EOF(intFile)
            Line Input #intFile, strSource
            strSParts = Split(strSource, "=", 2)
            If UBound(strSParts) > 0 Then
                strLibKey = UCase$(strSParts(0))
                Select Case strLibKey
                    Case "REFERENCE", "OBJECT"
                        Set ldLib = New LibData
                        With ldLib
                            If strLibKey = "REFERENCE" Then
                                If .LoadRefByProjFileRef(strSParts(1)) Then
                                    ProjectProcessFileScan = True
                                    strMsg = "Unregistered reference found in project: " _
                                           & strSParts(1)
                                    AddEvent sesFatal, strMsg
                                    frmLog.Log strMsg
                                    frmLog.Log
                                    Exit Do
                                End If
                            Else '"OBJECT"
                                If .LoadRefByProjFileObject(strSParts(1)) Then
                                    ProjectProcessFileScan = True
                                    strMsg = "Unregistered control reference found in project: " _
                                           & strSParts(1)
                                    AddEvent sesFatal, strMsg
                                    frmLog.Log strMsg
                                    frmLog.Log
                                    Exit Do
                                End If
                            End If
                            
                            'Test for Lib in collection already by adding it.
                            On Error Resume Next
                            G_ProjLibData.Add ldLib, .Name
                            If Err.Number = 0 Then
                                On Error GoTo 0
                                
                                'New dependency found.  Record settings.
                                With G_Package_IniDOM.Sections
                                    'Add library to list.
                                    !Dependencies.Keys.Add ldLib.Name
                                    
                                    'Add and populate dependency's Section.
                                    .Add "D:" & ldLib.Name
                                    With .Item("D:" & ldLib.Name).Keys
                                        .Add "Source", ldLib.Source
                                        .Add "Included", CStr(ldLib.Included)
                                        .Add "IncludedNever", CStr(ldLib.IncludedNever)
                                        .Add "FileLocation", ldLib.FileLocation, QuoteValue:=True
                                    End With
                                    .Add 'Blank line, Section level.
                                End With
                                
                                'Process any DEP file (recurses).
                                ProjectProcessDepFile .FileLocation
                            Else
                                On Error GoTo 0
                            End If
                        End With
                        
                    Case "NAME"
                        G_ProjName = DQ(strSParts(1))
                        
                    Case "MAJORVER"
                        G_ProjMajorVer = strSParts(1)
                        
                    Case "MINORVER"
                        G_ProjMinorVer = strSParts(1)
                        
                    Case "REVISIONVER"
                        G_ProjRevisionVer = strSParts(1)
                        
                    Case "DESCRIPTION"
                        G_ProjDescription = DQ(strSParts(1))
                        
                    Case "VERSIONCOMPANYNAME"
                        G_ProjCompany = DQ(strSParts(1))
                        
                    Case "VERSIONFILEDESCRIPTION"
                        G_ProjFileDescription = DQ(strSParts(1))
                        
                    Case "EXENAME32"
                        G_ProjEXE = DQ(strSParts(1))
                    
                    Case "PATH32"
                        G_ProjPath32 = DQ(strSParts(1)) & "\"
                End Select
            End If
        Loop
        Close #intFile
        G_Package_IniDOM.Sections!Dependencies.Keys.Add 'Blank line.
    End With
    
    G_ProjVersion = G_ProjMajorVer & "." _
                     & G_ProjMinorVer & ".0." _
                     & G_ProjRevisionVer
    G_Package_IniDOM!VBProject.Keys.Add "Version", G_ProjVersion
End Function

Private Sub ProjectProcessDepFile(ByVal LibFileLocation As String)
    'Process DEP file if any at LibFileLocation.
    Dim idDep As IniDOM
    Dim ikUses As IniKey
    Dim iksLib As IniKeys
    Dim ldLib As LibData
    Dim lngProbe As Long
    Dim strUses As String
    Dim strDependency As String

    Set idDep = IniDOMFromFile(DepFileName(LibFileLocation))
    If Not (idDep Is Nothing) Then
        'DEP file exists, has been parsed into an IniDOM.
        'Now test for this library's Section in the IniDOM, which
        'will be named for the library's simple filename with or
        'without an LCID suffix.
        For lngProbe = glsFirst To glsLast
            On Error Resume Next
            Set iksLib = _
                idDep.Sections.Item(SimpleFileName(LibFileLocation) & GetDepLCID(lngProbe)).Keys
            If Err.Number = 0 Then
                On Error GoTo 0
                For Each ikUses In iksLib
                    If ikUses.Name Like "Uses#*" Then
                        strUses = ikUses.Value
                        If Len(strUses) > 0 Then
                            'We have a dependency's simple filename.  See if we can
                            'locate this library.
                            strDependency = DLLSearch(PathOfFQFileName(LibFileLocation), _
                                            G_VBP_Path, _
                                            strUses)
                            If Len(strDependency) > 0 Then
                                Set ldLib = New LibData
                                With ldLib
                                    If .LoadDepRefByFileName(strDependency) Then
                                        'Warn unless we're not going to use it anyway.
                                        G_UnregWarning = Not .IncludedNever
                                    End If
                                    
                                    'Test for Lib in collection already by adding it.
                                    On Error Resume Next
                                    G_ProjLibData.Add ldLib, .Name
                                    If Err.Number = 0 Then
                                        On Error GoTo 0
                                    
                                        'New dependency found.  Record settings.
                                        With G_Package_IniDOM.Sections
                                            'Add library to list.
                                            !Dependencies.Keys.Add ldLib.Name
                                            
                                            'Add and populate dependency's Section.
                                            .Add "D:" & ldLib.Name
                                            With .Item("D:" & ldLib.Name).Keys
                                                .Add "Source", ldLib.Source
                                                .Add "Included", CStr(ldLib.Included)
                                                .Add "IncludedNever", CStr(ldLib.IncludedNever)
                                                .Add "FileLocation", ldLib.FileLocation, _
                                                     QuoteValue:=True
                                            End With
                                            .Add 'Blank line, Section level.
                                        End With
                                        
                                        'Process any DEP file (recurses).
                                        ProjectProcessDepFile .FileLocation
                                    Else
                                        On Error GoTo 0
                                    End If
                                End With
                            End If
                            'Else we can't find the dependency library for this "Uses" key.
                        End If
                        'Else an empty "Uses" entry, sometimes used to end the list (skip it).
                    End If
                    'Else not a "Uses" key and we can skip it.
                Next
                
                Exit For 'We found the Section.
            End If
            'Else DEP file doesn't have dependency entries for this Section name.
            'Loop to try again or fall through if we've exhausted the possibilities.
        Next
    End If
    'Else no DEP file.
End Sub

Private Function ProjectProcessLog() As Boolean
    'Returns True if an error occurred which must abort processing.
    Dim ldLib As LibData
    
    With frmLog
        'Log project properties.
        With frmProject
            .AddProp GPP_PROJNAME, G_ProjName
            .AddProp GPP_PROJCOMPANY, _
                     IIf(Len(G_ProjCompany) > 0, G_ProjCompany, """""")
            .AddProp GPP_PROJDESC, _
                     IIf(Len(G_ProjDescription) > 0, G_ProjDescription, """""")
            .AddProp GPP_PROJFILEDESC, _
                     IIf(Len(G_ProjFileDescription) > 0, G_ProjFileDescription, """""")
            .AddProp GPP_PROJEXE, G_ProjEXE
            .AddProp GPP_PROJPATH32, _
                     IIf(Len(G_ProjPath32) > 0, G_ProjPath32, """""")
            .AddProp GPP_PROJVERS, G_ProjVersion
        End With
        .Log "Project:"
        .Log "  Name = " & G_ProjName
        If Len(G_ProjCompany) > 0 Then
            .Log "  Company = " & G_ProjCompany
        End If
        If Len(G_ProjDescription) > 0 Then
            .Log "  Description = " & G_ProjDescription
        End If
        If Len(G_ProjFileDescription) > 0 Then
            .Log "  File description = " & G_ProjFileDescription
        End If
        .Log "  EXE name = " & G_ProjEXE
        .Log "  EXE path = " & G_ProjPath32
        .Log "  Version = " & G_ProjVersion
        .Log
        
        'View frmScan now.
        With frmMain
            .mnuViewScan.Enabled = True
            With .tbrMain.Buttons(GC_TBBTN_SCAN)
                .Enabled = True
                .Value = tbrPressed
            End With
        End With
        Show
        ZOrder vbBringToFront
        
        'Log libraries.
        frmProject.AddProp GPP_DEPLIBS, CStr(G_ProjLibData.Count)
        .Log "Libraries: " & CStr(G_ProjLibData.Count)
        .Log
        For Each ldLib In G_ProjLibData
            If LogLibrary(ldLib) Then ProjectProcessLog = True
        Next
        frmProject.AddProp GPP_PERMEXCLUDED, G_PermExcludedCount
    End With
End Function

'----- Event Handlers -----

Private Sub cmdBack_Click()
    frmMain.mnuViewProject_Click
End Sub

Private Sub cmdCancel_Click()
    If G_ExitMode Then
        Unload frmMain
    Else
        G_ExitMode = True
        UICancelCleanup
        cmdNext.Enabled = False
        frmMain.sbMain.SimpleText = "Project processing canceled."
    End If
End Sub

Private Sub cmdNext_Click()
    If blnPhaseComplete Then
        frmMain.mnuViewDependencies_Click
    Else
        blnPhaseComplete = True
        Load frmDeps
    End If
End Sub

Private Sub Form_GotFocus()
    cmdNext.SetFocus
End Sub

Private Sub Form_Load()
    FontWiz.AdjustControls Me.Controls
    
    sngScanEventsHeightDelta = FORM_SCALEHEIGHT - (lvwScanEvents.Top + lvwScanEvents.Height)
    
    MousePointer = vbHourglass
    frmProject.cmdNext.Enabled = True
    tmrStartProcessing.Enabled = True 'Trigger the scan.
End Sub

Private Sub Form_Resize()
    Const WIDTH_DELTA As Single = 60 'Diff. between Form width and canvas width in lvwScanEvents.
    Dim sngVal As Single
        
    If WindowState <> vbMinimized Then
        If ScaleWidth - GC_CANCELLEFTDELTA > 0 _
           And ScaleHeight - (lvwScanEvents.Top + sngScanEventsHeightDelta) > 0 Then
            txtScanEvent.Width = ScaleWidth

            With lvwScanEvents
                .Width = ScaleWidth
                .Height = ScaleHeight - (.Top + sngScanEventsHeightDelta)
                .ColumnHeaders(SCANEVT_SUBITEM_EVENT + 1).Width = _
                        ScaleWidth - WIDTH_DELTA _
                     - .ColumnHeaders(SCANEVT_SUBITEM_SEVERITY + 1).Width
            End With

            sngVal = ScaleHeight - GC_BUTTONSTOPDELTA
            cmdBack.Top = sngVal
            cmdBack.Left = ScaleWidth - GC_BACKLEFTDELTA
            cmdCancel.Top = sngVal
            cmdCancel.Left = ScaleWidth - GC_CANCELLEFTDELTA
            cmdNext.Top = sngVal
            cmdNext.Left = ScaleWidth - GC_NEXTLEFTDELTA
        End If
    End If
End Sub


Private Sub lvwScanEvents_BeforeLabelEdit(Cancel As Integer)
    'Disallow editing of label (Text property) values.
    
    Cancel = 1
End Sub

Private Sub lvwScanEvents_Click()
    txtScanEvent.Text = Replace$(lvwScanEvents.SelectedItem.SubItems(SCANEVT_SUBITEM_EVENT), _
                                 "·", vbNewLine)
End Sub

Private Sub tmrStartProcessing_Timer()
    tmrStartProcessing.Enabled = False
    
    If ProjectProcess() Then
        'Some error was encountered.
        With frmLog
            .Log "Fatal:"
            .Log "  Can't build XCopy package.  A serious problem was encountered"
            .Log "  while processing your project's dependencies.  Examine the log"
            .Log "  entries above for the reason(s)."
        End With
        MsgBox "Can't build XCopy package.  A serious problem was encountered" _
             & vbNewLine _
             & "while processing your project's dependencies.  Examine the log" _
             & vbNewLine _
             & "or Project scan for the reason(s).", _
               vbOKOnly Or vbCritical, _
               "MMM - Dependency processing problem"
            
        UICancelCleanup
        frmMain.sbMain.SimpleText = "VB 6.0 project scan failed.  See log."
    Else
        'Verify EXE name in project file.
        If Len(G_ProjEXE) = 0 Then
            With frmLog
                .Log "Fatal:"
                .Log "  Can't build XCopy package.  No EXE name found in project file."
                .Log "  Compile your program, save the project, and try again."
            End With
            MsgBox "Can't build XCopy package.  No EXE name found in project file." _
                 & vbNewLine _
                 & "Compile your program and try again.", _
                   vbOKOnly Or vbCritical, _
                   "MMM - No EXE name in project file"
            
            UICancelCleanup
            frmMain.sbMain.SimpleText = "VB 6.0 project scan failed.  See log."
        Else
            'Verify EXE present.
            If Left$(G_ProjPath32, 2) = "\\" Or Mid$(G_ProjPath32, 2, 2) = ":\" Then
                G_ProjEXE_FQFileName = G_ProjPath32 & G_ProjEXE
            Else
                G_ProjEXE_FQFileName = GetFullPath(G_VBP_Path & G_ProjPath32 & G_ProjEXE)
            End If
            If Not FilePresent(G_ProjEXE_FQFileName) Then
                With frmLog
                    .Log "Fatal:"
                    .Log "  Can't build XCopy package.  No EXE file found in project folder."
                    .Log "  Compile your program and try again."
                End With
                MsgBox "Can't build XCopy package.  No EXE file found in project folder." _
                     & vbNewLine _
                     & "Compile your program and try again.", _
                       vbOKOnly Or vbCritical, _
                       "MMM - No EXE file in project folder"
                
                UICancelCleanup
                frmMain.sbMain.SimpleText = "No compiled EXE to deploy.  See log."
            Else
                'Good scan!
                With lvwScanEvents
                    If .ListItems.Count > 0 Then
                        Set .SelectedItem = .ListItems(1)
                        lvwScanEvents_Click
                    End If
                End With
                cmdBack.Enabled = True
                cmdCancel.Enabled = True
                cmdNext.Enabled = True
                
                frmMain.sbMain.SimpleText = "VB 6.0 project scan complete."
            End If
        End If
    End If
        
    MousePointer = vbDefault
End Sub

