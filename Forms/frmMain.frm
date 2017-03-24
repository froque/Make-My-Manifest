VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "MMM"
   ClientHeight    =   5685
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   8625
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   5160
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   5400
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   503
      Style           =   1
      SimpleText      =   "Idle"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open VB 6.0 project..."
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OpenVBP"
                  Text            =   "Open VB 6.0 project..."
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OpenMMMP"
                  Text            =   "Open MMM project..."
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   200
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Log"
            Object.ToolTipText     =   "MMM Activity Log"
            ImageIndex      =   9
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Project"
            Object.ToolTipText     =   "Project properties"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Scan"
            Object.ToolTipText     =   "Project scan report"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Dependencies"
            Object.ToolTipText     =   "Project dependencies"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AddedFiles"
            Object.ToolTipText     =   "Package added files"
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Settings"
            Object.ToolTipText     =   "Package settings"
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Make"
            Object.ToolTipText     =   "Make package"
            ImageIndex      =   8
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   200
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help contents..."
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4425
      Top             =   3690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":07C2
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08D4
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09E6
            Key             =   "Project"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B40
            Key             =   "Scan"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C9A
            Key             =   "Dependencies"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DF4
            Key             =   "AddedFiles"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F4E
            Key             =   "Settings"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10A8
            Key             =   "Make"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1202
            Key             =   "Log"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":135C
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpenVBP 
         Caption         =   "&Open VB 6.0 project..."
      End
      Begin VB.Menu mnuFileOpenMMMP 
         Caption         =   "Open &MMM project..."
      End
      Begin VB.Menu mnuFileDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit MMM"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewLog 
         Caption         =   "MMM activity &log"
      End
      Begin VB.Menu mnuViewProject 
         Caption         =   "Project &properties"
      End
      Begin VB.Menu mnuViewScan 
         Caption         =   "Project scan &report"
      End
      Begin VB.Menu mnuViewDependencies 
         Caption         =   "Project &dependencies"
      End
      Begin VB.Menu mnuViewAddedFiles 
         Caption         =   "Package &added files"
      End
      Begin VB.Menu mnuViewSettings 
         Caption         =   "Package &settings"
      End
      Begin VB.Menu mnuViewMake 
         Caption         =   "Ma&ke package"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----- Public Methods -----

'----- Private Methods -----

Private Function GetProjectFile() As String
    With dlgFile
        .CancelError = True
        .DialogTitle = "Select project file"
        .Filter = "VB project files (*.vbp)|*.vbp"
        .Flags = cdlOFNExplorer _
              Or cdlOFNFileMustExist _
              Or cdlOFNHideReadOnly _
              Or cdlOFNLongNames _
              Or cdlOFNPathMustExist _
              Or cdlOFNShareAware
        .InitDir = CurDir$()
        On Error Resume Next
        .ShowOpen
        If Err.Number <> 0 Then
            On Error GoTo 0
            GetProjectFile = ""
        Else
            On Error GoTo 0
            GetProjectFile = .FileName
            .FileName = ""
        End If
    End With
End Function

'----- Event Handlers ------

Private Sub MDIForm_Load()
    Dim intButton As Integer
    Dim intButtonMenu As Integer
    
    FontWiz.AdjustControls Me.Controls
    
    sbMain.SimpleText = "Idle"
    
    'Menu setup.
    mnuFileOpenMMMP.Enabled = False '@@@@@@@@ Future use of MMMPs.
    mnuViewProject.Enabled = False
    mnuViewScan.Enabled = False
    mnuViewDependencies.Enabled = False
    mnuViewAddedFiles.Enabled = False
    mnuViewSettings.Enabled = False
    mnuViewMake.Enabled = False
    mnuViewLog.Enabled = False
    mnuHelpContents.Enabled = False
    
    'Toolbar setup.
    With tbrMain
        For intButton = GC_TBBTN_LOG To GC_TBBTN_HELP
            .Buttons(intButton).Enabled = False
        Next
        .Buttons(GC_TBBTN_OPEN).ButtonMenus(2).Enabled = False '@@@@@ (Open MMMP) @@@@@@
    End With
End Sub

Private Sub MDIForm_Resize()
    If WindowState <> vbMinimized Then
        If Width < 7005 Then Width = 7005
        If Height < 7350 Then Height = 7350
        If WindowState <> vbMaximized Then
            If Top + Height > Screen.Height - 500 Then
                Top = Screen.Height - 500 - Height
            End If
        End If
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim frmClose As Form
    
    'Visually cleaner to close out the children first.
    For Each frmClose In Forms
        If frmClose.Name <> Name Then
            Unload frmClose
        End If
    Next
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOpenMMMP_Click()
    MsgBox "Not implemented" '@@@@@@@@@@@@@@@@@@@@@@@@@@@@
End Sub

Private Sub mnuFileOpenVBP_Click()
    G_VBP_FQFileName = GetProjectFile()
    
    If Len(G_VBP_FQFileName) > 0 Then
        Load frmLog
        mnuViewLog.Enabled = True
        tbrMain.Buttons(GC_TBBTN_LOG).Enabled = True
        
        With frmLog
            .Log G_Product_Identity
            .Log FormatDateTime(Date, vbLongDate) & " @ " & FormatDateTime(Time(), vbLongTime)
            .Log
            .Log "Host OS = Windows " & FormatNumber(AppEx.OSVersion(aioVersion), 2) _
               & " " & AppEx.OSVersion(aioServicePackString) _
               & " (Build " & CStr(AppEx.OSVersion(aioBuildNumber)) & ")"
            .Log "Host = " & AppEx.ComputerName
            .Log "User = " & AppEx.UserName
            .Log
            Set G_Product_IniDOM = _
                IniDOMFromFile(G_Product_Settings_Path & GC_PRODUCT_SETTINGS_FILE)
            If G_Product_IniDOM Is Nothing Then
                .Log "Missing settings file " & GC_PRODUCT_SETTINGS_FILE
                MsgBox "Missing settings file " & GC_PRODUCT_SETTINGS_FILE, _
                       vbOKOnly Or vbCritical
                mnuFileOpenVBP.Enabled = False
                mnuFileOpenMMMP.Enabled = False
                tbrMain.Buttons(GC_TBBTN_OPEN).Enabled = False
                sbMain.SimpleText = "Fatal error: missing " & GC_PRODUCT_SETTINGS_FILE
            Else
                G_VBP_Path = PathOfFQFileName(G_VBP_FQFileName)
                G_VBP_FileName = SimpleFileName(G_VBP_FQFileName)
                Caption = G_VBP_FileName & " - MMM: Make My Manifest"
                
                'Create MMM package settings DOM, begin populating Sections.
                Set G_Package_IniDOM = New IniDOM
                With G_Package_IniDOM.Sections
                    .Add , "MMM package settings."
                    .Add
                    .Add "MMM"
                    !MMM.Keys.Add "Version", CStr(App.Major) & "." _
                                           & CStr(App.Minor) & ".0." _
                                           & CStr(App.Revision)
                    .Add
                    IniDOMCloneSection "Redist", G_Product_IniDOM, G_Package_IniDOM
                    .Add
                    .Add "VBProject"
                    With !VBProject.Keys
                        .Add "File", G_VBP_FileName
                        .Add "Folder", G_VBP_Path
                    End With
                    .Add
                    .Add "MMMPackage"
                    .Add
                End With
                
                mnuViewProject.Enabled = True
                tbrMain.Buttons(GC_TBBTN_PROJECT).Enabled = True
                tbrMain.Buttons(GC_TBBTN_PROJECT).Value = tbrPressed
                With frmProject
                    .Show 'And Load.
                    .AddProp GPP_PROJPATH, G_VBP_Path
                    .AddProp GPP_VBPROJFILE, G_VBP_FileName
                End With
                .Log GPP_PROJPATH & " = " & G_VBP_Path
                .Log GPP_VBPROJFILE & " = " & G_VBP_FileName
                .Log
                
                sbMain.SimpleText = "Scanning VBP file and dependencies"
                
                'Disable Open options.
                Dim intButtonMenu As Integer
                
                With tbrMain.Buttons(GC_TBBTN_OPEN)
                    .Enabled = False
                    For intButtonMenu = 1 To 2
                        .ButtonMenus(intButtonMenu).Enabled = False
                    Next
                End With
                mnuFileOpenVBP.Enabled = False
                mnuFileOpenMMMP.Enabled = False
                
                'Start scanning.
                Load frmScan
            End If
        End With
    End If
    'Else canceled open.
End Sub

Private Sub mnuHelpAbout_Click()
    frmSplash.Show vbModal, Me
End Sub

Private Sub mnuHelpContents_Click()
    MsgBox "Not implemented" '@@@@@@@@@@@@@@@@@@@@@@@@@@@@
End Sub

Public Sub mnuViewAddedFiles_Click()
    tbrMain.Buttons!AddedFiles.Value = tbrPressed
    frmAddedFiles.ZOrder vbBringToFront
    With frmAddedFiles
        If .cmdNext.Visible Then If .cmdNext.Enabled Then .cmdNext.SetFocus
    End With
End Sub

Public Sub mnuViewDependencies_Click()
    tbrMain.Buttons!Dependencies.Value = tbrPressed
    frmDeps.ZOrder vbBringToFront
    With frmDeps
        If .cmdNext.Visible Then If .cmdNext.Enabled Then .cmdNext.SetFocus
    End With
End Sub

Private Sub mnuViewLog_Click()
    tbrMain.Buttons!Log.Value = tbrPressed
    frmLog.Show 'We don't present this until/unless requested.
    frmLog.ZOrder vbBringToFront
End Sub

Public Sub mnuViewMake_Click()
    tbrMain.Buttons!Make.Value = tbrPressed
    frmMake.ZOrder vbBringToFront
    With frmMake
        If .cmdFinish.Visible Then If .cmdFinish.Enabled Then .cmdFinish.SetFocus
    End With
End Sub

Public Sub mnuViewProject_Click()
    tbrMain.Buttons!Project.Value = tbrPressed
    frmProject.ZOrder vbBringToFront
    With frmProject
        If .cmdNext.Visible Then If .cmdNext.Enabled Then .cmdNext.SetFocus
    End With
End Sub

Public Sub mnuViewScan_Click()
    tbrMain.Buttons!Scan.Value = tbrPressed
    frmScan.ZOrder vbBringToFront
    With frmScan
        If .cmdNext.Visible Then If .cmdNext.Enabled Then .cmdNext.SetFocus
    End With
End Sub

Public Sub mnuViewSettings_Click()
    tbrMain.Buttons!Settings.Value = tbrPressed
    frmSettings.ZOrder vbBringToFront
    With frmSettings
        If .cmdNext.Visible Then If .cmdNext.Enabled Then .cmdNext.SetFocus
    End With
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Open"
            mnuFileOpenVBP_Click
        Case "Project"
            mnuViewProject_Click
        Case "Scan"
            mnuViewScan_Click
        Case "Dependencies"
            mnuViewDependencies_Click
        Case "AddedFiles"
            mnuViewAddedFiles_Click
        Case "Settings"
            mnuViewSettings_Click
        Case "Make"
            mnuViewMake_Click
        Case "Log"
            mnuViewLog_Click
        Case "Help"
            mnuHelpContents_Click
        Case Else
            MsgBox "Toolbar error (non-fatal).  Please report.  Details:" & vbNewLine _
                 & vbNewLine _
                 & G_Product_Identity & vbNewLine _
                 & "Button.Key was [" & Button.Key & "]", _
                   vbOKOnly Or vbExclamation
    End Select
End Sub

Private Sub tbrMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Parent.Key
        Case "Open"
            Select Case ButtonMenu.Key
                Case "OpenVBP"
                    mnuFileOpenVBP_Click
                Case "OpenMMMP"
                    mnuFileOpenMMMP_Click
            End Select
    End Select
End Sub
