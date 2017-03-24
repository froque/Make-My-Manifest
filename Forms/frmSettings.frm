VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Package settings"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   ControlBox      =   0   'False
   LinkTopic       =   "frmSettings"
   MDIChild        =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   5895
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Height          =   375
      Left            =   2580
      TabIndex        =   20
      Top             =   5460
      Width           =   975
   End
   Begin VB.CheckBox chkSaveLog 
      Caption         =   "Save log to disk when saving package settings"
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Top             =   1320
      Value           =   1  'Checked
      Width           =   4575
   End
   Begin VB.Frame fraManifest 
      Caption         =   "Manifest settings"
      Height          =   3615
      Left            =   60
      TabIndex        =   25
      Top             =   1680
      Width           =   5775
      Begin VB.Frame fraVista 
         Height          =   1455
         Left            =   240
         TabIndex        =   27
         Top             =   2040
         Visible         =   0   'False
         Width           =   5295
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   915
            Left            =   540
            ScaleHeight     =   915
            ScaleWidth      =   2295
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   420
            Width           =   2295
            Begin VB.OptionButton optLevel 
               Caption         =   "asInvoker"
               Height          =   240
               Index           =   0
               Left            =   60
               TabIndex        =   14
               Top             =   60
               Value           =   -1  'True
               Width           =   1395
            End
            Begin VB.OptionButton optLevel 
               Caption         =   "highestAvailable"
               Height          =   240
               Index           =   1
               Left            =   60
               TabIndex        =   15
               Top             =   360
               Width           =   2115
            End
            Begin VB.OptionButton optLevel 
               Caption         =   "requireAdministrator"
               Height          =   240
               Index           =   2
               Left            =   60
               TabIndex        =   16
               Top             =   660
               Width           =   2115
            End
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   795
            Left            =   3600
            ScaleHeight     =   795
            ScaleWidth      =   1575
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   420
            Width           =   1575
            Begin VB.OptionButton optUIAccess 
               Caption         =   "false"
               Height          =   240
               Index           =   0
               Left            =   60
               TabIndex        =   17
               Top             =   60
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.OptionButton optUIAccess 
               Caption         =   "true"
               Height          =   240
               Index           =   1
               Left            =   60
               TabIndex        =   18
               Top             =   360
               Width           =   675
            End
            Begin VB.CommandButton cmdUIAccessTrueInfo 
               Caption         =   "i"
               CausesValidation=   0   'False
               BeginProperty Font 
                  Name            =   "Webdings"
                  Size            =   9.75
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1020
               TabIndex        =   19
               ToolTipText     =   "Click for cautions"
               Top             =   300
               Visible         =   0   'False
               Width           =   315
            End
         End
         Begin VB.Label Label4 
            Caption         =   "UI Access"
            ForeColor       =   &H80000011&
            Height          =   315
            Left            =   3480
            TabIndex        =   31
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Execution Level"
            ForeColor       =   &H80000011&
            Height          =   315
            Left            =   420
            TabIndex        =   30
            Top             =   180
            Width           =   1395
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   60
         ScaleHeight     =   1815
         ScaleWidth      =   5475
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   5475
         Begin VB.CommandButton cmdCompatibilityInfo 
            Caption         =   "i"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4800
            TabIndex        =   11
            ToolTipText     =   "Click for cautions"
            Top             =   1080
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CheckBox chkCompatibility 
            Caption         =   "Include Compatibility section in manifest"
            Height          =   315
            Left            =   60
            TabIndex        =   10
            Top             =   1080
            Width           =   4455
         End
         Begin VB.CheckBox chkDPIAware 
            Caption         =   "Include dpiAware section in manifest"
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   720
            Width           =   4455
         End
         Begin VB.CommandButton cmdDPIAwareInfo 
            Caption         =   "i"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4800
            TabIndex        =   9
            ToolTipText     =   "Click for cautions"
            Top             =   720
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton cmdTrustInfo 
            Caption         =   "i"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4800
            TabIndex        =   13
            ToolTipText     =   "Click for cautions"
            Top             =   1440
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton cmdKB828629Info 
            Caption         =   "i"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4800
            TabIndex        =   7
            ToolTipText     =   "Click for cautions"
            Top             =   360
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CheckBox chkKB828629 
            Caption         =   "Include VB 6.0 KB 828629 remediation"
            Height          =   315
            Left            =   60
            TabIndex        =   6
            Top             =   360
            Width           =   4455
         End
         Begin VB.CheckBox chkTrustInfo 
            Caption         =   "Include trustInfo section in manifest"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   1440
            Width           =   4455
         End
         Begin VB.CommandButton cmdCC6Info 
            Caption         =   "i"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4800
            TabIndex        =   5
            ToolTipText     =   "Click for cautions"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CheckBox chkCC6 
            Caption         =   "Include Common Controls 6.0 in manifest"
            Height          =   315
            Left            =   60
            TabIndex        =   4
            Top             =   0
            Width           =   4455
         End
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   21
      Top             =   5460
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4860
      TabIndex        =   22
      Top             =   5460
      Width           =   975
   End
   Begin VB.TextBox txtXCopy 
      Height          =   360
      Left            =   2820
      TabIndex        =   0
      Text            =   "MMMPack"
      Top             =   97
      Width           =   2955
   End
   Begin VB.TextBox txtDeps 
      Height          =   360
      Left            =   2820
      TabIndex        =   1
      Top             =   517
      Width           =   2955
   End
   Begin VB.CheckBox chkEmbedMan 
      Caption         =   "Embed application manifest in EXE"
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label lblXCopy 
      Alignment       =   1  'Right Justify
      Caption         =   "Package folder (within project):"
      Height          =   255
      Left            =   60
      TabIndex        =   24
      Top             =   120
      Width           =   2715
   End
   Begin VB.Label lblDeps 
      Alignment       =   1  'Right Justify
      Caption         =   "Dependencies folder (or none):"
      Height          =   255
      Left            =   60
      TabIndex        =   23
      Top             =   540
      Width           =   2715
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----- Private Data -----

'Local string constants.
Private Const VISTA_EXEC_LEVEL As String = _
    "asInvoker%highestAvailable%requireAdministrator"

'Misc.
Private blnPhaseComplete As Boolean 'Already completed this phase.

'----- Public Methods -----

Public Function ExecutionLevel() As String
    Dim intOption As Integer
    
    For intOption = 0 To optLevel.UBound
        If optLevel(intOption).Value Then Exit For
    Next
    ExecutionLevel = Split(VISTA_EXEC_LEVEL, "%")(intOption)
End Function

Public Sub UICancelCleanup()
    frmAddedFiles.UICancelCleanup
    cmdCancel.Caption = "E&xit"
End Sub

'----- Private Methods -----

'----- Event Handlers -----

Private Sub chkCC6_Click()
    cmdCC6Info.Visible = chkCC6.Value = vbChecked
End Sub

Private Sub chkCompatibility_Click()
    cmdCompatibilityInfo.Visible = chkCompatibility.Value = vbChecked
End Sub

Private Sub chkKB828629_Click()
    cmdKB828629Info.Visible = chkKB828629.Value = vbChecked
End Sub

Private Sub chkTrustInfo_Click()
    Dim blnVistaVis As Boolean
    
    blnVistaVis = chkTrustInfo.Value = vbChecked
    cmdTrustInfo.Visible = blnVistaVis
    fraVista.Visible = blnVistaVis
End Sub

Private Sub chkDPIAware_Click()
    cmdDPIAwareInfo.Visible = chkDPIAware.Value = vbChecked
End Sub

Private Sub cmdBack_Click()
    frmMain.mnuViewAddedFiles_Click
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

Private Sub cmdCC6Info_Click()
    MsgBox "You must ensure that you call InitCommonControls() " _
         & "properly to avoid causing your program to fail to " _
         & "start.", _
           vbOKOnly Or vbInformation, _
           "MMM - InitCommonControls info"
End Sub

Private Sub cmdCompatibilityInfo_Click()
    MsgBox "Include this option to select ""Windows 7"" behaviors under Windows 7, " _
         & "Windows 2010 Server, or later.  Omit this option to allow your program " _
         & "to default to ""Vista"" behaviors under these post-Vista OSs." _
         & vbNewLine & vbNewLine _
         & "The presence of this option is ignored by Vista and earlier OSs.", _
           vbOKOnly Or vbInformation, _
           "MMM - Compatibility info"
End Sub

Private Sub cmdDPIAwareInfo_Click()
    MsgBox "This is only used to mark your program as DPI Aware in Vista or later. " _
         & "Your program must be written to handle DPI Awareness.", _
           vbOKOnly Or vbInformation, _
           "MMM - DPI Aware info"
End Sub

Private Sub cmdNext_Click()
    If blnPhaseComplete Then
        frmMain.mnuViewMake_Click
    Else
        blnPhaseComplete = True
        With G_Package_IniDOM!MMMPackage.Keys
            .Add "Folder", txtXCopy.Text
            .Add "DepsFolder", txtDeps.Text
            .Add "EmbedManifest", CStr(chkEmbedMan.Value = vbChecked)
            .Add "SaveLog", CStr(chkSaveLog.Value = vbChecked)
            .Add "CC6", CStr(chkCC6.Value = vbChecked)
            .Add "KB828629", CStr(chkKB828629.Value = vbChecked)
            .Add "DPIAware", CStr(chkDPIAware.Value = vbChecked)
            .Add "Compatibility", CStr(chkCompatibility.Value = vbChecked)
            .Add "TrustInfo", CStr(chkTrustInfo.Value = vbChecked)
            .Add "ExecutionLevel", ExecutionLevel()
            .Add "UIAccess", CStr(optUIAccess(1).Value)
        End With
        Load frmMake
    End If
End Sub

Private Sub cmdKB828629Info_Click()
    MsgBox "This option addresses the ""KB 828629 issue.""  " _
         & "Programs using this information in their manifests " _
         & "require the VB6 SP6 runtimes and XP SP2 or later." & vbNewLine & vbNewLine _
         & "Microsoft Visual Basic 6.0 ActiveX controls are " _
         & "essentially COM DLL modules with .OCX file name " _
         & "extensions. If you try to configure these modules " _
         & "for SxS operation in Windows XP, you receive the " _
         & "following error message:" & vbNewLine & vbNewLine _
         & vbTab & "Runtime Error '336' Component not correctly " _
         & "registered." & vbNewLine & vbNewLine _
         & "Please see Microsoft KB article 828629 for details.", _
           vbOKOnly Or vbInformation, _
           "MMM - KB 828629 remediation info"
End Sub

Private Sub cmdTrustInfo_Click()
    MsgBox "Including this in the manifest will bypass Application Compatibility in Windows " _
         & "Vista or later for legacy applications.  This might be undesireable." & vbNewLine & vbNewLine _
         & "Also, if you do NOT include this information it is possible for " _
         & "your application to be erroneously detected as an installer or other " _
         & "type of program that requires a form of elevation.  This may result " _
         & "in an undesireable prompt, and also may cause your program " _
         & "to run in unintended ways.", _
           vbOKOnly Or vbInformation, _
           "MMM - Manifest trustinfo"
End Sub

Private Sub cmdUIAccessTrueInfo_Click()
    MsgBox "Programs using this option in their manifests must be " _
         & "signed.  In addition, the application must reside in " _
         & "a protected location in the file system. \Program Files\ " _
         & "and \windows\system32\ are currently the two allowable " _
         & "protected locations." & vbNewLine & vbNewLine _
         & "When true, the application is allowed to bypass user " _
         & "interface control levels to drive input to higher " _
         & "privilege windows on the desktop. This setting should " _
         & "only be used for user interface Assistive Technology " _
         & "applications.", _
           vbOKOnly Or vbInformation, _
           "MMM - Vista UI Access info"
End Sub

Private Sub Form_GotFocus()
    cmdNext.SetFocus
End Sub

Private Sub Form_Load()
    Dim sctManTool As IniSection
    
    FontWiz.AdjustControls Me.Controls
    
    On Error Resume Next
    Set sctManTool = G_Product_IniDOM!Manifest
    If Err.Number = 0 Then
        'We have a ManifestTool Section.
        'On Error GoTo 0
        
        'On Error Resume Next
        chkEmbedMan.Value = IIf(CBool(sctManTool!EmbedDefault.Value), _
                                vbChecked, _
                                vbUnchecked)
    End If
    On Error GoTo 0

    'View frmSettings now.
    With frmMain
        .mnuViewSettings.Enabled = True
        With .tbrMain.Buttons(GC_TBBTN_SETTINGS)
            .Enabled = True
            .Value = tbrPressed
        End With
        .sbMain.SimpleText = "Accepting package settings."
    End With
    Show
    ZOrder vbBringToFront
    cmdNext.SetFocus
End Sub

Private Sub Form_Resize()
    Dim sngVal As Single
    
    If WindowState <> vbMinimized Then
        sngVal = ScaleHeight - GC_BUTTONSTOPDELTA
        If ScaleWidth - GC_CANCELLEFTDELTA > 0 And sngVal > 0 Then
            cmdBack.Top = sngVal
            cmdBack.Left = ScaleWidth - GC_BACKLEFTDELTA
            cmdNext.Top = sngVal
            cmdNext.Left = ScaleWidth - GC_NEXTLEFTDELTA
            cmdCancel.Top = sngVal
            cmdCancel.Left = ScaleWidth - GC_CANCELLEFTDELTA
        End If
    End If
End Sub

Private Sub optUIAccess_Click(Index As Integer)
    cmdUIAccessTrueInfo.Visible = optUIAccess(1).Value = True
End Sub

Private Sub txtDeps_Validate(Cancel As Boolean)
    With txtDeps
        .Text = Trim$(.Text)
        If InStr(.Text, "\") > 0 Then
            MsgBox "Sorry, the Dependencies folder must be a simple folder name " _
                 & "that will go within the XCopy folder, or blank if you wish " _
                 & "to place your dependent libraries into the XCopy folder itself.", _
                 vbOKOnly Or vbExclamation, _
                 "MMM - Correct the Dependencies folder name"
            .SelStart = 0
            .SelLength = Len(.Text)
            Cancel = True
        End If
    End With
End Sub

Private Sub txtXCopy_Validate(Cancel As Boolean)
    With txtXCopy
        .Text = Trim$(.Text)
        If Len(.Text) = 0 Or InStr(.Text, "\") > 0 Then
            MsgBox "Sorry, the XCopy folder must be a simple folder name " _
                 & "that will go within the Project folder.", _
                 vbOKOnly Or vbExclamation, _
                 "MMM - Correct the XCopy folder name"
            .SelStart = 0
            .SelLength = Len(.Text)
            Cancel = True
        End If
    End With
End Sub

