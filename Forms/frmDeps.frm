VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDeps 
   Caption         =   "Dependencies"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   ControlBox      =   0   'False
   LinkTopic       =   "frmDeps"
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   6915
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   5700
      Width           =   975
   End
   Begin VB.CommandButton cmdUnregWarning 
      Caption         =   "i"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3060
      TabIndex        =   5
      ToolTipText     =   "Click for cautions"
      Top             =   5700
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   5700
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Height          =   375
      Left            =   4740
      TabIndex        =   7
      Top             =   5700
      Width           =   975
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   1260
      TabIndex        =   4
      ToolTipText     =   "Remove selected manual dependency"
      Top             =   5700
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Manually add a dependency"
      Top             =   5700
      Width           =   975
   End
   Begin MSComctlLib.ListView lvwDeps 
      Height          =   4935
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   8705
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Library"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Src*"
         Object.Width           =   1094
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Version"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Location"
         Object.Width           =   5080
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description"
         Object.Width           =   10160
      EndProperty
   End
   Begin VB.Label lblNotice 
      Caption         =   "* VBP - project file,  DEP - DEP file,  MAN - manual"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   5340
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Check the dependencies to isolate:"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4635
   End
End
Attribute VB_Name = "frmDeps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----- Private Data -----

'Form resize events.
Private Const FORM_SCALEHEIGHT As Single = 6135# 'UPDATE THESE IF FORM DESIGN LAYOUT CHANGES
Private Const FORM_SCALEWIDTH As Single = 6915#

Private sngDepsHeightDelta As Single
Private sngDepsWidthDelta As Single
Private sngNoticeTopDelta As Single

Private blnPhaseComplete As Boolean 'Already completed this phase, now we're just viewing.

'lvwDeps Subitems.
Private Const DEPS_SUBITEM_SOURCE = 1
Private Const DEPS_SUBITEM_VERSION = 2
Private Const DEPS_SUBITEM_FILELOCATION = 3
Private Const DEPS_SUBITEM_DESCRIPTION = 4
'Private Const DEPS_SUBITEM_NOTE = 5 '@@ Future use!

'----- Public Methods -----

Public Sub UnregWarning()
    If cmdUnregWarning.Visible Then Exit Sub
    UnregWarningContent
    cmdUnregWarning.Visible = True
End Sub

Public Sub UICancelCleanup()
    frmScan.UICancelCleanup
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Caption = "E&xit"
End Sub

'----- Private Methods -----

Private Sub PopulateDepsList()
    Dim ldLib As LibData
    Dim itmLib As ListItem
    
    For Each ldLib In G_ProjLibData
        On Error Resume Next
        Set itmLib = lvwDeps.ListItems(ldLib.Name)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Set itmLib = lvwDeps.ListItems.Add(, ldLib.Name)
            With itmLib
                .Text = ldLib.Name
                If ldLib.IncludedNever Then
                    .ForeColor = vbGrayText
                    .Ghosted = True 'Use as a "disabled" flag.
                    .ToolTipText = "Excluded by MMM"
                End If
                .Checked = ldLib.Included
                .SubItems(DEPS_SUBITEM_SOURCE) = ldLib.Source
                .SubItems(DEPS_SUBITEM_VERSION) = ldLib.Version
                .SubItems(DEPS_SUBITEM_FILELOCATION) = ldLib.FileLocation
                .SubItems(DEPS_SUBITEM_DESCRIPTION) = ldLib.Description
                .Selected = False
            End With
        End If
        'Else exists in list already, skip it.
    Next
End Sub

Private Sub UnregWarningContent()
    MsgBox "A standard or unregistered library has been included.  " _
         & "Inspect the list of dependencies here closely " _
         & "before deciding to proceed.", _
           vbOKOnly Or vbExclamation, _
           "MMM - Possible unregistered dependency"
End Sub

'----- Event Handlers -----

Private Sub cmdAdd_Click()
    frmAddDepDlg.Show vbModal, frmMain
    PopulateDepsList
End Sub

Private Sub cmdBack_Click()
    frmMain.mnuViewScan_Click
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
        frmMain.mnuViewAddedFiles_Click
    Else
        blnPhaseComplete = True
        Load frmAddedFiles
    End If
End Sub

Private Sub cmdRemove_Click()
    Dim strLibKey As String
    
    strLibKey = lvwDeps.SelectedItem.Key
    lvwDeps.ListItems.Remove strLibKey
    
    If G_ProjLibData(strLibKey).IncludedNever Then G_PermExcludedCount = G_PermExcludedCount - 1
    G_ProjLibData.Remove strLibKey
    
    With G_Package_IniDOM
        .Sections.Remove "D:" & strLibKey
        !Dependencies.Keys.Remove strLibKey
    End With
    
    'Log it here.
    With frmLog
        .Log "Manually removed:"
        .Log "  " & strLibKey
        .Log
    End With
    frmProject.UpdateProp GPP_PERMEXCLUDED, G_PermExcludedCount
End Sub

Private Sub cmdUnregWarning_Click()
    UnregWarningContent
End Sub

Private Sub Form_GotFocus()
    If cmdNext.Enabled Then cmdNext.SetFocus
End Sub

Private Sub Form_Load()
    FontWiz.AdjustControls Me.Controls

    sngDepsHeightDelta = FORM_SCALEHEIGHT - lvwDeps.Height
    sngDepsWidthDelta = FORM_SCALEWIDTH - lvwDeps.Width
    sngNoticeTopDelta = FORM_SCALEHEIGHT - lblNotice.Top

    PopulateDepsList

    'View frmDeps now.
    With frmMain
        .mnuViewDependencies.Enabled = True
        With .tbrMain.Buttons(GC_TBBTN_DEPS)
            .Enabled = True
            .Value = tbrPressed
        End With
        .sbMain.SimpleText = "VB 6.0 project dependencies cataloged." _
                           & IIf(G_UnregWarning, _
                                 "  Standard or unregistered libraries found.", _
                                 "")
    End With
    Show
    ZOrder vbBringToFront

    cmdNext.SetFocus
    If G_UnregWarning Then UnregWarning
End Sub

Private Sub Form_Resize()
    Dim sngVal As Single
    
    If WindowState <> vbMinimized Then
        sngVal = ScaleHeight - GC_BUTTONSTOPDELTA
        If ScaleWidth - GC_CANCELLEFTDELTA > 0 And sngVal > 0 Then
            lvwDeps.Height = ScaleHeight - sngDepsHeightDelta
            lvwDeps.Width = ScaleWidth - sngDepsWidthDelta
            lblNotice.Top = ScaleHeight - sngNoticeTopDelta
            cmdAdd.Top = sngVal
            cmdRemove.Top = sngVal
            cmdUnregWarning.Top = sngVal
            cmdUnregWarning.Left = ScaleWidth - GC_UNREGWARNLEFTDELTA
            cmdBack.Top = sngVal
            cmdBack.Left = ScaleWidth - GC_BACKLEFTDELTA
            cmdNext.Top = sngVal
            cmdNext.Left = ScaleWidth - GC_NEXTLEFTDELTA
            cmdCancel.Top = sngVal
            cmdCancel.Left = ScaleWidth - GC_CANCELLEFTDELTA
        End If
    End If
End Sub

Private Sub lvwDeps_BeforeLabelEdit(Cancel As Integer)
    'Disallow editing of label (Text property) values.
    
    Cancel = 1
End Sub

Private Sub lvwDeps_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    'Disallow checking of permanently excluded libraries,
    'capture new status for other libraries.
    Dim ldLib As LibData
    
    With Item
        If .Ghosted Then
            If .Checked Then .Checked = False
        Else
            Set ldLib = G_ProjLibData.Item(.Text)
            ldLib.Included = .Checked
        End If
    End With
End Sub

'@@ Future use, when we decide to persist MMM project settings
'@@ to the project folder.  This is part of a "dependency note"
'@@ implementation allowing the user to enter comments such as
'@@ why a dependency was excluded, etc.
'
'Private Sub lvwDeps_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    Dim strNote As String
'
'    With Item
'        strNote = InputBox("Edit note on this dependency:", _
'                           "MMM - Edit dependency note", _
'                           .SubItems(DEPS_SUBITEM_NOTE))
'        .SubItems(DEPS_SUBITEM_NOTE) = strNote
'    End With
'End Sub


