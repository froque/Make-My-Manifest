VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAddedFiles 
   Caption         =   "Added files"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   ControlBox      =   0   'False
   LinkTopic       =   "frmAddedFiles"
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   6915
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   5700
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   5700
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Height          =   375
      Left            =   4740
      TabIndex        =   5
      Top             =   5700
      Width           =   975
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   1260
      TabIndex        =   3
      ToolTipText     =   "Remove selected added file"
      Top             =   5700
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Add a file to the package"
      Top             =   5700
      Width           =   975
   End
   Begin MSComctlLib.ListView lvwAdded 
      Height          =   5235
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   9234
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
         Text            =   "File ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Target folder"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Source location"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Note"
         Object.Width           =   10160
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Add non-code files to your package:"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4635
   End
End
Attribute VB_Name = "frmAddedFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----- Private Data -----

'Form resize events.
Private Const FORM_SCALEHEIGHT As Single = 6135# 'UPDATE THESE IF FORM DESIGN LAYOUT CHANGES
Private Const FORM_SCALEWIDTH As Single = 6915#

Private sngAddedHeightDelta As Single
Private sngAddedWidthDelta As Single

Private blnPhaseComplete As Boolean 'Already completed this phase.

'lvwAdded Subitems.
Private Const ADDED_SUBITEM_FILENAME = 1
Private Const ADDED_SUBITEM_TARGET = 2
Private Const ADDED_SUBITEM_SOURCELOCATION = 3
Private Const ADDED_SUBITEM_NOTE = 4

'----- Public Methods -----

Public Sub UICancelCleanup()
    frmDeps.UICancelCleanup
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Caption = "E&xit"
End Sub

'----- Private Methods -----

Private Sub PopulateAddedList()
    Dim sctAddeds As IniSection
    Dim kyAdded As IniKey
    Dim sctAdded As IniSection
    Dim itmAdded As ListItem
    
    On Error Resume Next
    Set sctAddeds = G_Package_IniDOM!Additions
    If Err.Number = 0 Then
        'We have entries!
        On Error GoTo 0
        For Each kyAdded In sctAddeds.Keys
            On Error Resume Next
            Set itmAdded = lvwAdded.ListItems("#" & kyAdded.Name) 'Can't be numeric, add #.
            If Err.Number <> 0 Then
                On Error GoTo 0
                Set itmAdded = lvwAdded.ListItems.Add(, "#" & kyAdded.Name)
                With itmAdded
                    .Text = kyAdded.Name
                    Set sctAdded = G_Package_IniDOM.Sections("A:" & kyAdded.Name)
                    .Checked = sctAdded!Included
                    .SubItems(ADDED_SUBITEM_FILENAME) = sctAdded!FileName
                    .SubItems(ADDED_SUBITEM_TARGET) = sctAdded!TargetFolder
                    .SubItems(ADDED_SUBITEM_SOURCELOCATION) = sctAdded!SourceLocation
                    .SubItems(ADDED_SUBITEM_NOTE) = sctAdded!Note
                    .Selected = False
                End With
            End If
            'Else exists in list already, skip it.
        Next
    End If
End Sub

'----- Event Handlers -----

Private Sub cmdAdd_Click()
    frmAddFileDlg.Show vbModal, frmMain
    PopulateAddedList
End Sub

Private Sub cmdBack_Click()
    frmMain.mnuViewDependencies_Click
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
        frmMain.mnuViewSettings_Click
    Else
        blnPhaseComplete = True
        Load frmSettings
    End If
End Sub

Private Sub cmdRemove_Click()
    Dim strAddedKey As String
    
    strAddedKey = lvwAdded.SelectedItem.Key
    lvwAdded.ListItems.Remove strAddedKey
    
    strAddedKey = Mid$(strAddedKey, 2) 'Strip the # prefix.
    With G_Package_IniDOM
        .Sections.Remove "A:" & strAddedKey
        !Additions.Keys.Remove strAddedKey
    End With
End Sub

Private Sub Form_GotFocus()
    If cmdNext.Enabled Then cmdNext.SetFocus
End Sub

Private Sub Form_Load()
    FontWiz.AdjustControls Me.Controls
    
    sngAddedHeightDelta = FORM_SCALEHEIGHT - lvwAdded.Height
    sngAddedWidthDelta = FORM_SCALEWIDTH - lvwAdded.Width

    'View frmAddedFiles now.
    With frmMain
        .mnuViewAddedFiles.Enabled = True
        With .tbrMain.Buttons(GC_TBBTN_ADDEDFILES)
            .Enabled = True
            .Value = tbrPressed
        End With
        .sbMain.SimpleText = "Accepting other package files."
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
            lvwAdded.Height = ScaleHeight - sngAddedHeightDelta
            lvwAdded.Width = ScaleWidth - sngAddedWidthDelta
            cmdAdd.Top = sngVal
            cmdRemove.Top = sngVal
            cmdBack.Top = sngVal
            cmdBack.Left = ScaleWidth - GC_BACKLEFTDELTA
            cmdNext.Top = sngVal
            cmdNext.Left = ScaleWidth - GC_NEXTLEFTDELTA
            cmdCancel.Top = sngVal
            cmdCancel.Left = ScaleWidth - GC_CANCELLEFTDELTA
        End If
    End If
End Sub

Private Sub lvwAdded_BeforeLabelEdit(Cancel As Integer)
    'Disallow editing of label (Text property) values.
    
    Cancel = 1
End Sub

Private Sub lvwAdded_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    'Capture new status for addition.
    Dim sctAdded As IniSection
    
    With Item
        Set sctAdded = G_Package_IniDOM.Sections("A:" & .Text)
        sctAdded!Included = CStr(.Checked)
    End With
End Sub
