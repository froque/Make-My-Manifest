VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProject 
   Caption         =   "Project properties"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "frmProject"
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   6975
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5940
      TabIndex        =   1
      Top             =   4680
      Width           =   975
   End
   Begin MSComctlLib.ListView lvwProjProps 
      CausesValidation=   0   'False
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Property"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----- Private Data -----

'Form resize events.
Private Const FORM_SCALEHEIGHT As Single = 5115# 'UPDATE THESE IF FORM DESIGN LAYOUT CHANGES

Private sngProjPropsHeightDelta As Single

'lvwProjProps Subitems.
Private Const PROP_SUBITEM_NAME = 0
Private Const PROP_SUBITEM_VALUE = 1

'----- Public Methods -----

Public Sub AddProp(ByVal Name As String, ByVal Value As String)
    Dim itmProp As ListItem
    
    Set itmProp = lvwProjProps.ListItems.Add(Key:=Name)
    With itmProp
        .Text = Name
        .Bold = True
        .SubItems(PROP_SUBITEM_VALUE) = Value
        .Selected = False
    End With
End Sub

Public Sub UICancelCleanup()
    cmdCancel.Caption = "E&xit"
End Sub

Public Sub UpdateProp(ByVal Name As String, ByVal NewValue As String)
    Dim itmProp As ListItem
    
    On Error Resume Next
    Set itmProp = lvwProjProps.ListItems.Item(Name)
    If Err.Number = 0 Then
        On Error GoTo 0
        With itmProp
            .SubItems(PROP_SUBITEM_VALUE) = NewValue
            .Selected = False
        End With
    Else
        On Error GoTo 0
        AddProp Name, NewValue
    End If
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
    frmMain.mnuViewScan_Click
End Sub

'----- Event Handlers -----

Private Sub Form_GotFocus()
    cmdNext.SetFocus
End Sub

Private Sub Form_Load()
    FontWiz.AdjustControls Me.Controls
    
    sngProjPropsHeightDelta = FORM_SCALEHEIGHT - (lvwProjProps.Top + lvwProjProps.Height)
End Sub

Private Sub Form_Resize()
    Const WIDTH_DELTA As Single = 60 'Diff. between Form width and canvas width in lvwProjProps.
    Dim sngVal As Single
        

    If WindowState <> vbMinimized Then
        sngVal = ScaleHeight - GC_BUTTONSTOPDELTA
        If ScaleWidth - GC_CANCELLEFTDELTA > 0 And sngVal > 0 Then
            With lvwProjProps
                .Width = ScaleWidth
                .Height = ScaleHeight - (.Top + sngProjPropsHeightDelta)
                .ColumnHeaders(PROP_SUBITEM_VALUE + 1).Width = _
                        ScaleWidth - WIDTH_DELTA _
                     - .ColumnHeaders(PROP_SUBITEM_NAME + 1).Width
            End With

            cmdCancel.Top = sngVal
            cmdCancel.Left = ScaleWidth - GC_CANCELLEFTDELTA
            cmdNext.Top = sngVal
            cmdNext.Left = ScaleWidth - GC_NEXTLEFTDELTA
        End If
    End If
End Sub

Private Sub lvwProjProps_BeforeLabelEdit(Cancel As Integer)
    'Disallow editing of label (Text property) values.
    
    Cancel = 1
End Sub

