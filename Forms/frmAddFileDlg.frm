VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddFileDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MMM - Add a non-code file"
   ClientHeight    =   3210
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmAddFileDlg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNote 
      Height          =   795
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1800
      Width           =   5355
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSubfolder 
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   1080
      Width           =   5355
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5580
      TabIndex        =   1
      ToolTipText     =   "Browse for file location"
      Top             =   360
      Width           =   315
   End
   Begin VB.TextBox txtFileLocation 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   5355
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4980
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Note:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1500
      Width           =   2955
   End
   Begin VB.Label Label3 
      Caption         =   "Package subfolder:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   780
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "File location:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   2955
   End
End
Attribute VB_Name = "frmAddFileDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----- Public Data -----

'----- Private Data -----

Private intFileID As Integer
Private strInitDir As String
Private strFileLocation As String

'----- Public Methods -----

'----- Private Methods -----

'----- Event Handlers -----

Private Sub cmdBrowse_Click()
    Dim strCD As String
    
    If Len(strInitDir) = 0 Then
        strInitDir = G_VBP_Path
        strInitDir = Left$(strInitDir, Len(strInitDir) - 1)
    End If
    
    With dlgFile
        .CancelError = True
        .DialogTitle = "Select a file to add"
        .Filter = "All files (*.*)|*.*"
        .Flags = cdlOFNExplorer _
              Or cdlOFNFileMustExist _
              Or cdlOFNHideReadOnly _
              Or cdlOFNLongNames _
              Or cdlOFNPathMustExist _
              Or cdlOFNShareAware
        .InitDir = strInitDir
        strCD = CurDir$()
        On Error Resume Next
        .ShowOpen
        If Err.Number = 0 Then
            On Error GoTo 0
            txtFileLocation.Text = .FileName
            txtFileLocation.SetFocus
            strInitDir = TrimSlash(PathOfFQFileName(.FileName))
            .FileName = ""
        Else
            On Error GoTo 0
        End If
        ChDrive strCD
        ChDir strCD
    End With
End Sub

Private Sub cmdCancel_Click()
    txtFileLocation.Text = ""
    txtSubfolder.Text = ""
    txtNote.Text = ""
    Hide
End Sub

Private Sub cmdOk_Click()
    If Len(strFileLocation) > 0 Then
        intFileID = intFileID + 1
        
        'Get Additions Section.
        Dim sctAdditions As IniSection
        On Error Resume Next
        Set sctAdditions = G_Package_IniDOM!Additions
        If Err.Number <> 0 Then
            'Section doesn't exist.
            On Error GoTo 0
            
            Set sctAdditions = G_Package_IniDOM.Sections.Add("Additions")
        
            'Add Section-level blank.
            G_Package_IniDOM.Sections.Add
        End If
        On Error GoTo 0
        
        'Add new Addition.
        sctAdditions.Keys.Add CStr(intFileID)
            
        'Add new addition Section.
        Dim sctAddition As IniSection
        Set sctAddition = G_Package_IniDOM.Sections.Add("A:" & CStr(intFileID))
        
        'Populate Keys.
        txtFileLocation.Text = Trim$(txtFileLocation.Text)
        txtSubfolder.Text = Trim$(txtSubfolder.Text)
        With sctAddition.Keys
            .Add "Included", "True"
            .Add "FileName", Mid$(txtFileLocation.Text, InStrRev(txtFileLocation.Text, "\") + 1)
            .Add "TargetFolder", txtSubfolder.Text
            .Add "SourceLocation", txtFileLocation.Text, QuoteValue:=True
            .Add "Note", txtNote.Text, QuoteValue:=True
        End With
        
        'Add Section-level blank.
        G_Package_IniDOM.Sections.Add
    End If
    cmdCancel_Click
End Sub

Private Sub Form_GotFocus()
    txtFileLocation.SetFocus
End Sub

Private Sub Form_Load()
    FontWiz.AdjustControls Me.Controls
End Sub

Private Sub txtFileLocation_Change()
    cmdOk.Enabled = Len(txtFileLocation.Text) > 0
End Sub

Private Sub txtFileLocation_Validate(Cancel As Boolean)
    With txtFileLocation
        If Len(.Text) > 0 Then
            If FilePresent(.Text) Then
                strFileLocation = .Text
            Else
                MsgBox "File doesn't exist.  Correct the file name" & vbNewLine _
                     & "or Cancel.", _
                       vbOKOnly Or vbExclamation, _
                       "MMM - Added file does not exist"
                .SelStart = 0
                .SelLength = Len(.Text)
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub txtNote_Validate(Cancel As Boolean)
    With txtNote
        If Len(.Text) > 0 Then
            .Text = Trim$(.Text)
            .Text = Replace$(.Text, """", "'") 'Can't have quotes!
        End If
    End With
End Sub

Private Sub txtSubfolder_Validate(Cancel As Boolean)
    Dim strCLSID As String
    
    With txtSubfolder
        .Text = Trim$(.Text)
        If InStr(.Text, "\") > 0 Then
            MsgBox "Sorry, the subfolder must be a simple folder name " _
                 & "that will go within the XCopy folder, or blank if you wish " _
                 & "to place this added file into the XCopy folder itself.", _
                 vbOKOnly Or vbExclamation, _
                 "MMM - Correct the subfolder name"
            .SelStart = 0
            .SelLength = Len(.Text)
            Cancel = True
        End If
    End With
End Sub
