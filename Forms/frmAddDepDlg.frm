VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddDepDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MMM - Manually add an isolated dependency"
   ClientHeight    =   2595
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmAddDepDlg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   120
      Top             =   1980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtProgID 
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   1440
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
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Add by ProgID:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1140
      Width           =   2955
   End
   Begin VB.Label Label2 
      Caption         =   "- or -"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   780
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Add by file location:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   2955
   End
End
Attribute VB_Name = "frmAddDepDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----- Public Data -----

'----- Private Data -----

Private strInitDir As String
Private strFileLocation As String

'----- Public Methods -----

'----- Private Methods -----

'----- Event Handlers -----

Private Sub cmdBrowse_Click()
    Dim strCD As String
    
    If Len(strInitDir) = 0 Then
        strInitDir = GetSystemPath()
        strInitDir = Left$(strInitDir, Len(strInitDir) - 1)
    End If
    
    With dlgFile
        .CancelError = True
        .DialogTitle = "Select manual dependency"
        .Filter = "Dynamic-link Libraries (*.dll;*.ocx)|*.dll;*.ocx"
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
    txtProgID.Text = ""
    Hide
End Sub

Private Sub cmdOk_Click()
    If Len(strFileLocation) > 0 Then
        If frmScan.ManualRefByFile(strFileLocation) Then
            frmDeps.UnregWarning
        End If
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
    cmdOk.Enabled = Len(txtFileLocation.Text) > 0 Or Len(txtProgID) > 0
End Sub

Private Sub txtFileLocation_Validate(Cancel As Boolean)
    With txtFileLocation
        .Text = Trim$(.Text)
        If Len(.Text) > 0 Then
            If Len(txtProgID.Text) > 0 Then
                MsgBox "Can't Add by both file name and ProgID." & vbNewLine _
                     & "Clear one other the other, or Cancel.", _
                       vbOKOnly Or vbExclamation, _
                       "MMM - Can't Add by both file name and ProgID"
                .SelStart = 0
                .SelLength = Len(.Text)
                Cancel = True
            ElseIf FilePresent(.Text) Then
                strFileLocation = .Text
            Else
                MsgBox "File doesn't exist.  Correct the file name or" & vbNewLine _
                     & "clear it to Add by ProgID or Cancel.", _
                       vbOKOnly Or vbExclamation, _
                       "MMM - Dependency file does not exist"
                .SelStart = 0
                .SelLength = Len(.Text)
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub txtProgID_Change()
    txtFileLocation_Change
End Sub

Private Sub txtProgID_Validate(Cancel As Boolean)
    Dim strCLSID As String
    
    With txtProgID
        .Text = Trim$(.Text)
        If Len(.Text) > 0 Then
            If Len(txtFileLocation.Text) > 0 Then
                MsgBox "Can't Add by both file name and ProgID." & vbNewLine _
                     & "Clear one other the other, or Cancel.", _
                       vbOKOnly Or vbExclamation, _
                       "MMM - Can't Add by both file name and ProgID"
                .SelStart = 0
                .SelLength = Len(.Text)
                Cancel = True
            Else
                strCLSID = _
                    GetRegistryValue(HKEY_CLASSES_ROOT, .Text & "\CLSID", "")
                If StrPtr(strCLSID) = 0 Then
                    MsgBox "Can't locate CLSID for this ProgID.  Correct the" & vbNewLine _
                         & "ProgID, clear and Add by file name, or Cancel.", _
                           vbOKOnly Or vbExclamation, _
                           "MMM - Can't find CLSID for ProgID"
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    Cancel = True
                Else
                    strFileLocation = _
                        GetRegistryValue(HKEY_CLASSES_ROOT, _
                                         "CLSID\" & strCLSID & "\InprocServer32", _
                                         "")
                    If StrPtr(strFileLocation) = 0 Then
                        MsgBox "Can't locate library file for ProgID.  Correct the" & vbNewLine _
                             & "ProgID, clear and Add by file name, or Cancel." & vbNewLine _
                             & vbNewLine _
                             & "Note: MMM can't isolate out of process servers.", _
                               vbOKOnly Or vbExclamation, _
                               "MMM - Can't find library for ProgID"
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        Cancel = True
                    Else
                        strFileLocation = ExpandEnv(strFileLocation)
                        
                        If Not FilePresent(strFileLocation) Then
                            MsgBox "File doesn't exist.  Correct the ProgID or" & vbNewLine _
                                 & "clear it to Add by file name or Cancel.", _
                                   vbOKOnly Or vbExclamation, _
                                   "MMM - Dependency file does not exist"
                            .SelStart = 0
                            .SelLength = Len(.Text)
                            Cancel = True
                        End If
                        'Else good.
                    End If
                End If
            End If
        End If
    End With
End Sub
