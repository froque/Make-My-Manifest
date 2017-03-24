VERSION 5.00
Begin VB.Form frmLog 
   Caption         =   "Log"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   ControlBox      =   0   'False
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "frmLog"
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   6975
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtLog 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5835
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----- Public Methods -----

Public Sub Log(Optional ByVal Text As String = "")
    With txtLog
        TBSetSel txtLog, &H7FFFFFFF
        TBWriteSel txtLog, Text & vbNewLine
        TBSetSel txtLog, &H7FFFFFFF
        .Refresh
    End With
    '''DoEvents
End Sub

Public Sub SaveLog(ByVal Path As String)
    Dim intFile As Integer
    Dim strLog As String
        
    intFile = FreeFile()
    Open Path & "MMMLog.txt" For Output As #intFile
    strLog = TBRead(txtLog, 262144)
    Print #intFile, strLog
    Close #intFile
End Sub

Private Sub Form_Load()
    FontWiz.AdjustControls Me.Controls
End Sub

'----- Event Handlers -----

Private Sub Form_Resize()
    txtLog.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
