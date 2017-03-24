VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   4200
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   4590
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmSplash"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4200
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblCompany 
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3900
      Width           =   4335
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   780
      TabIndex        =   2
      Top             =   1920
      Width           =   3675
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Left            =   120
      Picture         =   "frmSplash.frx":0943
      Top             =   1500
      Width           =   720
   End
   Begin VB.Label lblProductName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   780
      TabIndex        =   1
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   780
      TabIndex        =   0
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   4095
      Left            =   60
      Top             =   60
      Width           =   4470
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----- Event Handlers -----

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    FontWiz.AdjustControls Me.Controls

    lblProductName.Caption = App.ProductName
    lblDescription.Caption = App.FileDescription
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblCompany.Caption = App.CompanyName
End Sub

Private Sub imgIcon_Click()
    Unload Me
End Sub

Private Sub lblDescription_Click()
    Unload Me
End Sub

Private Sub lblProductName_Click()
    Unload Me
End Sub

Private Sub lblVersion_Click()
    Unload Me
End Sub
