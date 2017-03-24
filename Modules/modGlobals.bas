Attribute VB_Name = "modGlobals"
Option Explicit

'----- Public Constants -----

Public Const GC_PRODUCT_SETTINGS_FILE As String = "MMM.ini"

'tbrMain constants.
Public Const GC_TBBTN_OPEN As Integer = 1
            'Placeholder = 2
Public Const GC_TBBTN_LOG As Integer = 3
Public Const GC_TBBTN_PROJECT As Integer = 4
Public Const GC_TBBTN_SCAN As Integer = 5
Public Const GC_TBBTN_DEPS As Integer = 6
Public Const GC_TBBTN_ADDEDFILES As Integer = 7
Public Const GC_TBBTN_SETTINGS As Integer = 8
Public Const GC_TBBTN_MAKE As Integer = 9
            'Placeholder = 10
Public Const GC_TBBTN_HELP As Integer = 11

'Wizard navigation button offsets.
Public Const GC_BUTTONSTOPDELTA As Single = 435
Public Const GC_BACKLEFTDELTA As Single = 3315
Public Const GC_NEXTLEFTDELTA As Single = 2175
Public Const GC_CANCELLEFTDELTA As Single = 1035
Public Const GC_UNREGWARNLEFTDELTA As Single = 3855

'Project Properties.
'lvwProjProps constants.
'These are used as collection Keys as well as Text values:
Public Const GPP_DEPLIBS As String = "Dependency libraries"
Public Const GPP_PERMEXCLUDED As String = "Hard-excluded libraries"
Public Const GPP_PROJCOMPANY As String = "VB project company name"
Public Const GPP_PROJDESC As String = "VB project description"
Public Const GPP_PROJEXE As String = "VB project EXE name"
Public Const GPP_PROJPATH32 As String = "VB project EXE path"
Public Const GPP_PROJFILEDESC As String = "VB project file description"
Public Const GPP_PROJNAME As String = "VB project name"
Public Const GPP_PROJPATH As String = "Project folder"
Public Const GPP_PROJVERS As String = "VB project version"
Public Const GPP_VBPROJFILE As String = "VB project file"

'----- Public Data -----

Public G_Manifest As ADODB.Stream
Public G_Product_IniDOM As IniDOM
Public G_Package_IniDOM As IniDOM
Public G_ProjLibData As New Collection

Public G_PermExcludedCount As Long
Public G_UnregWarning As Boolean
Public G_ExitMode As Boolean

Public G_Decimal As String 'Char used as decimal point.
Public G_DepsFolder As String
Public G_PackageFQFolder As String
Public G_PackageFolder As String
Public G_Product_Identity As String
Public G_Product_Settings_Path As String
Public G_ProjCompany As String
Public G_ProjDescription As String
Public G_ProjEXE As String
Public G_ProjPath32 As String
Public G_ProjEXE_FQFileName As String
Public G_ProjFileDescription As String
Public G_ProjMajorVer As String
Public G_ProjMinorVer As String
Public G_ProjName As String
Public G_ProjRevisionVer As String
Public G_ProjVersion As String
Public G_VBP_FileName As String
Public G_VBP_FQFileName As String
Public G_VBP_Path As String

