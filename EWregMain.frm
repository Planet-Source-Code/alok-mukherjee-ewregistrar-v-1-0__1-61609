VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form EWregMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  EWregistrar"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   ForeColor       =   &H00000000&
   Icon            =   "EWregMain.frx":0000
   LinkTopic       =   "EWregMain"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "EWregMain.frx":57E2
   ScaleHeight     =   6795
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMemory 
      Interval        =   1
      Left            =   3150
      Top             =   3960
   End
   Begin VB.Frame FrameJob 
      Caption         =   "[ Job Done ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2325
      Left            =   390
      TabIndex        =   6
      Top             =   4020
      Width           =   8295
      Begin EWregistrar.chameleonButton ExitButton 
         Default         =   -1  'True
         Height          =   645
         Left            =   6735
         TabIndex        =   20
         Top             =   1590
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1138
         BTYPE           =   3
         TX              =   "&Exit     "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "EWregMain.frx":5B24
         PICN            =   "EWregMain.frx":5E3E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   1
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin EWregistrar.chameleonButton HelpButton 
         Height          =   645
         Left            =   6735
         TabIndex        =   19
         Top             =   930
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1138
         BTYPE           =   3
         TX              =   "&Help    "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "EWregMain.frx":C0D8
         PICN            =   "EWregMain.frx":C3F2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   1
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin EWregistrar.chameleonButton ReportButton 
         Height          =   645
         Left            =   2040
         TabIndex        =   17
         Top             =   1590
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1138
         BTYPE           =   3
         TX              =   "Re&port  "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "EWregMain.frx":F474
         PICN            =   "EWregMain.frx":F78E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   1
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin EWregistrar.chameleonButton RegButton 
         Height          =   645
         Left            =   90
         TabIndex        =   16
         Top             =   1590
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1138
         BTYPE           =   3
         TX              =   "&Register  "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "EWregMain.frx":12810
         PICN            =   "EWregMain.frx":12B2A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   1
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ListBox lstJobDone 
         Height          =   1230
         ItemData        =   "EWregMain.frx":18DC4
         Left            =   120
         List            =   "EWregMain.frx":18DC6
         TabIndex        =   15
         Top             =   270
         Width           =   6525
      End
      Begin VB.OptionButton optReg 
         Caption         =   "&Register"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3765
         TabIndex        =   12
         Top             =   1650
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optUnReg 
         Caption         =   "&Un-Register"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4995
         TabIndex        =   11
         Top             =   1650
         Width           =   1425
      End
      Begin VB.CheckBox chkCopy 
         Caption         =   "Copy Files to System Directory"
         Height          =   225
         Left            =   3765
         TabIndex        =   10
         Top             =   1980
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin EWregistrar.chameleonButton AboutButton 
         Height          =   645
         Left            =   6735
         TabIndex        =   18
         Top             =   270
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1138
         BTYPE           =   3
         TX              =   "&About  "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "EWregMain.frx":18DC8
         PICN            =   "EWregMain.frx":190E2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   1
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   345
      Left            =   330
      TabIndex        =   4
      Top             =   6435
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "21/07/2005"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "2:00 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11218
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   9375
      Left            =   0
      ScaleHeight     =   9375
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   0
      Width           =   315
   End
   Begin VB.Frame frameMain 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3015
      Left            =   390
      TabIndex        =   1
      Top             =   870
      Width           =   8295
      Begin VB.CheckBox chkSubDir 
         Caption         =   "Find in Sub Directories also"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5460
         TabIndex        =   8
         Top             =   2640
         Value           =   1  'Checked
         Width           =   2685
      End
      Begin MSComctlLib.ListView lstFiles 
         Height          =   2265
         Left            =   3180
         TabIndex        =   5
         Top             =   180
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   3995
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList2"
         ForeColor       =   16711680
         BackColor       =   12648447
         BorderStyle     =   1
         Appearance      =   0
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
            Text            =   "ActiveX EXE/DLL/OCX/TLB/OLB"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Information"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   3045
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2640
         Width           =   7485
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   3045
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2280
         Width           =   4695
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2280
         Width           =   2025
      End
      Begin VB.ListBox lstTypeLibs 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   2010
         ItemData        =   "EWregMain.frx":199BC
         Left            =   120
         List            =   "EWregMain.frx":199BE
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   180
         Width           =   8055
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Path   :"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   2670
         Width           =   510
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Version :"
         Height          =   195
         Index           =   2
         Left            =   5490
         TabIndex        =   25
         Top             =   2310
         Width           =   615
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "GUID :"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   2310
         Width           =   495
      End
      Begin VB.Label lblFilesFound 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   45
      End
   End
   Begin VB.Label lblTPM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total physical memory : "
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   510
      TabIndex        =   14
      Top             =   330
      Width           =   1695
   End
   Begin VB.Label lblAPM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Availbale physical memory : "
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   510
      TabIndex        =   13
      Top             =   540
      Width           =   1980
   End
   Begin VB.Label lblWinVer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OS : Windows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   510
      TabIndex        =   9
      Top             =   60
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   4950
      Picture         =   "EWregMain.frx":199C0
      Top             =   60
      Width           =   3660
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   600
      Left            =   5010
      Top             =   120
      Width           =   3660
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   795
      Left            =   330
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "EWregMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim cTip As New cTooltip
  Dim cLogo As New cLogo

  Dim DuplicateForm As Boolean
  Dim blnHasRun As Boolean
  Dim itmX As ListItem
  Dim i, ButtonEnable, strLastDir As String
  Dim strWinDirPath As String * 255
  Dim strSysDirPath As String * 255
  Dim lonStatusW As Long
  Dim lonStatusS As Long
  Dim FileCount As Integer
  Dim DirCount As Integer
  Dim mEnum As STATUS
  
  'Get the memory status
  Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
  Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
  Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
  Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
  Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
  Private Declare Function GetWindowsDirectory& Lib "kernel32" _
        Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
        ByVal nSize As Long)
  Private Declare Function GetSystemDirectory& Lib "kernel32" _
        Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
        ByVal nSize As Long)
  Private Declare Function SendMessageList Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
  Private Declare Function GetDialogBaseUnits Lib "user32" () As Long
  Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, _
    ByVal lpszWindow As String) As Long
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
  Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
  Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

  'Common Dialog Error
  Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
  Private Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathA" (ByVal pszPath As String)
  Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
  Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

  Private Const ERROR_SUCCESS = &H0
  Private Const FS_CASE_IS_PRESERVED = &H2
  Private Const FS_CASE_SENSITIVE = &H1
  Private Const FS_UNICODE_STORED_ON_DISK = &H4
  Private Const FS_PERSISTENT_ACLS = &H8
  Private Const FS_FILE_COMPRESSION = &H10
  Private Const FS_VOL_IS_COMPRESSED = &H8000
  Private Const MAX_PATH = 260
  Private Const MAXDWORD = &HFFFF
  Private Const INVALID_HANDLE_VALUE = -1
  Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
  Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
  Private Const FILE_ATTRIBUTE_HIDDEN = &H2
  Private Const FILE_ATTRIBUTE_NORMAL = &H80
  Private Const FILE_ATTRIBUTE_READONLY = &H1
  Private Const FILE_ATTRIBUTE_SYSTEM = &H4
  Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
  Private Const LB_ADDSTRING = &H180
  Private Const SW_SHOWNORMAL = 1
  Private Const MAX_FILENAME_LEN = 260
  Private Const EC_LEFTMARGIN = &H1
  Private Const EC_RIGHTMARGIN = &H2
  Private Const EC_USEFONTINFO = &HFFFF&
  Private Const EM_SETMARGINS = &HD3&
  Private Const EM_GETMARGINS = &HD4&
  Private Const LB_GETHORIZONTALEXTENT = &H193
  Private Const LB_SETHORIZONTALEXTENT = &H194
  Private Const LB_SETTABSTOPS = &H192
  
  Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
  End Type

  Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
  End Type
  
  Private Type MEMORYSTATUS
    dwLength                As Long
    dwMemoryLoad            As Long
    dwTotalPhys             As Long
    dwAvailPhys             As Long
    dwTotalPageFile         As Long
    dwAvailPageFile         As Long
    dwTotalVirtual          As Long
    dwAvailVirtual          As Long
  End Type

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As cTypeLibInfo
  Dim objT As cTypeLibInfo
  ' Bruce McKinney's code for getting an Object from the
  ' object pointer:
  CopyMemory objT, lPtr, 4
  Set ObjectFromPtr = objT
  CopyMemory objT, 0&, 4
End Property

Public Property Get FileExists(ByVal sFile As String) As Boolean
  On Error Resume Next
  sFile = Dir(sFile)
  FileExists = ((Err.Number = 0) And sFile <> "")
  On Error GoTo 0
End Property

Public Sub Populate()
Dim iSectCount As Long, iSect As Long, sSections() As String
Dim iVerCount As Long, iVer As Long, sVersions() As String
Dim iExeSectCount As Long, sExeSect() As String
Dim iExeSect As Long
Dim bFoundExeSect As Boolean
Dim sExists As String
Dim cTLI As cTypeLibInfo
Dim i As IShellFolderEx_TLB.IUnknown

   pClearList
   lstTypeLibs.Clear
   lstTypeLibs.Visible = False

   Dim cR As New cRegistry
   cR.ClassKey = HKEY_CLASSES_ROOT
   cR.ValueType = REG_SZ
   cR.SectionKey = "TypeLib"
   ' Get the registered Type Libs:
   If cR.EnumerateSections(sSections(), iSectCount) Then
      For iSect = 1 To iSectCount
         ' Enumerate the versions for each typelib:
         cR.SectionKey = "TypeLib\" & sSections(iSect)
         'MsgBox (cR.SectionKey)
         If cR.EnumerateSections(sVersions(), iVerCount) Then
            For iVer = 1 To iVerCount
               Set cTLI = New cTypeLibInfo
               cTLI.CLSID = sSections(iSect)
               cTLI.Ver = sVersions(iVer)
               cR.SectionKey = "TypeLib\" & sSections(iSect) & "\" & sVersions(iVer)
               cTLI.Name = cR.Value
               cR.EnumerateSections sExeSect(), iExeSectCount
               If iExeSectCount > 0 Then
                  bFoundExeSect = False
                  For iExeSect = 1 To iExeSectCount
                     If IsNumeric(sExeSect(iExeSect)) Then
                        cR.SectionKey = cR.SectionKey & "\" & sExeSect(iExeSect) & "\win32"
                        bFoundExeSect = True
                        Exit For
                     End If
                  Next iExeSect
                  If bFoundExeSect Then
                     cTLI.Path = cR.Value
                     If FileExists(cTLI.Path) Then
                        sExists = "Y"
                     Else
                        sExists = "N"
                     End If
                  Else
                     sExists = "N"
                  End If
               Else
                  sExists = "N"
               End If
               cTLI.Exists = (StrComp(sExists, "Y") = 0)
               If Len(cTLI.Name) > 0 Then
                  lstTypeLibs.AddItem cTLI.Name & vbTab & sExists
               Else
                  lstTypeLibs.AddItem cTLI.CLSID & vbTab & sExists
               End If
               lstTypeLibs.ItemData(lstTypeLibs.NewIndex) = ObjPtr(cTLI)
               Set i = cTLI
               i.AddRef
            Next iVer
         End If
      Next iSect
   End If
   
   lstTypeLibs.Visible = True
   
End Sub

Function StripNulls(OriginalStr As String) As String
  If (InStr(OriginalStr, Chr(0)) > 0) Then
    OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
  End If
  StripNulls = OriginalStr
End Function

Public Sub TabStop(ByVal hWndA As Long, lTabPositions() As Long)
  Dim lCount As Long
  Dim lBaseUnitX As Long
  Dim lBaseUnit As Long
  Dim lTabDlgUnitPos() As Long
  Dim i As Long

  On Error Resume Next
  lCount = UBound(lTabPositions) - LBound(lTabPositions) + 1
  If lCount > 0 Then
    lBaseUnit = GetDialogBaseUnits()
    lBaseUnitX = lBaseUnit And &HFFFF&
    ReDim lTabDlgUnitPos(0 To lCount - 1) As Long
    For i = 0 To lCount - 1
      lTabDlgUnitPos(i) = (lTabPositions(i + LBound(lTabPositions)) * 4) / lBaseUnitX
    Next i
    i = SendMessage(hWndA, LB_SETTABSTOPS, lCount, lTabDlgUnitPos(0))
  End If
End Sub

Private Sub pDeleteEntry(ByVal lPtr As Long)
  Dim cTLI As cTypeLibInfo
  Set cTLI = ObjectFromPtr(lPtr)
  Dim cR As New cRegistry
  cR.ClassKey = HKEY_CLASSES_ROOT
  cR.SectionKey = "TypeLib\" & cTLI.CLSID & "\" & cTLI.Ver
  On Error Resume Next
  If cR.DeleteKey Then
    If Err.Number = 0 Then
      MsgBox "Successfully deleted the item " & cTLI.Name & ", version " & cTLI.Ver, vbInformation
    End If
  End If
  Err.Clear
  On Error GoTo 0
End Sub

Function FindFilesAPI(Path As String, SearchStr As String, FileCount As Integer, DirCount As Integer)
  Dim FileName As String ' Walking filename variable...
  Dim DirName As String ' SubDirectory Name
  Dim dirNames() As String ' Buffer for directory name entries
  Dim nDir As Integer ' Number of directories in this path
  Dim i As Integer ' For-loop counter...
  Dim hSearch As Long ' Search Handle
  Dim WFD As WIN32_FIND_DATA
  Dim Cont As Integer
  Dim cTLI As TypeLibInfo

  If Right(Path, 1) <> "\" Then Path = Path & "\"
  ' Search for subdirectories.
  nDir = 0
  ReDim dirNames(nDir)
  Cont = True
  hSearch = FindFirstFile(Path & "*", WFD)
  If hSearch <> INVALID_HANDLE_VALUE Then
    Do While Cont
      DirName = StripNulls(WFD.cFileName)
      ' Ignore the current and encompassing directories.
      If (DirName <> ".") And (DirName <> "..") Then
        ' Check for directory with bitwise comparison.
        If GetFileAttributes(Path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
          dirNames(nDir) = DirName
          DirCount = DirCount + 1
          nDir = nDir + 1
          ReDim Preserve dirNames(nDir)
        End If
      End If
      Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
    Loop
    Cont = FindClose(hSearch)
  End If
  ' Walk through this directory and sum file sizes.
  hSearch = FindFirstFile(Path & SearchStr, WFD)
  Cont = True
  If hSearch <> INVALID_HANDLE_VALUE Then
    While Cont
      FileName = StripNulls(WFD.cFileName)
      If (FileName <> ".") And (FileName <> "..") Then
        FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
        FileCount = FileCount + 1
        Set itmX = lstFiles.ListItems.Add(, , Path & FileName)
        Select Case UCase(Right(Path & FileName, 3))
          Case "TLB", "OLB", "DLL", "OCX"
            On Error Resume Next
            Set cTLI = TLI.TypeLibInfoFromFile(Path & FileName)
            If Not IsNull(cTLI.Name) Then
              itmX.SubItems(1) = CStr(cTLI.Name & " (" & cTLI.HelpString & ")")
            Else
              itmX.SubItems(1) = "-- No Information Available --"
            End If
          Case Else
            '
        End Select
        Set cTLI = Nothing
      End If
      Cont = FindNextFile(hSearch, WFD) ' Get next file
    Wend
    Cont = FindClose(hSearch)
  End If
  ' ===================================
  ' Do not Serch in Sub directories
  ' ===================================
  If chkSubDir.Value = 1 Then
    ' If there are sub-directories...
    If nDir > 0 Then
      ' Recursively walk into them...
      For i = 0 To nDir - 1
        FindFilesAPI = FindFilesAPI + FindFilesAPI(Path & dirNames(i) & "\", SearchStr, FileCount, DirCount)
      Next i
    End If
  End If
End Function

Private Sub AboutButton_Click()
  Me.WindowState = vbMinimized
  Dim frmEWabout As New EWabout
  Load frmEWabout
  frmEWabout.Show
  Set frmEWabout = Nothing
End Sub

Private Sub AboutButton_MouseOver()
  With cTip
    .Style = TTBalloon
    .Icon = TTIconInfo
    .Centered = True
    .BackColor = &HC0FFFF
    .ForeColor = &H0
    Set .ParentControl = AboutButton
    .TipText = "About the Developers"
    .Title = "About Us"
    .Create
  End With
End Sub

Private Sub chkSubDir_Click()
  Call FillListView
End Sub

Private Sub Dir1_Change()
  Call FillListView
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub ExitButton_Click()
  Unload Me
End Sub

Private Sub ExitButton_MouseOver()
  With cTip
    .Style = TTBalloon
    .Icon = TTIconError
    .Centered = True
    .BackColor = &HC0FFFF
    .ForeColor = &H0
    Set .ParentControl = ExitButton
    .TipText = "Close Everything and Exit the Program"
    .Title = "Close Me"
    .Create
  End With
End Sub

Private Sub Form_Activate()
  Screen.MousePointer = vbHourglass
  If blnHasRun Then
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  Call HideUnhide(True)
  Call API_DoEvents
  blnHasRun = True
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  cLogo.DrawingObject = picLogo
  cLogo.Caption = "ActiveX EXE/DLL/OCX/TLB/OLB Registration Utility"
  Me.Caption = Me.Caption & " Build " & App.Major & _
               "." & App.Minor & "." & App.Revision
  Me.BackColor = frameMain.BackColor
  lblWinVer.Caption = "OS : " & GetVersionString
  lonStatusW = GetWindowsDirectory&(strWinDirPath, 255)
  lonStatusS = GetSystemDirectory&(strSysDirPath, 255)
  strLastDir = Trim(Dir1.Path)
  Call Populate
  Call FillListView
  Call ChkLstChecked
  ReportButton.Enabled = False
  StatusBar1.Panels(3).Text = "Please select File(s) to Register"
  ReDim lTabPos(0 To 0) As Long
  lTabPos(0) = (lstTypeLibs.Width \ Screen.TwipsPerPixelX) - 32
  TabStop lstTypeLibs.hwnd, lTabPos()
  If lstTypeLibs.ListCount > 0 Then
    lstTypeLibs.ListIndex = 0
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call pClearList
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  picLogo.Height = Me.ScaleHeight
  On Error GoTo 0
  cLogo.Draw
End Sub

Private Function FillListView()
  Dim SearchPath As String, FindStr As String
  Dim FileSize As Long
  Dim NumFiles As Integer, NumDirs As Integer
  Screen.MousePointer = vbHourglass
  lblFilesFound.Caption = ""
  StatusBar1.Panels(3).Text = "Searching DLL files..!"
  lstFiles.ListItems.Clear
  SearchPath = Dir1.Path
  FindStr = "*.DLL"
  FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
  lblFilesFound.Caption = NumFiles & " Files found in " & NumDirs + 1 & " Directories"
  
  StatusBar1.Panels(3).Text = "Searching OCX files..!"
  FindStr = "*.OCX"
  FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
  lblFilesFound.Caption = NumFiles & " Files found in " & NumDirs + 1 & " Directories"

  StatusBar1.Panels(3).Text = "Searching TLB files..!"
  FindStr = "*.TLB"
  FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
  lblFilesFound.Caption = NumFiles & " Files found in " & NumDirs + 1 & " Directories"

  StatusBar1.Panels(3).Text = "Searching OLB files..!"
  FindStr = "*.OLB"
  FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
  lblFilesFound.Caption = NumFiles & " Files found in " & NumDirs + 1 & " Directories"
  
  StatusBar1.Panels(3).Text = "Searching ActiveX EXE files..!"
  FindStr = "*.EXE"
  FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
  lblFilesFound.Caption = NumFiles & " Files found in " & NumDirs + 1 & " Directories"
  
  ' Now Auto Size the Listview
  ' ==========================
  StatusBar1.Panels(3).Text = "Trimming List Details..!"
  Dim i As Integer
  For i = 0 To lstFiles.ColumnHeaders.Count - 1
    Call AutoSizeListView(lstFiles, i)
  Next
  StatusBar1.Panels(3).Text = "Searching Completed..!"
  Screen.MousePointer = vbDefault
End Function

Private Sub Form_Unload(Cancel As Integer)
  Set cTip = Nothing
  Set cLogo = Nothing
End Sub

Private Sub HelpButton_Click()
  Dim frmTip As New EWtip
  Load frmTip
  frmTip.Show
  Set frmTip = Nothing
End Sub

Private Sub HelpButton_MouseOver()
  With cTip
    .Style = TTBalloon
    .Icon = TTIconInfo
    .Centered = True
    .BackColor = &HC0FFFF
    .ForeColor = &H0
    Set .ParentControl = HelpButton
    .TipText = "Help on Operations"
    .Title = "Help"
    .Create
  End With
End Sub

Private Sub lstFiles_GotFocus()
  Call ChkLstChecked
End Sub

Private Sub lstFiles_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  Call ChkLstChecked
End Sub

Function ChkLstChecked()
  Dim FileCount As Integer
  Dim itemChecked As Integer
  itemChecked = 0
  For FileCount = 1 To lstFiles.ListItems.Count
    If lstFiles.ListItems(FileCount).Checked = True Then _
      itemChecked = itemChecked + 1
  Next
  If itemChecked = 0 Then RegButton.Enabled = False _
     Else RegButton.Enabled = True
End Function

Private Sub lstFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Call ChkLstChecked
End Sub

Private Sub lstFiles_LostFocus()
  Call ChkLstChecked
End Sub

Private Sub lstTypeLibs_Click()
  Dim lI As Long
  Dim lPtr As Long
  Dim cTLI As cTypeLibInfo
  lI = lstTypeLibs.ListIndex
  If lI > -1 Then
    lPtr = lstTypeLibs.ItemData(lI)
    If Not (lPtr = 0) Then
      Set cTLI = ObjectFromPtr(lPtr)
      'txtInfo(0).Text = cTLI.Name
      txtInfo(1).Text = cTLI.CLSID
      txtInfo(2).Text = cTLI.Ver
      txtInfo(3).Text = cTLI.Path
      If Not cTLI.Exists Then
        txtInfo(3).ForeColor = &HC0&  ' Red
      Else
        txtInfo(3).ForeColor = vbWindowText
      End If
    End If
  End If
End Sub

Private Sub optReg_Click()
  Call HideUnhide(True)
  RegButton.Caption = "&Register  "
  chkCopy.Caption = "Copy Files to System Directory"
  Drive1.Drive = Left(strLastDir, 1)
  Dir1.Path = strLastDir
  Call ChkLstChecked
End Sub

Private Sub optUnReg_Click()
  Call HideUnhide(False)
  RegButton.Caption = "&Un-Register"
  chkCopy.Caption = "Delete Files from System Directory"
  strLastDir = Dir1.Path
  'Drive1.Drive = Left(strSysDirPath, 1)
  'Dir1.Path = strSysDirPath
  RegButton.Enabled = True
End Sub

Private Sub HideUnhide(bHideUnhide As Boolean)
  Dim i As Integer
  lstTypeLibs.Visible = Not bHideUnhide
  For i = 1 To 3
    lblInfo(i).Visible = Not bHideUnhide
    txtInfo(i).Visible = Not bHideUnhide
  Next
  Drive1.Visible = bHideUnhide
  Dir1.Visible = bHideUnhide
  lstFiles.Visible = bHideUnhide
  chkSubDir.Visible = bHideUnhide
  lblFilesFound.Visible = bHideUnhide
  Me.Refresh
  Call API_DoEvents
End Sub

Private Sub ReportButton_Click()
  Dim i As Integer
  For i = 0 To lstJobDone.ListCount - 1
    Call WriteALogFile("[ " & Str(Date) & " " & Str(Time) & " ] - " & lstJobDone.List(i))
  Next
  'Check if the file exists
  If Dir(sLogFile) = "" Or sLogFile = "" Then
    MsgBox "Could not read Log File / not created!", vbCritical
    Exit Sub
  End If
  ShellExecute Me.hwnd, vbNullString, sLogFile, vbNullString, App.Path, SW_SHOWNORMAL
End Sub

Private Sub ReportButton_MouseOver()
  With cTip
    .Style = TTBalloon
    .Icon = TTIconInfo
    .Centered = True
    .BackColor = &HC0FFFF
    .ForeColor = &H0
    Set .ParentControl = ReportButton
    .TipText = "Print Report on Registration / Un-Registration"
    .Title = "Print"
    .Create
  End With
End Sub

Private Sub RegButton_Click()
  Dim FileCount As Integer
  Dim sFile As String, sFileName As String
  Dim lPtr As Long
  Dim lPtrCurrent As Long
  Dim lI As Long
  Dim cTLI As cTypeLibInfo
  Dim bAlso As Boolean
  Dim sPath As String
  Dim lOrigIndex As Long
  
  On Error Resume Next
  If optReg.Value = True Then
    ' Register EXE/TLB/DLL/OLB/OCX
    ' ============================
    For FileCount = 1 To lstFiles.ListItems.Count
      If lstFiles.ListItems(FileCount).Checked = True Then
        sFile = StripTerminator(lstFiles.ListItems(FileCount).Text)
        sFileName = sFile
        PathStripPath sFileName
        If chkCopy.Value = 1 Then
          ' Copy the file to Systems Directory &
          ' Register the file from System Directory
          ' =======================================
          If (CopyFile(sFile, StripTerminator(strSysDirPath) & _
            "\" & StripTerminator(sFileName), 0)) <> 0 Then
            lstJobDone.AddItem (sFile & _
              " Copied to " & StripTerminator(strSysDirPath) & ".")
          Else
            lstJobDone.AddItem (sFile & _
              " could not be Copied to " & StripTerminator(strSysDirPath) & ".")
          End If
          Select Case UCase(Right(StripTerminator(sFileName), 3))
            Case "EXE", "DLL", "OCX"
              mEnum = RegisterComponent(StripTerminator(strSysDirPath) & _
                "\" & StripTerminator(sFileName), DllRegisterServer) 'to Register
              If mEnum = [File Could Not Be Loaded Into Memory Space] Then
                lstJobDone.AddItem StripTerminator(strSysDirPath) & _
                "\" & StripTerminator(sFileName) & " ( Error: File Could Not Be Loaded Into Memory Space )"
              ElseIf mEnum = [Not A Valid ActiveX Component] Then
                lstJobDone.AddItem StripTerminator(strSysDirPath) & _
                "\" & StripTerminator(sFileName) & " ( Error: Not A Valid ActiveX Component )"
              ElseIf mEnum = [ActiveX Component Registration Failed] Then
                lstJobDone.AddItem StripTerminator(strSysDirPath) & _
                "\" & StripTerminator(sFileName) & " ( Error: ActiveX Component Registration Failed )"
              ElseIf mEnum = [ActiveX Component Registered Successfully] Then
                lstJobDone.AddItem StripTerminator(strSysDirPath) & _
                "\" & StripTerminator(sFileName) & " ( ActiveX Component Registered Successfully )"
              End If
            Case "TLB", "OLB"
              UIRegisterTypeLib StripTerminator(strSysDirPath) & _
                "\" & StripTerminator(sFileName), True, True
          End Select
        Else
          ' Do not copy the file
          ' Register from Original File Location
          ' ====================================
          Select Case UCase(Right(StripTerminator(sFile), 3))
            Case "EXE", "DLL", "OCX"
              mEnum = RegisterComponent(StripTerminator(sFile), DllRegisterServer) 'to Register
              If mEnum = [File Could Not Be Loaded Into Memory Space] Then
                lstJobDone.AddItem StripTerminator(sFile) & " ( Error: File Could Not Be Loaded Into Memory Space )"
              ElseIf mEnum = [Not A Valid ActiveX Component] Then
                lstJobDone.AddItem StripTerminator(sFile) & " ( Error: Not A Valid ActiveX Component )"
              ElseIf mEnum = [ActiveX Component Registration Failed] Then
                lstJobDone.AddItem StripTerminator(sFile) & " ( Error: ActiveX Component Registration Failed )"
              ElseIf mEnum = [ActiveX Component Registered Successfully] Then
                lstJobDone.AddItem StripTerminator(sFile) & " ( ActiveX Component Registered Successfully )"
              End If
            Case "TLB", "OLB"
              UIRegisterTypeLib StripTerminator(sFile), True, True
          End Select
        End If  ' If Copy to System Directory
      End If  ' If Checked
    Next
    Call FillListView
    API_DoEvents
  Else
    ' Un-Register EXE/TLB/DLL/OLB/OCX
    ' ===============================
    sFile = StripTerminator(txtInfo(3).Text)
    If lstTypeLibs.ListIndex > 0 Then
      lOrigIndex = lstTypeLibs.ListIndex
      lstTypeLibs_Click
      lPtrCurrent = lstTypeLibs.ItemData(lstTypeLibs.ListIndex)
      If Not lPtrCurrent = 0 Then
        If txtInfo(3).ForeColor = &HC0& Then  'Red
          ' typelib not on system, just delete the entry:
          pDeleteEntry lPtrCurrent
        Else
          ' check the same file is not being used elsewhere.
          ' if it is, delete the offending entry & then re-register the
          ' typelib, otherwise just unregister the existing type lib:
          sPath = UCase$(txtInfo(3).Text)
          For lI = 0 To lstTypeLibs.ListCount - 1
            lPtr = lstTypeLibs.ItemData(lI)
            If Not (lPtr = lPtrCurrent) Then
              If Not lPtr = 0 Then
                Set cTLI = ObjectFromPtr(lPtr)
                If StrComp(UCase$(cTLI.Path), sPath) = 0 Then
                  bAlso = True
                  Exit For
                End If
              End If
            End If
          Next lI
          If bAlso Then
            ' Multi use:
            pDeleteEntry lPtrCurrent
            UIRegisterTypeLib txtInfo(3).Text, True, True
          Else
            ' The only use, unregister it:
            UIRegisterTypeLib txtInfo(3).Text, False, True
          End If
        End If
        Call Populate
        If lOrigIndex < lstTypeLibs.ListCount Then
          lstTypeLibs.ListIndex = lOrigIndex
        Else
          lstTypeLibs.ListIndex = lstTypeLibs.ListCount
        End If
      End If
    End If
    
    If chkCopy.Value = 1 Then
      If Not txtInfo(3).ForeColor = &HC0& Then  'Red
        ' if typelib is on system, Delete the file From Systems Directory
        If (DeleteFile(sFile)) <> 0 Then
          lstJobDone.AddItem (sFile & " Deleted.")
        Else
          lstJobDone.AddItem (sFile & " could not be Deleted.")
        End If
      End If
    End If
  End If
  If lstJobDone.ListCount > 0 Then ReportButton.Enabled = True _
     Else ReportButton.Enabled = False
End Sub

Private Sub RegButton_MouseOver()
  With cTip
    .Style = TTBalloon
    .Icon = TTIconWarning
    .Centered = True
    .BackColor = &HC0FFFF
    .ForeColor = &H0
    Set .ParentControl = RegButton
    If optReg Then
      .TipText = "Register Checked Files"
      .Title = "Register"
    Else
      .TipText = "Un-Register Checked Files"
      .Title = "Un-Register"
    End If
    .Create
  End With
End Sub

Private Sub tmrMemory_Timer()
  Dim MEMStat As MEMORYSTATUS
  GlobalMemoryStatus MEMStat
  lblAPM.Caption = "Availbale physical memory : " & Round((MEMStat.dwAvailPhys / 1014) / 1014, 2) & " mb"
  'lblAVM.Caption = "Availbale virtual memory : " & Round((MEMStat.dwAvailVirtual / 1014) / 1014, 2) & " mb"
  lblTPM.Caption = "Total physical memory : " & Round((MEMStat.dwTotalPhys / 1014) / 1014, 2) & " mb"
  'lblTVM.Caption = "Total virtual memory : " & Round((MEMStat.dwTotalVirtual / 1014) / 1014, 2) & " mb"
End Sub

Private Sub pClearList()
  Dim lI As Long
  Dim lPtr As Long
  Dim i As IShellFolderEx_TLB.IUnknown
  Dim cTLI As cTypeLibInfo

  For lI = 0 To lstTypeLibs.ListCount - 1
    lPtr = lstTypeLibs.ItemData(lI)
    If Not (lPtr = 0) Then
      Set cTLI = ObjectFromPtr(lPtr)
      Set i = cTLI
      i.Release
      Set i = Nothing
      Set cTLI = Nothing
    End If
  Next lI
End Sub
