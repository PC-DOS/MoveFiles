VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shadow Move - PC-DOS Workshop"
   ClientHeight    =   5715
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10035
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   10035
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.FileListBox File1 
      Height          =   450
      Left            =   3885
      TabIndex        =   14
      Top             =   6270
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "“∆Ñ”(&M)"
      Height          =   510
      Left            =   8175
      TabIndex        =   8
      Top             =   5160
      Width           =   1800
   End
   Begin VB.Frame Frame1 
      Caption         =   "”∞◊”Œƒº˛≈‰÷√"
      Height          =   1110
      Index           =   2
      Left            =   90
      TabIndex        =   5
      Top             =   4020
      Width           =   9900
      Begin VB.OptionButton Option2 
         Caption         =   "”√∫¨”–“∆Ñ”ƒøòÀŒª÷√ŸY”çµƒŒƒº˛ÃÊìQ‘¥Œƒº˛(&C)"
         Height          =   420
         Left            =   135
         TabIndex        =   7
         Top             =   630
         Value           =   -1  'True
         Width           =   7605
      End
      Begin VB.OptionButton Option1 
         Caption         =   "”√0◊÷πùµƒø’Œƒº˛ÃÊìQ‘¥Œƒº˛(&R)"
         Height          =   420
         Left            =   135
         TabIndex        =   6
         Top             =   240
         Width           =   7605
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Œƒº˛“∆Ñ”ƒøòÀ"
      Height          =   645
      Index           =   1
      Left            =   75
      TabIndex        =   2
      Top             =   3315
      Width           =   9900
      Begin VB.CommandButton Command1 
         Caption         =   "ûg”[(&B)..."
         Height          =   325
         Left            =   8370
         TabIndex        =   4
         Top             =   210
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Height          =   325
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   210
         Width           =   8235
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ÅÌ‘¥Œƒº˛≈‰÷√"
      Height          =   2430
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   765
      Width           =   9900
      Begin VB.CommandButton Command7 
         Caption         =   "ôz“ïﬂx÷–Ìó(&V)"
         Height          =   375
         Left            =   105
         TabIndex        =   16
         Top             =   1950
         Width           =   3060
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Ñh≥˝ﬂx÷–Ìó(&D)"
         Height          =   375
         Left            =   105
         TabIndex        =   15
         Top             =   1530
         Width           =   3060
      End
      Begin VB.CommandButton Command5 
         Caption         =   "«Âø’Œƒº˛¡–±Ì(&L)"
         Height          =   375
         Left            =   105
         TabIndex        =   13
         Top             =   1094
         Width           =   3060
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ÃÌº”Œƒº˛äA÷–µƒŒƒº˛(&F)..."
         Height          =   375
         Left            =   105
         TabIndex        =   12
         Top             =   667
         Width           =   3060
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ÃÌº”…¢¡–Œƒº˛(&A)..."
         Height          =   375
         Left            =   105
         TabIndex        =   11
         Top             =   240
         Width           =   3060
      End
      Begin VB.ListBox List1 
         Height          =   1860
         Left            =   3225
         TabIndex        =   9
         Top             =   465
         Width           =   6570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¥˝“∆Ñ”Œƒº˛¡–±Ì"
         Height          =   180
         Left            =   3225
         TabIndex        =   10
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0ECA
      Height          =   660
      Left            =   900
      TabIndex        =   0
      Top             =   120
      Width           =   9165
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   75
      Picture         =   "Form1.frx":0FD1
      Top             =   -15
      Width           =   720
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Œƒº˛(&F)"
      Begin VB.Menu mnuAddFile 
         Caption         =   "ÃÌº”…¢¡–Œƒº˛(&A)..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFolder 
         Caption         =   "ÃÌº”Œƒº˛äA÷–µƒŒƒº˛(&F)..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuExplorer 
         Caption         =   "‘⁄ŸY‘¥π‹¿Ì∆˜÷–¥ÚÈ_Œƒº˛“∆Ñ”ƒøòÀƒø‰õ(&E)"
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "«Âø’Œƒº˛¡–±Ì(&C)"
      End
      Begin VB.Menu mnuRemoveSelected 
         Caption         =   "“∆≥˝ﬂx÷–Ìó(&R)"
      End
      Begin VB.Menu mnuViewSelected 
         Caption         =   "ôz“ïﬂx÷–Ìó(&V)"
      End
      Begin VB.Menu b2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Ô@ æ¥˝“∆Ñ”Œƒº˛‘îºö¬∑èΩ(&S)..."
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "ﬂxÌó(&O)"
      Begin VB.Menu mnuZeroBytes 
         Caption         =   " π”√0◊÷πùµƒø’Œƒº˛ÅÌÑìΩ®”∞◊”Œƒº˛(&R)"
      End
      Begin VB.Menu mnuWithInfo 
         Caption         =   " π”√∫¨”–“∆Ñ”ƒøòÀŸY”çµƒŒƒº˛ÅÌÑìΩ®”∞◊”Œƒº˛(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu b3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "‘O÷√(&S)..."
      End
      Begin VB.Menu b5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "èÕŒªƒ¨’J‘O÷√(&E)"
      End
      Begin VB.Menu mnuInit 
         Caption         =   "≥ı ºªØ(&I)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "éÕ÷˙(&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "ÍPÏ∂Shadow Move(&A)..."
      End
   End
   Begin VB.Menu mnuHideList 
      Caption         =   "mnuHideList"
      Visible         =   0   'False
      Begin VB.Menu mnuHideRemove 
         Caption         =   "“∆≥˝ﬂx÷–Ìó(&R)"
      End
      Begin VB.Menu mnuHideView 
         Caption         =   "ôz“ïﬂx÷–Ìó(&V)"
      End
      Begin VB.Menu b4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHideClear 
         Caption         =   "«Âø’Œƒº˛¡–±Ì(&C)"
      End
      Begin VB.Menu b6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHideAdd 
         Caption         =   "ÃÌº”…¢¡–Œƒº˛(&A)..."
      End
      Begin VB.Menu mnuHideFolder 
         Caption         =   "ÃÌº”Œƒº˛äA÷–µƒŒƒº˛(&F)..."
      End
   End
   Begin VB.Menu mnuHideR 
      Caption         =   "mnuHideR"
      Visible         =   0   'False
      Begin VB.Menu mnuHRRemove 
         Caption         =   "“∆≥˝ﬂx÷–Ìó(&R)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHRView 
         Caption         =   "ôz“ïﬂx÷–Ìó(&V)"
         Enabled         =   0   'False
      End
      Begin VB.Menu b7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHRClear 
         Caption         =   "«Âø’Œƒº˛¡–±Ì(&C)"
      End
      Begin VB.Menu b9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHRAdd 
         Caption         =   "ÃÌº”…¢¡–Œƒº˛(&A)..."
      End
      Begin VB.Menu mnuHRFolder 
         Caption         =   "ÃÌº”Œƒº˛äA÷–µƒŒƒº˛(&F)..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 ' demo project showing how to use the SHFileOperation API function
  ' by Bryan Stafford of New Vision SoftwareÆ - newvision@imt.net
  ' this demo is released into the public domain "as is" without
  ' warranty or guaranty of any kind.  In other words, use at
  ' your own risk.

  ' ###########################################################
  ' see the comments section at the end of this module for
  ' more info about this API function.
  ' ###########################################################

  ' constants used in the wFunc member of the UDT
  ' to determine which file operation to use.  see
  ' the comments section for further explaination
  Private Const FO_MOVE As Long = &H1&
  Private Const FO_COPY As Long = &H2&
  Private Const FO_DELETE As Long = &H3&
  Private Const FO_RENAME As Long = &H4&

  ' flag constants used in the fFlags member of the UDT
  ' to set the behavior of the dialog when it is displayed.
  ' see the comments section for further explaination
  Private Const FOF_CONFIRMMOUSE As Integer = &H2
  Private Const FOF_SILENT As Integer = &H4
  Private Const FOF_RENAMEONCOLLISION As Integer = &H8
  Private Const FOF_NOCONFIRMATION As Integer = &H10
  Private Const FOF_WANTMAPPINGHANDLE As Integer = &H20
  Private Const FOF_CREATEPROGRESSDLG As Integer = &H0
  Private Const FOF_ALLOWUNDO As Integer = &H40
  Private Const FOF_FILESONLY As Integer = &H80
  Private Const FOF_SIMPLEPROGRESS As Integer = &H100
  Private Const FOF_NOCONFIRMMKDIR As Integer = &H200
  Private Const FOF_NOERRORUI As Integer = &H400
  Private Const FOF_NOCOPYSECURITYATTRIBS As Integer = &H800

  ' UDT to hold the file information
  Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
  End Type ' SHFILEOPSTRUCT

  ' variable to hold the last dir copied to for comparison in the recycle test
  Private sLastCopyDir As String
  
  ' declaration for the API functions
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, _
                                                  hpvSource As Any, ByVal cbCopy As Long)
                        
  Private Declare Function SHFileOperation& Lib "shell32.dll" Alias "SHFileOperationA" _
                                                      (lpFileOp As Any)
Private Enum ShadowFileCreatingOverwritingMethods
EmptyFile = 0
WithDestination = 1
End Enum
Dim lpOWM As ShadowFileCreatingOverwritingMethods
Private Declare Function EmptyClipboard Lib "user32" () As Long
Dim bShowWin As Boolean
Dim bEmpty As Boolean
Private Declare Function SetClipboardViewer Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ChangeClipboardChain Lib "user32" (ByVal hWnd As Long, ByVal hWndNext As Long) As Long
Private Const WM_DRAWCLIPBOARD = &H308
Private Const WM_CHANGECBCHAIN = &H30D
Private Const WM_DESTROY = &H2
Private Const WM_HOTKEY = &H312
Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_FMSYNTH = 4
Private Const MOD_MAPPER = 5
Private Const MOD_MIDIPORT = 1
Private Const MOD_SHIFT = &H4
Private Const MOD_SQSYNTH = 3
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long) As Long
Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Const DEFAULT_SIZE_VALUE = 810
Dim CommonDialog1 As New CCommonDialog
 Private Type PROCESSENTRY32
 dwSize As Long
 cntUsage As Long
 th32ProcessID As Long 'Ω¯≥ÃID
 th32DefaultHeapID As Long '∂—’ªID
 th32ModuleID As Long 'ƒ£øÈID
 cntThreads As Long
 th32ParentProcessID As Long '∏∏Ω¯≥ÃID
 pcPriClassBase As Long
 dwFlags As Long
 szExeFile As String * 260
 End Type
 Private Type MEMORYSTATUS
 dwLength As Long
 dwMemoryLoad As Long
 dwTotalPhys As Long
 dwAvailPhys As Long
 dwTotalPageFile As Long
 dwAvailPageFile As Long
 dwTotalVirtual As Long
 dwAvailVirtual As Long
 End Type
 Private Declare Function NtQuerySystemInformation Lib "ntdll" (ByVal dwInfoType As Long, ByVal lpStructure As Long, ByVal dwSize As Long, ByVal dwReserved As Long) As Long
Private Const SYSTEM_BASICINFORMATION = 0&
Private Const SYSTEM_PERFORMANCEINFORMATION = 2&
Private Const SYSTEM_TIMEINFORMATION = 3&
Private Const NO_ERROR = 0
Private Type LARGE_INTEGER
    dwLow As Long
    dwHigh As Long
End Type

Private Type SYSTEM_PERFORMANCE_INFORMATION
    liIdleTime As LARGE_INTEGER
    dwSpare(0 To 75) As Long
End Type
Private Type SYSTEM_BASIC_INFORMATION
    dwUnknown1 As Long
    uKeMaximumIncrement As Long
    uPageSize As Long
    uMmNumberOfPhysicalPages As Long
    uMmLowestPhysicalPage As Long
    uMmHighestPhysicalPage As Long
    uAllocationGranularity As Long
    pLowestUserAddress As Long
    pMmHighestUserAddress As Long
    uKeActiveProcessors As Long
    bKeNumberProcessors As Byte
    bUnknown2 As Byte
    wUnknown3 As Integer
End Type
Private Type SYSTEM_TIME_INFORMATION
    liKeBootTime As LARGE_INTEGER
    liKeSystemTime As LARGE_INTEGER
    liExpTimeZoneBias As LARGE_INTEGER
    uCurrentTimeZoneId As Long
    dwReserved As Long
End Type

Private lidOldIdle As LARGE_INTEGER
Private liOldSystem As LARGE_INTEGER
 Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
 Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long 'ªÒ»° ◊∏ˆΩ¯≥Ã
 Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long 'ªÒ»°œ¬∏ˆΩ¯≥Ã
 Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long ' Õ∑≈æ‰±˙
 Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
 Private Const TH32CS_SNAPPROCESS = &H2&
 Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Dim IsHideToTray As Boolean
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const VK_LWIN = &H5B
Private Const WM_KEYUP = &H101
Private Const WM_KEYDOWN = &H100
Private Declare Function DebugActiveProcess Lib "kernel32" (ByVal dwProcessId As Long) As Long
Private Declare Sub DebugBreak Lib "kernel32" ()
Private Const SM_DEBUG = 22
Private Const DEBUG_ONLY_THIS_PROCESS = &H2
Private Const DEBUG_PROCESS = &H1
Private Type USER_DIALOG_CONFIG
lpTitle As String
lpIcon As Integer
lpMessage As String
End Type
Private Type USER_APP_RUN
lpAppPath As String
lpAppParam As String
lpRunMode As Integer
End Type
Private Type APP_TASK_PARAM
lpTimerType As Integer
lpDelay As Long
lpRunHour As Integer
lpRunMinute As Integer
lpRunSecond As Integer
lpCurrentHour As Integer
lpCurrentMinute As Integer
lpCurrentSecond As Integer
lpTaskEnum As Integer
lpTaskFriendlyDisplayName As String
lpRunning As Boolean
End Type
Dim lpDialogCfg As USER_DIALOG_CONFIG
Dim lpAppCfg As USER_APP_RUN
Dim lpTaskCfg As APP_TASK_PARAM
Const SC_SCREENSAVE = &HF140&
Dim IsCodeUse As Boolean
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GWL_WNDPROC = (-4)
Private Const WM_SYSCOMMAND = &H112
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Dim lpSize As Long
Dim bchk As Boolean
Dim lpFilePath As String
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1
Const REG_DWORD = 4
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Const MAX_FILE_SIZE = 1.5 * (1024 ^ 3)
Private Const HWND_BOTTOM = 1
Private Const HWND_BROADCAST = &HFFFF&
Private Const HWND_DESKTOP = 0
Private Const HWND_NOTOPMOST = -2
Private Const WS_EX_TRANSPARENT = &H20&
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
'∫‹∂‡≈Û”—∂ºº˚µΩπ˝ƒ‹‘⁄Õ–≈ÃÕº±Í…œ≥ˆœ÷∆¯«ÚÃ· æµƒ»Ìº˛£¨≤ªÀµ»Ìº˛£¨æÕ «‘⁄°∞¥≈≈Ãø’º‰≤ª◊„°± ±Windows∏¯≥ˆµƒÃ· ææÕ Ù”⁄∆¯«ÚÃ· æ£¨ƒ«√¥‘ı—˘‘⁄◊‘º∫µƒ≥Ã–Ú÷–ÃÌº”’‚—˘µƒ∆¯«ÚÃ· æƒÿ£ø
   
'∆‰ µ≤¢≤ªƒ—£¨πÿº¸æÕ‘⁄ÃÌº”Õ–≈ÃÕº±Í ±À˘ π”√µƒNOTIFYICONDATAΩ·ππ£¨‘¥¥˙¬Î»Áœ¬£∫
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
   
Private Type NOTIFYICONDATA
cbSize   As Long     '   Ω·ππ¥Û–°(◊÷Ω⁄)
hWnd   As Long     '   ¥¶¿Ìœ˚œ¢µƒ¥∞ø⁄µƒæ‰±˙
uID   As Long     '   Œ®“ªµƒ±Í ∂∑˚
uFlags   As Long     '   Flags
uCallbackMessage   As Long     '   ¥¶¿Ìœ˚œ¢µƒ¥∞ø⁄Ω” ’µƒœ˚œ¢
hIcon   As Long     '   Õ–≈ÃÕº±Íæ‰±˙
szTip   As String * 128         '   Tooltip   Ã· æŒƒ±æ
dwState   As Long     '   Õ–≈ÃÕº±Í◊¥Ã¨
dwStateMask   As Long     '   ◊¥Ã¨—⁄¬Î
szInfo   As String * 256         '   ∆¯«ÚÃ· æŒƒ±æ
uTimeoutOrVersion   As Long     '   ∆¯«ÚÃ· æœ˚ ß ±º‰ªÚ∞Ê±æ
'   uTimeout   -   ∆¯«ÚÃ· æœ˚ ß ±º‰(µ•Œª:ms,   10000   --   30000)
'   uVersion   -   ∞Ê±æ(0   for   V4,   3   for   V5)
szInfoTitle   As String * 64         '   ∆¯«ÚÃ· æ±ÍÃ‚
dwInfoFlags   As Long     '   ∆¯«ÚÃ· æÕº±Í
End Type
   
'   dwState   to   NOTIFYICONDATA   structure
Private Const NIS_HIDDEN = &H1           '   “˛≤ÿÕº±Í
Private Const NIS_SHAREDICON = &H2           '   π≤œÌÕº±Í
   
'   dwInfoFlags   to   NOTIFIICONDATA   structure
Private Const NIIF_NONE = &H0           '   ŒﬁÕº±Í
Private Const NIIF_INFO = &H1           '   "œ˚œ¢"Õº±Í
Private Const NIIF_WARNING = &H2           '   "æØ∏Ê"Õº±Í
Private Const NIIF_ERROR = &H3           '   "¥ÌŒÛ"Õº±Í
   
'   uFlags   to   NOTIFYICONDATA   structure
Private Const NIF_ICON       As Long = &H2
Private Const NIF_INFO       As Long = &H10
Private Const NIF_MESSAGE       As Long = &H1
Private Const NIF_STATE       As Long = &H8
Private Const NIF_TIP       As Long = &H4
   
'   dwMessage   to   Shell_NotifyIcon
Private Const NIM_ADD       As Long = &H0
Private Const NIM_DELETE       As Long = &H2
Private Const NIM_MODIFY       As Long = &H1
Private Const NIM_SETFOCUS       As Long = &H3
Private Const NIM_SETVERSION       As Long = &H4
Private Type RECTL
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Dim cRect As RECT
Const LCR_UNLOCK = 0
Dim dwMouseFlag As Integer
Const ME_LBCLICK = 1
Const ME_LBDBLCLICK = 2
Const ME_RBCLICK = 3
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Const MOUSETRAILS = 39
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Const SWP_NOACTIVATE = &H10
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Dim lpszCaptionNew As String
Private Const SC_MINIMIZE = &HF020&
Private Const WS_MAXIMIZEBOX = &H10000
Dim HKStateCtrl As Integer
Dim HKStateFn As Integer
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_ICONIC = WS_MINIMIZE
Const SC_ICON = SC_MINIMIZE
Const SC_TASKLIST = &HF130&
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Dim bCodeUse As Boolean
Private Const WS_CAPTION = &HC00000
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Const PROCESS_ALL_ACCESS = &H1F0FFF
Const SC_RESTORE = &HF120&
Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Dim lMeWinStyle As Long
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOOWNERZORDER = &H200
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SC_MOVE = &HF010&
Private Const SC_SIZE = &HF000&
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Const WS_EX_APPWINDOW = &H40000
Private Type WINDOWINFORMATION
hWindow As Long
hWindowDC As Long
hThreadProcess As Long
hThreadProcessID As Long
lpszCaption As String
lpszClassName As String
lpszThreadProcessName As String * 1024
lpszThreadProcessPath As String
lpszExe As String
lpszPath As String
End Type
Private Type WINDOWPARAM
bEnabled As Boolean
bHide As Boolean
bTrans As Boolean
bClosable As Boolean
bSizable As Boolean
bMinisizable As Boolean
bTop As Boolean
lpTransValue As Integer
End Type
Dim lpWindow As WINDOWINFORMATION
Dim lpWindowParam() As WINDOWPARAM
Dim lpCur As Long
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Dim lpRtn As Long
Dim hWindow As Long
Dim lpLength As Long
Dim lpArray() As Byte
Dim lpArray2() As Byte
Dim lpBuff As String
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const LWA_COLORKEY = &H1
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
Private Const WS_SYSMENU = &H80000
Private Const GWL_STYLE = (-16)
Private Const MF_BYCOMMAND = &H0
Private Const SC_CLOSE = &HF060&
Private Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Private Const MF_INSERT = &H0&
Private Const SC_MAXIMIZE = &HF030&
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Type WINDOWINFOBOXDATA
lpszCaption As String
lpszClass As String
lpszThread As String
lpszHandle As String
lpszDC As String
End Type
Dim dwWinInfo As WINDOWINFOBOXDATA
Dim bError As Boolean
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Const WM_CLOSE = &H10
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOMOVE = &H2
Dim mov As Boolean
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Const ANYSIZE_ARRAY = 1
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Private Type LUID
UsedPart As Long
IgnoredForNowHigh32BitPart As Long
End Type
Private Type TOKEN_PRIVILEGES
PrivilegeCount As Long
TheLuid As LUID
Attributes As Long
End Type
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal _
ProcessHandle As Long, _
ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
Alias "LookupPrivilegeValueA" _
(ByVal lpSystemName As String, ByVal lpName As String, lpLuid _
As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
(ByVal TokenHandle As Long, _
ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES _
, ByVal BufferLength As Long, _
PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Type TestCounter
TimesLeft As Integer
ResetTime As Integer
End Type
Dim PassTest As TestCounter
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
X As Long
y As Long
End Type
Private Const VK_ADD = &H6B
Private Const VK_ATTN = &HF6
Private Const VK_BACK = &H8
Private Const VK_CANCEL = &H3
Private Const VK_CAPITAL = &H14
Private Const VK_CLEAR = &HC
Private Const VK_CONTROL = &H11
Private Const VK_CRSEL = &HF7
Private Const VK_DECIMAL = &H6E
Private Const VK_DELETE = &H2E
Private Const VK_DIVIDE = &H6F
Private Const VK_DOWN = &H28
Private Const VK_END = &H23
Private Const VK_EREOF = &HF9
Private Const VK_ESCAPE = &H1B
Private Const VK_EXECUTE = &H2B
Private Const VK_EXSEL = &HF8
Private Const VK_F1 = &H70
Private Const VK_F10 = &H79
Private Const VK_F11 = &H7A
Private Const VK_F12 = &H7B
Private Const VK_F13 = &H7C
Private Const VK_F14 = &H7D
Private Const VK_F15 = &H7E
Private Const VK_F16 = &H7F
Private Const VK_F17 = &H80
Private Const VK_F18 = &H81
Private Const VK_F19 = &H82
Private Const VK_F2 = &H71
Private Const VK_F20 = &H83
Private Const VK_F21 = &H84
Private Const VK_F22 = &H85
Private Const VK_F23 = &H86
Private Const VK_F24 = &H87
Private Const VK_F3 = &H72
Private Const VK_F4 = &H73
Private Const VK_F5 = &H74
Private Const VK_F6 = &H75
Private Const VK_F7 = &H76
Private Const VK_F8 = &H77
Private Const VK_F9 = &H78
Private Const VK_HELP = &H2F
Private Const VK_HOME = &H24
Private Const VK_INSERT = &H2D
Private Const VK_LBUTTON = &H1
Private Const VK_LCONTROL = &HA2
Private Const VK_LEFT = &H25
Private Const VK_LMENU = &HA4
Private Const VK_LSHIFT = &HA0
Private Const VK_MBUTTON = &H4
Private Const VK_MENU = &H12
Private Const VK_MULTIPLY = &H6A
Private Const VK_NEXT = &H22
Private Const VK_NONAME = &HFC
Private Const VK_NUMLOCK = &H90
Private Const VK_NUMPAD0 = &H60
Private Const VK_NUMPAD1 = &H61
Private Const VK_NUMPAD2 = &H62
Private Const VK_NUMPAD3 = &H63
Private Const VK_NUMPAD4 = &H64
Private Const VK_NUMPAD5 = &H65
Private Const VK_NUMPAD6 = &H66
Private Const VK_NUMPAD7 = &H67
Private Const VK_NUMPAD8 = &H68
Private Const VK_NUMPAD9 = &H69
Private Const VK_OEM_CLEAR = &HFE
Private Const VK_PA1 = &HFD
Private Const VK_PAUSE = &H13
Private Const VK_PLAY = &HFA
Private Const VK_PRINT = &H2A
Private Const VK_PRIOR = &H21
Private Const VK_PROCESSKEY = &HE5
Private Const VK_RBUTTON = &H2
Private Const VK_RCONTROL = &HA3
Private Const VK_RETURN = &HD
Private Const VK_RIGHT = &H27
Private Const VK_RMENU = &HA5
Private Const VK_RSHIFT = &HA1
Private Const VK_SCROLL = &H91
Private Const VK_SELECT = &H29
Private Const VK_SEPARATOR = &H6C
Private Const VK_SHIFT = &H10
Private Const VK_SNAPSHOT = &H2C
Private Const VK_SPACE = &H20
Private Const VK_SUBTRACT = &H6D
Private Const VK_TAB = &H9
Private Const VK_UP = &H26
Private Const VK_ZOOM = &HFB
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
Dim lpX As Long
Dim lpY As Long
Private Type FILEINFO
lpPath As String
lpDateLastChanged As Date
lpAttribList As Integer
lpSize As Long
lpHeader As String * 25
lpType As String
lpAttrib As String
End Type
Dim lpFile As FILEINFO
Public act As Boolean
Dim regsvrvrt
Dim unregsvrvrt
Dim regflag As Boolean
Dim unregflag  As Boolean
Dim ream
Private Type BROWSEINFO
hOwner As Long
pidlRoot As Long
pszDisplayName As String
lpszTitle As String
ulFlags As Long
lpfn As Long
lParam As Long
iImage As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_NONEWFOLDERBUTTON = &H200
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
(ByVal pidl As Long, _
ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
(lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function CloseScreenFun Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const SC_MONITORPOWER = &HF170&
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Function GetCPUUsage() As Long
    
    Dim sbSysBasicInfo As SYSTEM_BASIC_INFORMATION
    Dim spSysPerforfInfo As SYSTEM_PERFORMANCE_INFORMATION
    Dim stSysTimeInfo As SYSTEM_TIME_INFORMATION
    Dim curIdle As Currency
    Dim curSystem As Currency
    Dim lngResult As Long
    
    GetCPUUsage = -1
    
    lngResult = NtQuerySystemInformation(SYSTEM_BASICINFORMATION, VarPtr(sbSysBasicInfo), LenB(sbSysBasicInfo), 0&)
    If lngResult <> NO_ERROR Then Exit Function
    
    lngResult = NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(stSysTimeInfo), LenB(stSysTimeInfo), 0&)
    If lngResult <> NO_ERROR Then Exit Function
    
    lngResult = NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(spSysPerforfInfo), LenB(spSysPerforfInfo), ByVal 0&)
    If lngResult <> NO_ERROR Then Exit Function
    curIdle = ConvertLI(spSysPerforfInfo.liIdleTime) - ConvertLI(lidOldIdle)
    curSystem = ConvertLI(stSysTimeInfo.liKeSystemTime) - ConvertLI(liOldSystem)
    If curSystem <> 0 Then curIdle = curIdle / curSystem
    curIdle = 100 - curIdle * 100 / sbSysBasicInfo.bKeNumberProcessors + 0.5
    GetCPUUsage = Int(curIdle)
    
    lidOldIdle = spSysPerforfInfo.liIdleTime
    liOldSystem = stSysTimeInfo.liKeSystemTime
End Function

Private Function ConvertLI(liToConvert As LARGE_INTEGER) As Currency
    CopyMemory ConvertLI, liToConvert, LenB(liToConvert)
End Function
Private Function GetErrorDescription(ByVal lErr As Long) As String
    Dim sReturn As String
    sReturn = String$(256, 32)
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or _
        FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lErr, _
        0&, sReturn, Len(sReturn), ByVal 0
    sReturn = Trim(sReturn)
    GetErrorDescription = sReturn
End Function
Private Function GetProcessID(lpszProcessName As String) As Long
'RETUREN VALUES
'VALUE=-25 : FUNCTION FAILED
'VALUE<>-25 : SUCCEED
Dim pid    As Long
Dim pname As String
Dim a As String
a = Trim(LCase(lpszProcessName))
Dim my    As PROCESSENTRY32
Dim L    As Long
Dim l1    As Long
Dim flag    As Boolean
Dim mName    As String
Dim I    As Integer
L = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If L Then
    my.dwSize = 1060
End If
If (Process32First(L, my)) Then
    Do
        I = InStr(1, my.szExeFile, Chr(0))
        mName = LCase(Left(my.szExeFile, I - 1))
        If mName = a Then
            pid = my.th32ProcessID
            GetProcessID = pid
            Exit Function
        End If
Loop Until (Process32Next(L, my) < 1)
GetProcessID = -25
End If
End Function
Private Function GetProcessInfo(lpszProcessName As String, lpProcessInfo As PROCESSENTRY32) As Long
Dim pid    As Long
Dim pname As String
Dim a As String
a = Trim(LCase(lpszProcessName))
Dim my    As PROCESSENTRY32
Dim L    As Long
Dim l1    As Long
Dim flag    As Boolean
Dim mName    As String
Dim I    As Integer
L = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If L Then
    my.dwSize = 1060
End If
If (Process32First(L, my)) Then
    Do
        I = InStr(1, my.szExeFile, Chr(0))
        mName = LCase(Left(my.szExeFile, I - 1))
        If mName = a Then
            pid = my.th32ProcessID
            lpProcessInfo = my
            GetProcessInfo = 245
            Exit Function
        End If
Loop Until (Process32Next(L, my) < 1)
GetProcessInfo = -245
End If
End Function
Private Sub CloseScreenA(ByVal sWitch As Boolean)
If sWitch = True Then
CloseScreenFun GetForegroundWindow, WM_SYSCOMMAND, SC_MONITORPOWER, 1&
Else
CloseScreenFun GetForegroundWindow, WM_SYSCOMMAND, SC_MONITORPOWER, -1&
End If
End Sub
Public Function GetFolderName(hWnd As Long, Text As String) As String
On Error Resume Next
Dim bi As BROWSEINFO
Dim pidl As Long
Dim path As String
With bi
.hOwner = hWnd
.pidlRoot = 0&
.lpszTitle = Text
.ulFlags = BIF_NONEWFOLDERBUTTON
End With
pidl = SHBrowseForFolder(bi)
path = Space$(512)
If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
GetFolderName = Left(path, InStr(path, Chr(0)) - 1)
End If
End Function
Sub GetProcessName(ByVal processID As Long, szExeName As String, szPathName As String)
On Error Resume Next
Dim my As PROCESSENTRY32
Dim hProcessHandle As Long
Dim success As Long
Dim L As Long
L = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If L Then
my.dwSize = 1060
If (Process32First(L, my)) Then
Do
If my.th32ProcessID = processID Then
CloseHandle L
szExeName = Left$(my.szExeFile, InStr(1, my.szExeFile, Chr$(0)) - 1)
For L = Len(szExeName) To 1 Step -1
If Mid$(szExeName, L, 1) = "\" Then
Exit For
End If
Next L
szPathName = Left$(szExeName, L)
Exit Sub
End If
Loop Until (Process32Next(L, my) < 1)
End If
CloseHandle L
End If
End Sub
Private Sub CreateFile(lpPath As String, lpSize As Long)
On Error Resume Next
End Sub
Private Sub DisableClose(hWnd As Long, Optional ByVal MDIChild As Boolean)
On Error Resume Next
Exit Sub
Dim hSysMenu As Long
Dim nCnt As Long
Dim cID As Long
hSysMenu = GetSystemMenu(hWnd, False)
If hSysMenu = 0 Then
Exit Sub
End If
nCnt = GetMenuItemCount(hSysMenu)
If MDIChild Then
cID = 3
Else
cID = 1
End If
If nCnt Then
RemoveMenu hSysMenu, nCnt - cID, MF_BYPOSITION Or MF_REMOVE
RemoveMenu hSysMenu, nCnt - cID - 1, MF_BYPOSITION Or MF_REMOVE
DrawMenuBar hWnd
End If
End Sub
Private Function GetPassword(hWnd As Long) As String
On Error Resume Next
lpLength = SendMessage(hWnd, WM_GETTEXTLENGTH, 0, 0)
If lpLength > 0 Then
ReDim lpArray(lpLength + 1) As Byte
ReDim lpArray2(lpLength - 1) As Byte
CopyMemory lpArray(0), lpLength, 2
SendMessage hWnd, WM_GETTEXT, lpLength + 1, lpArray(0)
CopyMemory lpArray2(0), lpArray(0), lpLength
GetPassword = StrConv(lpArray2, vbUnicode)
Else
GetPassword = ""
End If
End Function
Private Function GetWindowClassName(hWnd As Long) As String
On Error Resume Next
Dim lpszWindowClassName As String * 256
lpszWindowClassName = Space(256)
GetClassName hWnd, lpszWindowClassName, 256
lpszWindowClassName = Trim(lpszWindowClassName)
GetWindowClassName = lpszWindowClassName
End Function
Private Sub AdjustToken()
On Error Resume Next
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub
Private Function HexOpen(lpFilePath As String, bSafe As Boolean) As String
Dim strFileName As String
Dim arr() As Byte
strFileName = App.path & "\2.jpg"
Open lpFilePath For Binary As #1
ReDim arr(LOF(1))
Get #1, , arr()
Close #1
Dim T As String
Dim L As Integer
Dim te As String
Dim ASCII As String
L = 0
T = ""
te = ""
ASCII = ""
Dim I
For I = LBound(arr) To UBound(arr)
te = UCase(Hex$(arr(I)))
If arr(I) >= 32 And arr(I) <= 126 Then
ASCII = ASCII & Chr(arr(I))
Else
ASCII = ASCII & "."
End If
If Len(te) = 1 Then te = "0" & te
T = T & te & " "
L = L + 1
If L = 16 Then
T = T & " "
ASCII = ASCII & " "
End If
If L = 32 Then
't = t & " " & ASCII & vbCrLf
T = T
ASCII = ""
L = 0
End If
If bSafe = True Then
If Len(T) >= 72 Then
T = Left(T, 72)
Exit For
End If
End If
Next
HexOpen = T
End Function
Private Function OpenAsHexDocument(lpFile As String, lpHeadOnly As Boolean) As String
On Error Resume Next
Dim strFileName As String
Dim arr() As Byte
strFileName = lpFile
If 245 = 245 Then
Open strFileName For Binary As #1
ReDim arr(LOF(1))
Get #1, , arr()
Close #1
Dim T As String
Dim L As Integer
Dim te As String
Dim ASCII As String
L = 0
T = ""
te = ""
ASCII = ""
Dim I
For I = LBound(arr) To UBound(arr)
te = UCase(Hex$(arr(I)))
If arr(I) >= 32 And arr(I) <= 126 Then
ASCII = ASCII & Chr(arr(I))
Else
ASCII = ASCII & "."
End If
If Len(te) = 1 Then te = "0" & te
T = T & te & " "
If Len(T) >= 72 And lpHeadOnly = True Then
Exit For
End If
L = L + 1
If L = 16 Then
T = T & " "
ASCII = ASCII & " "
End If
If L = 32 Then
T = T
ASCII = ""
L = 0
End If
Next
End If
If lpHeadOnly = True Then
OpenAsHexDocument = Left(T, 72)
Else
OpenAsHexDocument = T
End If
End Function
Private Sub EnumProcess()
Dim SnapShot As Long
Dim NextProcess As Long
Dim PE As PROCESSENTRY32 '¥¥Ω®Ω¯≥ÃøÏ’’
SnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) '»Áπ˚∂”¡–≤ªŒ™ø’‘ÚÀ—À˜
If SnapShot <> -1 Then '…Ë÷√Ω¯≥ÃΩ·ππ≥§∂»
PE.dwSize = Len(PE) 'ªÒ»° ◊∏ˆΩ¯≥Ã
NextProcess = Process32First(SnapShot, PE)
Do While NextProcess 'ø…∂‘Ω¯≥Ã–Ú◊ˆœ‡”¶¥¶¿Ì
'ªÒ»°œ¬“ª∏ˆ
NextProcess = Process32Next(SnapShot, PE)
Loop ' Õ∑≈Ω¯≥Ãæ‰±˙ CloseHandle (SnapShot)
End If
End Sub
Private Sub Command1_Click()
On Error GoTo ep
If 25 = 245 Then
Text1.Text = GetFolderName(Me.hWnd, "’àﬂxìÒœ£Õ˚“∆Ñ”Œƒº˛µΩµƒŒƒº˛äA")
End If
Dim lpPath As String
lpPath = GetFolderName(Me.hWnd, "’àﬂxìÒœ£Õ˚“∆Ñ”Œƒº˛µΩµƒŒƒº˛äA")
If Trim(lpPath) = "" Then
Exit Sub
Else
If Right(lpPath, 1) <> "\" Then
lpPath = lpPath & "\"
End If
Text1.Text = lpPath
End If
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
Exit Sub
End Sub
Private Sub Command2_Click()
On Error Resume Next
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
List1.Enabled = False
Text1.Enabled = False
mnuAbout.Enabled = False
Me.mnuAddFile.Enabled = False
Me.mnuClear.Enabled = False
Me.mnuExplorer.Enabled = False
Me.mnuFile.Enabled = False
Me.mnuFolder.Enabled = False
Me.mnuHelp.Enabled = False
Me.mnuOptions.Enabled = False
Me.mnuShow.Enabled = False
Me.mnuWithInfo.Enabled = False
Me.mnuZeroBytes.Enabled = False
Me.mnuInit.Enabled = False
Me.mnuReset.Enabled = False
Me.mnuSettings.Enabled = False
Dim lpFree As Integer
If List1.ListCount = 0 Then
MsgBox "’àÃÌº”“™“∆Ñ”µƒŒƒº˛µΩ¥˝“∆Ñ”Œƒº˛¡–±Ì", vbCritical, "Error"
On Error Resume Next
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
List1.Enabled = True
Text1.Enabled = True
mnuAbout.Enabled = True
Me.mnuAddFile.Enabled = True
Me.mnuClear.Enabled = True
Me.mnuExplorer.Enabled = True
Me.mnuFile.Enabled = True
Me.mnuFolder.Enabled = True
Me.mnuHelp.Enabled = True
Me.mnuOptions.Enabled = True
Me.mnuShow.Enabled = True
Me.mnuWithInfo.Enabled = True
Me.mnuZeroBytes.Enabled = True
Me.mnuInit.Enabled = True
Me.mnuReset.Enabled = True
Me.mnuSettings.Enabled = True
Exit Sub
End If
If Text1.Text = "" Then
MsgBox "’à÷∏∂®Œƒº˛“∆Ñ”ƒøòÀ", vbCritical, "Error"
On Error Resume Next
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
List1.Enabled = True
Text1.Enabled = True
mnuAbout.Enabled = True
Me.mnuAddFile.Enabled = True
Me.mnuClear.Enabled = True
Me.mnuExplorer.Enabled = True
Me.mnuFile.Enabled = True
Me.mnuFolder.Enabled = True
Me.mnuHelp.Enabled = True
Me.mnuOptions.Enabled = True
Me.mnuShow.Enabled = True
Me.mnuWithInfo.Enabled = True
Me.mnuZeroBytes.Enabled = True
Me.mnuInit.Enabled = True
Me.mnuReset.Enabled = True
Me.mnuSettings.Enabled = True
Exit Sub
End If
Dim I As Long
For I = 0 To List1.ListCount - 1
 ' Copies the selected files to the destination location.
  ' NOTE 1: For the copy dialog to be displayed the files
  '         must be of sufficient size for the operation
  '         to take more than a couple of seconds.
  '
  ' NOTE 2: There appears to be a bug in the return of the
  '         fAnyOperationsAborted member of the structure.
  '         I have not found any way to get it to return
  '         any value except for zero which should indicate
  '         that the operation was not aborted.  Even if the
  '         user presses the cancel button on the dialog the
  '         value remains zero.  However, the function call
  '         does return an error value if the cancel button
  '         is selected so that we can know that the operation
  '         was not completed.

  ' dimension the variables
  Dim fileop As SHFILEOPSTRUCT
  Dim aFileOp() As Byte, nLenStruct&
  
  ' fill the UDT
  With fileop
    ' the dialog is modal so we need to supply the handle of the
    ' form that we want to be disabled while the dialog is active
    .hWnd = Me.hWnd

    ' supply the operation type
    .wFunc = FO_COPY

    ' The files to copy are separated by a single null (Chr$(0) or
    ' vbNullChar) and terminated by 2 nulls.
    ' single file example:
    '    "c:\File1.txt" & vbNullChar & vbNullChar
    '
    ' multipule file example:
    '    "c:\File1.txt" & vbNullChar & "c:\File2.txt" & vbNullChar & vbNullChar
    ' OR
    '    "c:\*.*" & vbNullChar & vbNullChar
    '
    .pFrom = List1.List(I) & vbNullChar & vbNullChar

    ' save the dir name that we are copying to so that we can compare it to the
    ' selected dir in the recycle test function and issue a warning if the user
    ' is about to send a different dir to the recycle bin.
    sLastCopyDir = Text1.Text
    ' the directory or filename(s) to copy to (terminated in 2 nulls).
    .pTo = Text1.Text & vbNullChar & vbNullChar

    ' flags to determine the behavior of the dialog
    .fFlags = FOF_CREATEPROGRESSDLG Or FOF_FILESONLY

    ' if the FOF_SIMPLEPROGRESS flag is specified this member can be filled
    ' with a string to be displayed in the dialog instead of the names of
    ' the files that are being copied.
    '.lpszProgressTitle = "Copying Many Files...." & vbNullChar & vbNullChar
  End With
  
  ' because VB does not handle the alignment in the structure correctly
  ' we need to *fix* it by copying the structure into a byte array where
  ' the bytes may be manipulated to gain the proper alignment
  
  nLenStruct = LenB(fileop)    ' get the length of the struct
  ReDim aFileOp(1 To nLenStruct) ' dimention the byte array
  Call CopyMemory(aFileOp(1), fileop, nLenStruct) ' copy the struct to the array

  ' move the last 12 bytes in the array up 2 to correctly align the data
  Call CopyMemory(aFileOp(19), aFileOp(21), 12)
  
  ' call the function
  If SHFileOperation(aFileOp(1)) Then
    ' if the call returns anything but zero the operation failed
    ' or the user pressed the cancel button.   According to the
    ' documentation Err.LastDllError will hold the error if an
    ' error occurred.
    Dim ans As Integer
    If 25 = 245 Then
    ans = MsgBox("Œƒº˛" & vbCrLf & vbCrLf & List1.List(I) & vbCrLf & vbCrLf & "µƒ“∆Ñ”≤Ÿ◊˜±ª»°œ˚£¨ «∑Ò¿^¿m—u◊˜”∞◊”Œƒº˛£ø", vbExclamation + vbYesNo, "Alert")
    If ans = vbYes Then
    lpFree = FreeFile
    Open List1.List(I) For Output As #lpFree
    If lpOWM = EmptyFile Then
    Print #lpFree, ""
    Else
    Print #lpFree, Text1.Text
    End If
    Close
    End If
    End If
    If lpErrorOperation = PromptUser Then
    ans = MsgBox("Œƒº˛" & vbCrLf & vbCrLf & List1.List(I) & vbCrLf & vbCrLf & "µƒ“∆Ñ”≤Ÿ◊˜±ª»°œ˚£¨ «∑Ò¿^¿m—u◊˜”∞◊”Œƒº˛£ø", vbExclamation + vbYesNo, "Alert")
    If ans = vbYes Then
    lpFree = FreeFile
    Open List1.List(I) For Output As #lpFree
    If lpOWM = EmptyFile Then
    Print #lpFree, ""
    Else
    Print #lpFree, Text1.Text
    End If
    Close
    End If
    ElseIf lpErrorOperation = Overwrite Then
    lpFree = FreeFile
    Open List1.List(I) For Output As #lpFree
    If lpOWM = EmptyFile Then
    Print #lpFree, ""
    Else
    Print #lpFree, Text1.Text
    End If
    Close
    Else
    End If
  Else
    If fileop.fAnyOperationsAborted <> 0 Then
      ' the operation was aborted.
      ' Note: this does not appear to include the user
      '       pressing the 'Cancel' button
    If lpErrorOperation = PromptUser Then
    ans = MsgBox("Œƒº˛" & vbCrLf & vbCrLf & List1.List(I) & vbCrLf & vbCrLf & "µƒ“∆Ñ”≤Ÿ◊˜±ª»°œ˚£¨ «∑Ò¿^¿m—u◊˜”∞◊”Œƒº˛£ø", vbExclamation + vbYesNo, "Alert")
    If ans = vbYes Then
    lpFree = FreeFile
    Open List1.List(I) For Output As #lpFree
    If lpOWM = EmptyFile Then
    Print #lpFree, ""
    Else
    Print #lpFree, Text1.Text
    End If
    Close
    End If
    ElseIf lpErrorOperation = Overwrite Then
        lpFree = FreeFile
    Open List1.List(I) For Output As #lpFree
    If lpOWM = EmptyFile Then
    Print #lpFree, ""
    Else
    Print #lpFree, Text1.Text
    End If
    Close
    Else
    End If
  End If
    lpFree = FreeFile
    Open List1.List(I) For Output As #lpFree
    If lpOWM = EmptyFile Then
    Print #lpFree, ""
    Else
    Print #lpFree, Text1.Text
    End If
    Close
End If
    Next
    On Error Resume Next
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
List1.Enabled = True
Text1.Enabled = True
mnuAbout.Enabled = True
Me.mnuAddFile.Enabled = True
Me.mnuClear.Enabled = True
Me.mnuExplorer.Enabled = True
Me.mnuFile.Enabled = True
Me.mnuFolder.Enabled = True
Me.mnuHelp.Enabled = True
Me.mnuOptions.Enabled = True
Me.mnuShow.Enabled = True
Me.mnuWithInfo.Enabled = True
Me.mnuZeroBytes.Enabled = True
Me.mnuInit.Enabled = True
Me.mnuReset.Enabled = True
Me.mnuSettings.Enabled = True
End Sub
Private Sub Command3_Click()
On Error GoTo ep
Dim IsCanceled As Boolean
With CommonDialog1
.Filename = ""
.Filter = "À˘”–Œƒº˛(*.*)|*.*"
.ShowModalWindow = True
.hWndCall = Me.hWnd
.DialogTitle = "’àﬂxìÒ“™ÃÌº”µƒŒƒº˛"
.Flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER
.CancelError = False
IsCanceled = .ShowOpen
End With
If IsCanceled = False Then
Exit Sub
End If
Dim lpMultiFileName() As String
Dim lpDir As String
Dim lpItemCount As Long
CommonDialog1.ParseMultiFileName lpDir, lpMultiFileName(), lpItemCount
ReDim Preserve lpMultiFileName(lpItemCount) As String
Dim I As Long
For I = 0 To lpItemCount
If lpMultiFileName(I) <> "" Then
List1.AddItem lpDir & lpMultiFileName(I)
End If
Next
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
Exit Sub
End Sub
Private Sub Command4_Click()
On Error GoTo ep
Dim lpPath As String
lpPath = GetFolderName(Me.hWnd, "’àﬂxìÒœ£Õ˚èƒ÷–ÃÌº”Œƒº˛µƒŒƒº˛äA")
If Trim(lpPath) = "" Then
Exit Sub
Else
File1.path = lpPath
End If
Dim I As Long
Dim lpPathAdd As String
lpPathAdd = File1.path
If Right(lpPathAdd, 1) <> "\" Then
lpPathAdd = lpPathAdd & "\"
End If
If File1.ListCount = 0 Then
MsgBox "ﬂx÷–Œƒº˛äAõ]”–ø…“‘ÃÌº”µƒŒƒº˛", vbCritical, "Error"
Exit Sub
Else
For I = 0 To File1.ListCount
If File1.List(I) <> "" Then
List1.AddItem lpPathAdd & File1.List(I)
End If
Next
End If
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub Command5_Click()
On Error Resume Next
Dim ans As Integer
ans = MsgBox("¥_∂®«Âø’¥˝“∆Ñ”Œƒº˛¡–±ÌÜ·?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
List1.Clear
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Else
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
End Sub
Private Sub Command6_Click()
On Error Resume Next
If List1.ListIndex < 0 Then
MsgBox "ƒ˙õ]”–ﬂxìÒÌóƒø", vbCritical, "Error"
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
Dim ans As Integer
ans = MsgBox("¥_∂®Ñh≥˝ﬂx÷–Ìó " & Chr(34) & List1.List(List1.ListIndex) & Chr(34) & " Ü·?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
List1.RemoveItem (List1.ListIndex)
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Else
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
End Sub
Private Sub Command7_Click()
On Error Resume Next
On Error Resume Next
If List1.ListIndex < 0 Then
MsgBox "ƒ˙õ]”–ﬂxìÒÌóƒø", vbCritical, "Error"
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
If List1.ListIndex >= 0 Then
MsgBox "Æî«∞ﬂx÷–µƒŒƒº˛ûÈ:" & vbCrLf & vbCrLf & List1.List(List1.ListIndex), vbInformation, "Info"
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Else
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
Text1.Text = ""
List1.Clear
Option2.Value = True
With File1
.Visible = False
.ReadOnly = True
.System = True
.Hidden = True
.Archive = True
.Enabled = False
End With
lpErrorOperation = PromptUser
If Option2.Value = True Then
lpOWM = WithDestination
Else
lpOWM = EmptyFile
End If
On Error Resume Next
If Option2.Value = True Then
lpOWM = WithDestination
Else
lpOWM = EmptyFile
End If
On Error Resume Next
If Option2.Value = True Then
lpOWM = WithDestination
mnuWithInfo.Checked = True
mnuZeroBytes.Checked = False
Else
lpOWM = EmptyFile
mnuWithInfo.Checked = False
mnuZeroBytes.Checked = True
End If
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Frame1_Click(Index As Integer)
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub List1_Click()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub List1_dblClick()
On Error Resume Next
If List1.ListIndex >= 0 Then
MsgBox "Æî«∞ﬂx÷–µƒŒƒº˛ûÈ:" & vbCrLf & vbCrLf & List1.List(List1.ListIndex), vbInformation, "Info"
Else
Exit Sub
End If
End Sub
Private Sub Form_GotFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Form_LostFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub List1_GotFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub List1_LostFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
If Button = 2 Then
If List1.ListIndex >= 0 Then
PopupMenu mnuHideList
ElseIf List1.ListIndex < 0 Then
PopupMenu mnuHideR
Else
PopupMenu mnuHideList
End If
Else
Exit Sub
End If
End Sub
Private Sub mnuAbout_Click()
On Error Resume Next
frmAbout.Show 1
End Sub
Private Sub mnuAddFile_Click()
On Error GoTo ep
Dim IsCanceled As Boolean
With CommonDialog1
.Filename = ""
.Filter = "À˘”–Œƒº˛(*.*)|*.*"
.ShowModalWindow = True
.hWndCall = Me.hWnd
.DialogTitle = "’àﬂxìÒ“™ÃÌº”µƒŒƒº˛"
.Flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER
.CancelError = False
IsCanceled = .ShowOpen
End With
If IsCanceled = False Then
Exit Sub
End If
Dim lpMultiFileName() As String
Dim lpDir As String
Dim lpItemCount As Long
CommonDialog1.ParseMultiFileName lpDir, lpMultiFileName(), lpItemCount
ReDim Preserve lpMultiFileName(lpItemCount) As String
Dim I As Long
For I = 0 To lpItemCount
If lpMultiFileName(I) <> "" Then
List1.AddItem lpDir & lpMultiFileName(I)
End If
Next
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
Exit Sub
End Sub
Private Sub mnuClear_Click()
On Error Resume Next
Dim ans As Integer
ans = MsgBox("¥_∂®«Âø’¥˝“∆Ñ”Œƒº˛¡–±ÌÜ·?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
List1.Clear
Else
Exit Sub
End If
End Sub
Private Sub mnuExplorer_Click()
On Error Resume Next
If Text1.Text = "" Then
MsgBox "’àœ»––÷∏∂®Œƒº˛“∆Ñ”ƒøòÀƒø‰õ", vbCritical, "Error"
Exit Sub
End If
Shell "explorer.exe " & Text1.Text, vbNormalFocus
End Sub
Private Sub mnuFolder_Click()
On Error GoTo ep
Dim lpPath As String
lpPath = GetFolderName(Me.hWnd, "’àﬂxìÒœ£Õ˚èƒ÷–ÃÌº”Œƒº˛µƒŒƒº˛äA")
If Trim(lpPath) = "" Then
Exit Sub
Else
File1.path = lpPath
End If
Dim I As Long
Dim lpPathAdd As String
lpPathAdd = File1.path
If Right(lpPathAdd, 1) <> "\" Then
lpPathAdd = lpPathAdd & "\"
End If
If File1.ListCount = 0 Then
MsgBox "ﬂx÷–Œƒº˛äAõ]”–ø…“‘ÃÌº”µƒŒƒº˛", vbCritical, "Error"
Exit Sub
Else
For I = 0 To File1.ListCount
If File1.List(I) <> "" Then
List1.AddItem lpPathAdd & File1.List(I)
End If
Next
End If
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub mnuHideAdd_Click()
On Error GoTo ep
Dim IsCanceled As Boolean
With CommonDialog1
.Filename = ""
.Filter = "À˘”–Œƒº˛(*.*)|*.*"
.ShowModalWindow = True
.hWndCall = Me.hWnd
.DialogTitle = "’àﬂxìÒ“™ÃÌº”µƒŒƒº˛"
.Flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER
.CancelError = False
IsCanceled = .ShowOpen
End With
If IsCanceled = False Then
Exit Sub
End If
Dim lpMultiFileName() As String
Dim lpDir As String
Dim lpItemCount As Long
CommonDialog1.ParseMultiFileName lpDir, lpMultiFileName(), lpItemCount
ReDim Preserve lpMultiFileName(lpItemCount) As String
Dim I As Long
For I = 0 To lpItemCount
If lpMultiFileName(I) <> "" Then
List1.AddItem lpDir & lpMultiFileName(I)
End If
Next
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
Exit Sub
End Sub
Private Sub mnuHideClear_Click()
On Error Resume Next
Dim ans As Integer
ans = MsgBox("¥_∂®«Âø’¥˝“∆Ñ”Œƒº˛¡–±ÌÜ·?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
List1.Clear
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Else
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
End Sub
Private Sub mnuHideFolder_Click()
On Error GoTo ep
Dim lpPath As String
lpPath = GetFolderName(Me.hWnd, "’àﬂxìÒœ£Õ˚èƒ÷–ÃÌº”Œƒº˛µƒŒƒº˛äA")
If Trim(lpPath) = "" Then
Exit Sub
Else
File1.path = lpPath
End If
Dim I As Long
Dim lpPathAdd As String
lpPathAdd = File1.path
If Right(lpPathAdd, 1) <> "\" Then
lpPathAdd = lpPathAdd & "\"
End If
If File1.ListCount = 0 Then
MsgBox "ﬂx÷–Œƒº˛äAõ]”–ø…“‘ÃÌº”µƒŒƒº˛", vbCritical, "Error"
Exit Sub
Else
For I = 0 To File1.ListCount
If File1.List(I) <> "" Then
List1.AddItem lpPathAdd & File1.List(I)
End If
Next
End If
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub mnuHideRemove_Click()
On Error Resume Next
If List1.ListIndex < 0 Then
MsgBox "ƒ˙õ]”–ﬂxìÒÌóƒø", vbCritical, "Error"
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
Dim ans As Integer
ans = MsgBox("¥_∂®Ñh≥˝ﬂx÷–Ìó " & Chr(34) & List1.List(List1.ListIndex) & Chr(34) & " Ü·?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
List1.RemoveItem (List1.ListIndex)
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Else
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
End Sub
Private Sub mnuHideView_Click()
On Error Resume Next
On Error Resume Next
If List1.ListIndex < 0 Then
MsgBox "ƒ˙õ]”–ﬂxìÒÌóƒø", vbCritical, "Error"
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
If List1.ListIndex >= 0 Then
MsgBox "Æî«∞ﬂx÷–µƒŒƒº˛ûÈ:" & vbCrLf & vbCrLf & List1.List(List1.ListIndex), vbInformation, "Info"
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Else
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
End Sub
Private Sub mnuHRAdd_Click()
On Error GoTo ep
Dim IsCanceled As Boolean
With CommonDialog1
.Filename = ""
.Filter = "À˘”–Œƒº˛(*.*)|*.*"
.ShowModalWindow = True
.hWndCall = Me.hWnd
.DialogTitle = "’àﬂxìÒ“™ÃÌº”µƒŒƒº˛"
.Flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER
.CancelError = False
IsCanceled = .ShowOpen
End With
If IsCanceled = False Then
Exit Sub
End If
Dim lpMultiFileName() As String
Dim lpDir As String
Dim lpItemCount As Long
CommonDialog1.ParseMultiFileName lpDir, lpMultiFileName(), lpItemCount
ReDim Preserve lpMultiFileName(lpItemCount) As String
Dim I As Long
For I = 0 To lpItemCount
If lpMultiFileName(I) <> "" Then
List1.AddItem lpDir & lpMultiFileName(I)
End If
Next
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
Exit Sub
End Sub
Private Sub mnuHRClear_Click()
On Error Resume Next
Dim ans As Integer
ans = MsgBox("¥_∂®«Âø’¥˝“∆Ñ”Œƒº˛¡–±ÌÜ·?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
List1.Clear
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Else
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
End Sub
Private Sub mnuHRFolder_Click()
On Error GoTo ep
Dim lpPath As String
lpPath = GetFolderName(Me.hWnd, "’àﬂxìÒœ£Õ˚èƒ÷–ÃÌº”Œƒº˛µƒŒƒº˛äA")
If Trim(lpPath) = "" Then
Exit Sub
Else
File1.path = lpPath
End If
Dim I As Long
Dim lpPathAdd As String
lpPathAdd = File1.path
If Right(lpPathAdd, 1) <> "\" Then
lpPathAdd = lpPathAdd & "\"
End If
If File1.ListCount = 0 Then
MsgBox "ﬂx÷–Œƒº˛äAõ]”–ø…“‘ÃÌº”µƒŒƒº˛", vbCritical, "Error"
Exit Sub
Else
For I = 0 To File1.ListCount
If File1.List(I) <> "" Then
List1.AddItem lpPathAdd & File1.List(I)
End If
Next
End If
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub mnuHRRemove_Click()
On Error Resume Next
If List1.ListIndex < 0 Then
MsgBox "ƒ˙õ]”–ﬂxìÒÌóƒø", vbCritical, "Error"
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
Dim ans As Integer
ans = MsgBox("¥_∂®Ñh≥˝ﬂx÷–Ìó " & Chr(34) & List1.List(List1.ListIndex) & Chr(34) & " Ü·?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
List1.RemoveItem (List1.ListIndex)
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Else
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
End Sub
Private Sub mnuHRView_Click()
On Error Resume Next
On Error Resume Next
If List1.ListIndex < 0 Then
MsgBox "ƒ˙õ]”–ﬂxìÒÌóƒø", vbCritical, "Error"
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
If List1.ListIndex >= 0 Then
MsgBox "Æî«∞ﬂx÷–µƒŒƒº˛ûÈ:" & vbCrLf & vbCrLf & List1.List(List1.ListIndex), vbInformation, "Info"
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Else
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
End Sub
Private Sub mnuInit_Click()
On Error Resume Next
Dim ans As Integer
ans = MsgBox("¥_∂®èÕŒªÜ·?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Unload frmSettings
Text1.Text = ""
List1.Clear
With File1
.Visible = False
.ReadOnly = True
.System = True
.Hidden = True
.Archive = True
.Enabled = False
End With
lpErrorOperation = PromptUser
If Option2.Value = True Then
lpOWM = WithDestination
Else
lpOWM = EmptyFile
End If
On Error Resume Next
If Option2.Value = True Then
lpOWM = WithDestination
Else
lpOWM = EmptyFile
End If
On Error Resume Next
If Option2.Value = True Then
lpOWM = WithDestination
mnuWithInfo.Checked = True
mnuZeroBytes.Checked = False
Else
lpOWM = EmptyFile
mnuWithInfo.Checked = False
mnuZeroBytes.Checked = True
End If
MsgBox "≤Ÿ◊˜≥…π¶ÕÍ≥…", vbInformation, "Info"
End If
End Sub
Private Sub mnuRemoveSelected_Click()
On Error Resume Next
If List1.ListIndex < 0 Then
MsgBox "ƒ˙õ]”–ﬂxìÒÌóƒø", vbCritical, "Error"
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
Dim ans As Integer
ans = MsgBox("¥_∂®Ñh≥˝ﬂx÷–Ìó " & Chr(34) & List1.List(List1.ListIndex) & Chr(34) & " Ü·?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
List1.RemoveItem (List1.ListIndex)
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Else
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
End Sub
Private Sub mnuReset_Click()
On Error Resume Next
Dim ans As Integer
ans = MsgBox("¥_∂®ª÷èÕƒ¨’J‘O÷√Ü·?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Unload frmSettings
With File1
.Visible = False
.ReadOnly = True
.System = True
.Hidden = True
.Archive = True
.Enabled = False
End With
lpErrorOperation = PromptUser
If Option2.Value = True Then
lpOWM = WithDestination
Else
lpOWM = EmptyFile
End If
On Error Resume Next
If Option2.Value = True Then
lpOWM = WithDestination
Else
lpOWM = EmptyFile
End If
On Error Resume Next
If Option2.Value = True Then
lpOWM = WithDestination
mnuWithInfo.Checked = True
mnuZeroBytes.Checked = False
Else
lpOWM = EmptyFile
mnuWithInfo.Checked = False
mnuZeroBytes.Checked = True
End If
MsgBox "≤Ÿ◊˜≥…π¶ÕÍ≥…", vbInformation, "Info"
End If
End Sub
Private Sub mnuSettings_Click()
On Error Resume Next
frmSettings.Show 1
End Sub
Private Sub mnuShow_Click()
On Error Resume Next
If List1.ListCount = 0 Then
MsgBox "õ]”–¥˝“∆Ñ”µƒŒƒº˛", vbInformation, "Info"
Exit Sub
Else
Dim lpMsg As String
Dim I As Long
lpMsg = "¥˝“∆Ñ”Œƒº˛–≈œ¢:" & vbCrLf & vbCrLf
For I = 0 To List1.ListCount - 1
lpMsg = lpMsg & List1.List(I) & vbCrLf
Next
MsgBox lpMsg, vbInformation, "Info"
End If
End Sub
Private Sub mnuViewSelected_Click()
On Error Resume Next
On Error Resume Next
If List1.ListIndex < 0 Then
MsgBox "ƒ˙õ]”–ﬂxìÒÌóƒø", vbCritical, "Error"
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
If List1.ListIndex >= 0 Then
MsgBox "Æî«∞ﬂx÷–µƒŒƒº˛ûÈ:" & vbCrLf & vbCrLf & List1.List(List1.ListIndex), vbInformation, "Info"
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Else
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
Exit Sub
End If
End Sub
Private Sub mnuWithInfo_Click()
On Error Resume Next
Option2.Value = True
lpOWM = EmptyFile
mnuZeroBytes.Checked = False
mnuWithInfo.Checked = True
On Error Resume Next
If Option2.Value = True Then
lpOWM = WithDestination
Else
lpOWM = EmptyFile
End If
End Sub
Private Sub mnuZeroBytes_Click()
On Error Resume Next
Option1.Value = True
lpOWM = EmptyFile
mnuZeroBytes.Checked = True
mnuWithInfo.Checked = False
On Error Resume Next
If Option2.Value = True Then
lpOWM = WithDestination
Else
lpOWM = EmptyFile
End If
End Sub
Private Sub Option1_Click()
On Error Resume Next
If Option2.Value = True Then
lpOWM = WithDestination
mnuWithInfo.Checked = True
mnuZeroBytes.Checked = False
Else
lpOWM = EmptyFile
mnuWithInfo.Checked = False
mnuZeroBytes.Checked = True
End If
End Sub
Private Sub Option2_Click()
On Error Resume Next
If Option2.Value = True Then
lpOWM = WithDestination
Else
lpOWM = EmptyFile
End If
On Error Resume Next
If Option2.Value = True Then
lpOWM = WithDestination
mnuWithInfo.Checked = True
mnuZeroBytes.Checked = False
Else
lpOWM = EmptyFile
mnuWithInfo.Checked = False
mnuZeroBytes.Checked = True
End If
End Sub
Private Sub Command1_GotFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Command1_LostFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Command2_GotFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Command2_LostFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Command3_GotFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Command3_LostFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Command4_GotFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Command4_LostFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Command5_GotFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Command5_LostFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Command6_GotFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Command6_LostFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Command7_GotFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Command7_LostFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Text1_GotFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Text1_LostFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Option1_GotFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Option1_LostFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Option2_GotFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
Private Sub Option2_LostFocus()
On Error Resume Next
If List1.ListIndex < 0 Then
Command6.Enabled = False
Command7.Enabled = False
ElseIf List1.ListIndex >= 0 Then
Command6.Enabled = True
Command7.Enabled = True
Else
Command6.Enabled = True
Command7.Enabled = True
End If
End Sub
