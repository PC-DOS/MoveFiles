VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "‘O÷√"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "»°œ˚(&C)"
      Height          =   420
      Left            =   5235
      TabIndex        =   13
      Top             =   2760
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "¥_∂®(&O)"
      Default         =   -1  'True
      Height          =   420
      Left            =   3675
      TabIndex        =   12
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Caption         =   "Œƒº˛äAÃÌº”—°œÓ"
      Height          =   975
      Left            =   75
      TabIndex        =   5
      Top             =   1710
      Width           =   6690
      Begin VB.CheckBox Check1 
         Caption         =   "œµÕ≥(&S)"
         Height          =   255
         Index           =   4
         Left            =   4095
         TabIndex        =   10
         Top             =   585
         Width           =   1005
      End
      Begin VB.CheckBox Check1 
         Caption         =   "“˛≤ÿ(&H)"
         Height          =   255
         Index           =   3
         Left            =   5415
         TabIndex        =   9
         Top             =   585
         Width           =   1005
      End
      Begin VB.CheckBox Check1 
         Caption         =   "¥Êµµ(&A)"
         Height          =   255
         Index           =   2
         Left            =   2775
         TabIndex        =   8
         Top             =   585
         Width           =   1005
      End
      Begin VB.CheckBox Check1 
         Caption         =   "±Í◊º(&N)"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   585
         Width           =   1005
      End
      Begin VB.CheckBox Check1 
         Caption         =   "÷ª∂¡(&R)"
         Height          =   255
         Index           =   1
         Left            =   1455
         TabIndex        =   6
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "èƒŒƒº˛äA÷–ÃÌº”Œƒº˛ïr£¨ÃÌº”æﬂ”–“‘œ¬ Ù–‘µƒŒƒº˛"
         Height          =   180
         Left            =   135
         TabIndex        =   11
         Top             =   270
         Width           =   3960
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "—}—uﬂ^≥Ã“‚Õ‚ΩK÷π≤Ÿ◊˜"
      Height          =   1575
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   6690
      Begin VB.OptionButton Option1 
         Caption         =   "Ã· æ”√ëÙﬂxìÒ≤Ÿ◊˜(&P)"
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   510
         Value           =   -1  'True
         Width           =   3975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "≤ªÃ· æ«“¿^¿m—u◊˜”∞◊”Œƒº˛(&U)"
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   847
         Width           =   3975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "≤ªÃ· æ«“≤ª—u◊˜”∞◊”Œƒº˛(&D)"
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   1185
         Width           =   3975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "’àﬂxìÒŒƒº˛“∆Ñ”ﬂ^≥Ã±ª“‚Õ‚ΩK÷π(»Á∞l…˙Âe’`ªÚ’ﬂƒ˙»°œ˚¡À≤Ÿ◊˜)ïràÃ––µƒ≤Ÿ◊˜"
         Height          =   180
         Left            =   135
         TabIndex        =   4
         Top             =   255
         Width           =   6120
      End
   End
End
Attribute VB_Name = "frmSettings"
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
On Error Resume Next
If Option1.Value = True Then
lpErrorOperation = PromptUser
End If
If Option2.Value = True Then
lpErrorOperation = Overwrite
End If
If Option3.Value = True Then
lpErrorOperation = Skip
End If
If Check1(2).Value = 1 Then
Form1.File1.Archive = True
Else
Form1.File1.Archive = False
End If
If Check1(0).Value = 1 Then
Form1.File1.Normal = True
Else
Form1.File1.Normal = False
End If
If Check1(1).Value = 1 Then
Form1.File1.ReadOnly = True
Else
Form1.File1.ReadOnly = False
End If
If Check1(3).Value = 1 Then
Form1.File1.Hidden = True
Else
Form1.File1.Hidden = False
End If
If Check1(4).Value = 1 Then
Form1.File1.System = True
Else
Form1.File1.System = False
End If
Unload Me
End Sub
Private Sub Command2_Click()
On Error Resume Next
If lpErrorOperation = Skip Then
Option3.Value = True
ElseIf lpErrorOperation = Overwrite Then
Option2.Value = True
Else
Option1.Value = True
End If
If Form1.File1.Archive = True Then
Check1(2).Value = 1
Else
Check1(2).Value = 0
End If
If Form1.File1.Normal = True Then
Check1(0).Value = 1
Else
Check1(0).Value = 0
End If
If Form1.File1.ReadOnly = True Then
Check1(1).Value = 1
Else
Check1(1).Value = 0
End If
If Form1.File1.Hidden = True Then
Check1(3).Value = 1
Else
Check1(3).Value = 0
End If
If Form1.File1.System = True Then
Check1(4).Value = 1
Else
Check1(4).Value = 0
End If
Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next
If lpErrorOperation = Skip Then
Option3.Value = True
ElseIf lpErrorOperation = Overwrite Then
Option2.Value = True
Else
Option1.Value = True
End If
If Form1.File1.Archive = True Then
Check1(2).Value = 1
Else
Check1(2).Value = 0
End If
If Form1.File1.Normal = True Then
Check1(0).Value = 1
Else
Check1(0).Value = 0
End If
If Form1.File1.ReadOnly = True Then
Check1(1).Value = 1
Else
Check1(1).Value = 0
End If
If Form1.File1.Hidden = True Then
Check1(3).Value = 1
Else
Check1(3).Value = 0
End If
If Form1.File1.System = True Then
Check1(4).Value = 1
Else
Check1(4).Value = 0
End If
End Sub
