Attribute VB_Name = "Déclarations2"
Option Explicit

Public Const INFINITE = &HFFFF
'STARTINFO constants
Public Const STARTF_USESHOWWINDOW = &H1
Public Enum enSW
    SW_HIDE = 0
    SW_NORMAL = 1
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
End Enum

Public Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Public Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Byte
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
 
Public Enum enPriority_Class
    NORMAL_PRIORITY_CLASS = &H20
    IDLE_PRIORITY_CLASS = &H40
    HIGH_PRIORITY_CLASS = &H80
End Enum
 
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" _
        (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
        lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As _
        SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal _
        dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory _
        As String, lpStartupInfo As STARTUPINFO, lpProcessInformation _
        As PROCESS_INFORMATION) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

