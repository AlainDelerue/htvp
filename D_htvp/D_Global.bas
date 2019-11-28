Attribute VB_Name = "D_Global"
Public Const Version = "V3.00"

' V1.03 File transfert with ID
' V1.04 Lock files
' V2.00 New GUI design
' V2.01 Recover and Remove WallPaper
' V2.02 Enhanced remote background management and Schedule feature available for functions following file upload
' V2.03 Timelag management for scheduler
' V2.04 Fix: Ends the program when the main windows is closed
' V2.05 Fix: Problem when there was several actions in the scheduler
' V2.06 Permanant WelcomeScreen
' V2.07 Mouse speed change, Get htvp Status
' V2.08 Scheduler fix
' V2.09 Better PC_Name entry management
' V2.10 Several dom side tools doesn't interfere with each other anymore with sub side version egal or above V.8.00
' V2.20 PC are sorted with alias
' V2.30 Disable/Enable task manager
' V2.40 Removed the use of the TMP directory which is sometimes erased by the system.
' V3.00 Some cleaning in the GUI, and PC names checked in upercase mode

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
                 


Public Remote_Version As String
Public PC_To_Delete As String
Public S_Prefix As String
Public L_Prefix As String
Public PC_Name As String
Public ClpBrd As String
Public Fic_ini As String
Public PC_PW() As String
Public nb_PC_PW As Integer
Public ScheduleActions As String
Public ScheduleOn As Boolean
Public Local_minus_Remote_Time As Date
Public TimeLagRecieved As Boolean

' File sending
'-------------
Public Send_FilePathAndName As String
Public Send_FileNamesent As String
Public Send_Options As Integer
Public Send_LastNumChunk As Integer
Public Send_Content As String
Public Send_length As Long
Public Send_TargetCmd As String
Public Send_ChunkSize As Double
Public Send_DispProgress As String
Public Send_Timeout_Chunk As Integer
Public Send_ID As String

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

Public Type Action_Type
    Command As String
    t_Param_lbl() As String
    nb_t_Param_lbl As Integer
    Help As String
End Type
Public t_Actions_Account() As Action_Type
Public nb_t_Actions_Account As Integer
Public t_Actions_TV() As Action_Type
Public nb_t_Actions_TV As Integer
Public t_Actions_App() As Action_Type
Public nb_t_Actions_App As Integer
Public t_Actions_Mean() As Action_Type
Public nb_t_Actions_Mean As Integer
Public t_Actions_Miscellaneous() As Action_Type
Public nb_t_Actions_Miscellaneous As Integer

Public GDIPlusOK As Boolean
Public BackGroundTmpFile As String
Public BackGroundRemoteFileName As String

Public Aff_Preview_Background_t As Long
Public Aff_Preview_Background_l As Long
Public Aff_Preview_Background_h As Long
Public Aff_Preview_Background_w As Long

Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Public This_PC_Name As String

Public Function GetComputerName() As String
Dim sResult As String * 255
    GetComputerNameA sResult, 255
    GetComputerName = Left(sResult, InStr(sResult, Chr(0)) - 1)
End Function

