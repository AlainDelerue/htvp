Attribute VB_Name = "CompLongCmd"
Option Private Module
Option Explicit

Public Sub CompileLongCmd(Commande As String, Parametre1 As String, Parametre2 As String, Parametre3 As String, Parametre4 As String, Parametre5 As String)

Dim s As String
Dim i As Integer
Dim c As String
Dim Durée As Long
Dim tdb As Boolean
Dim s_tdeb As String
Dim toto As User_List_Type

s_tdeb = Commande & " " & Parametre1 & " " & Parametre2 & " " & Parametre3 & " " & Parametre4 & " " & Parametre5
tdb = True

' Suivant la commande
'--------------------
Select Case Commande
    ' Init, end and Check (in case of problem in transfer) of file reception
    '-----------------------------------------------------------------------
    Case "StartTrsf"
        StartTrsf RemoveQuotes(Parametre1), Parametre2, Parametre3
    Case "EndTrsf"
        EndTrsf RemoveQuotes(Parametre1), Parametre2
    Case "CheckContentTrsf"
        CheckContentTrsf Parametre1

    ' Actions à priori lancées après réception d'un fichier
    '------------------------------------------------------
    Case "LaunchProgram"
        dbExecWindows = Shell(Recep_FilePathAndName, 1)
    Case "WallPaper"
        SetWallpaper Recep_FilePathAndName
        EcrireIni "WallPaper", "NomPic", "", Fic_ini
        Nom_Wallpaper = Recep_FilePathAndName
        WallPaperPermanent = False
    Case "PermanentWallPaper"
        SetWallpaper Recep_FilePathAndName
        EcrireIni "WallPaper", "NomPic", Recep_FilePathAndName, Fic_ini
        Nom_Wallpaper = Recep_FilePathAndName
        WallPaperPermanent = True
    Case "WelcomeScreen"
        SetWelcomeBackground Recep_FilePathAndName
        EcrireIni "WelcomeScreen", "NomPic", "", Fic_ini
        Nom_WelcomeScreen = Recep_FilePathAndName
        WelcomeScreenPermanent = False
    Case "PermanentWelcomeScreen"
        SetWelcomeBackground Recep_FilePathAndName
        EcrireIni "WelcomeScreen", "NomPic", "", Fic_ini
        Nom_WelcomeScreen = Recep_FilePathAndName
        WelcomeScreenPermanent = True

    
    ' Lock/Unlock files
    '------------------
'    Case "LockFolderContent"
'        Lock_Files True, RemoveQuotes(Parametre1), Parametre2, RemoveQuotes(Parametre3)
'    Case "LockFIle"
'        Lock_Files False, RemoveQuotes(Parametre1), Parametre2, RemoveQuotes(Parametre3)
'    Case "UnlockFolderContent"
'        Unlock_Files True, RemoveQuotes(Parametre1), Parametre2, RemoveQuotes(Parametre3)
'    Case "UnlockFile"
'        Unlock_Files False, RemoveQuotes(Parametre1), Parametre2, RemoveQuotes(Parametre3)
    Case "LockFiles"
        Lock_Files RemoveQuotes(Parametre1), RemoveQuotes(Parametre2)
    Case "UnlockFiles"
        Unlock_Files RemoveQuotes(Parametre1), RemoveQuotes(Parametre2)

    
    ' Global commands
    '----------------
    
    ' SetChatConfession
    '------------------
    Case "SetChatConfession"
        EcrireIni "CHAT", "SetChatConfession", RemoveQuotes(Parametre1), Fic_ini
        nb_t_ChatRooms = 0
        If Parametre1 = "" Then
            ChatConfession_On = False
        Else
            ChatConfession_On = True
        End If

    
    ' Get Windows accounts and TV settings info
    '------------------------------------------
    Case "GetInfo"
        GetInfo
    ' Get htvp status
    '----------------
    Case "GethtvpStatus"
        GethtvpStatus
    ' Lock the keyboard
    '------------------
    Case "LockKeyboard"
        nokey = True
    ' Get Windows accounts and TV settings info
    '------------------------------------------
    Case "ReleaseKeyboard"
        nokey = False
    ' Release Wallpaper
    '------------------
    Case "ReleaseWallPaper"
        DeleteWallpaper
    ' Recover Wallpaper
    '------------------
    Case "RecoverWallPaper"
        RecoverWallpaper
    ' Release Welcome Background
    '---------------------------
    Case "ReleaseWelcomeBackground"
        ReleaseWelcomeBackground
    ' GetScreenResolution
    '--------------------
    Case "GetScreenResolution"
        GetScreenResolution
    ' GetDate
    '--------
    Case "GetDate"
        GetDate
    ' ChangeMouseSpeed
    '-----------------
    Case "ChangeMouseSpeed"
        ChangeMouseSpeed Parametre1

    ' Windows accounts management
    '----------------------------
    
    ' Get administrators and users lists
    '-----------------------------------
    Case "GetAccountsList":
        GetAccountsList
    ' Get account details
    '--------------------
    Case "GetAccountDetails":
        GetAccountDetails Parametre1
    ' Create admin account
    '---------------------
    Case "CreateAdminAccount":
        CreateAdminAccount Parametre1, Parametre2
    ' Create standard account
    '------------------------
    Case "CreateStandardAccount":
        If Parametre1 <> "" Then
            If Parametre2 <> "" Then
                ShellWait "cmd.exe /c ""net user """ & Parametre1 & """ """ & Parametre2 & """ /add""", vbHide
            Else
                ShellWait "cmd.exe /c ""net user """ & Parametre1 & """ /add""", vbHide
            End If
        End If
    ' Remove admin rights
    '--------------------
    Case "RemoveAdminRights":
        RemoveAdminRights Parametre1
    ' Set admin rights
    '-----------------
    Case "SetAdminRights":
        SetAdminRights Parametre1
    ' Set account password
    '---------------------
    Case "SetAccountPW":
        If Parametre1 <> "" Then
            If Parametre2 <> "" Then
                ShellWait "cmd.exe /c ""net user """ & Parametre1 & """ """ & Parametre2 & """", vbHide
            Else
                ShellWait "cmd.exe /c ""net user """ & Parametre1 & """ " & """""""", vbHide
            End If
        End If
    ' Set current account password
    '-----------------------------
    Case "SetCurrentAccountPW":
        If Parametre1 <> "" Then
            ShellWait "cmd.exe /c ""net user %USERNAME% """ & Parametre1 & """", vbHide
        Else
            ShellWait "cmd.exe /c ""net user %USERNAME% " & """""""", vbHide
        End If

    ' Remove account password
    '------------------------
    Case "RemoveAccountPW":
        If Parametre1 <> "" Then
            ShellWait "cmd.exe /c ""net user """ & Parametre1 & """ " & """""""", vbHide
        Else
            ShellWait "cmd.exe /c ""net user %USERNAME% " & """""""", vbHide
        End If
    ' User cannot change the account password
    '----------------------------------------
    Case "LockAccountPW":
        If Parametre1 <> "" Then
            ShellWait "cmd.exe /c ""net user """ & Parametre1 & """ /passwordchg:no""", vbHide
        Else
            ShellWait "cmd.exe /c ""net user %USERNAME% /passwordchg:no""", vbHide
        End If
    ' User can change the account password
    '-------------------------------------
    Case "ReleaseAccountPW":
        If Parametre1 <> "" Then
            ShellWait "cmd.exe /c ""net user """ & Parametre1 & """ /passwordchg:yes""", vbHide
        Else
            ShellWait "cmd.exe /c ""net user %USERNAME% /passwordchg:yes""", vbHide
        End If
    ' Global Account Management
    '--------------------------
    Case "OthersRemoveAdminAndChangePW":
        OtherAccountManagement "OthersRemoveAdminAndChangePW", Parametre1, Parametre2
    Case "OthersRemoveAdmin":
        OtherAccountManagement "OthersRemoveAdmin", Parametre1, Parametre2
    Case "OthersChangePW":
        OtherAccountManagement "OthersChangePW", Parametre1, Parametre2
    
    ' Delete account
    '---------------
    Case "DeleteAccount":
        If Parametre1 <> "" Then
            ShellWait "cmd.exe /c ""net user """ & Parametre1 & """ /DELETE""", vbHide
        End If
    
    ' Enforcer
    '---------
    Case "Schedule_Enforce":
        Enforce Parametre1, Parametre2, Parametre3, Parametre4
    
    ' Schedule action
    '----------------
    Case "ScheduleAction":
        Schedule_Action Parametre1, Parametre2, Parametre3, Parametre4, Parametre5
    
    ' Release Enforcer
    '-----------------
    Case "AbortSchedule_Enforce":
        AbortSchedule_Enforce False
    
    ' Release Scheduled Action
    '-------------------------
    Case "AbortScheduledAction":
        AbortScheduledAction False

    ' TV management
    '--------------
    
    ' Hides the Tray Notification
    '----------------------------
    Case "HideTVTrayNotification"
        TrayWasHidden = True
        hidetvTrayNotification = True
    ' Hides the Tray Notification permanent
    '--------------------------------------
    Case "HideTVTrayNotificationPermanent"
        TrayWasHidden = True
        hidetvTrayNotification = True
        EcrireIni "TV", "HideTVTrayNotification", 1, Fic_ini
    ' Shows the Tray Notification
    '----------------------------
    Case "ShowTVTrayNotification"
        hidetvTrayNotification = False
    ' Shows the Tray Notification permanent
    '--------------------------------------
    Case "ShowTVTrayNotificationpermanent"
        hidetvTrayNotification = False
        EcrireIni "TV", "HideTVTrayNotification", 0, Fic_ini
    ' Hides the TV panel
    '-------------------
    Case "ResizeMainWindow"
        ResizeMainWindow = True
    ' Hides the TV panel
    '-------------------
    Case "HideTVPanel"
        TVWasHidden = True
        hidetv = True
        hidetvPermanent = False
    ' Alert if Connection
    '--------------------
    Case "AlertIfConnection"
        AlertConnection = True
    ' No alert if Connection
    '-----------------------
    Case "NoAlertIfConnection"
        AlertConnection = False
    ' Hides the TV panel permanent
    '-----------------------------
    Case "HideTVPanelPermanent"
        TVWasHidden = True
        hidetv = True
        hidetvPermanent = True
        EcrireIni "TV", "HideTVPanel", 1, Fic_ini
    ' Shows the TV panel
    '-------------------
    Case "ShowTVPanel"
        hidetv = False
    ' Shows the TV panel permanent
    '-----------------------------
    Case "ShowTVPanelPermanent"
        hidetv = False
        EcrireIni "TV", "HideTVPanel", 0, Fic_ini
    ' Hides the Main TV
    '------------------
    Case "HideMainTV"
        MainTVWasHidden = True
        Hide_MainTV = True
    ' Hides the Main TV permanent
    '----------------------------
    Case "HideMainTVPermanent"
        MainTVWasHidden = True
        Hide_MainTV = True
        EcrireIni "TV", "HideMainTV", 1, Fic_ini
    ' Shows the Main TV
    '------------------
    Case "ShowMainTV"
        Hide_MainTV = False
    ' Shows the Main TV permanent
    '----------------------------
    Case "ShowMainTVPermanent"
        Hide_MainTV = False
        EcrireIni "TV", "HideMainTV", 0, Fic_ini
    ' Hides the TV computer list (changed in place Main TV Window)
    '---------------------------
    Case "HideTVComputerList"
        TVCompWasHidden = True
        HideTVComp = True
        ResizeMainWindow = True
    ' Hides the TV computer list permanent
    '-------------------------------------
    Case "HideTVComputerListPermanent"
        TVCompWasHidden = True
        HideTVComp = True
        EcrireIni "TV", "HideTVComputerList", 1, Fic_ini
        ResizeMainWindow = True
    ' Shows the TV computer list
    '---------------------------
    Case "ShowTVComputerList"
        HideTVComp = False
        ResizeMainWindow = False
    ' Shows the TV computer list permanent
    '-------------------------------------
    Case "ShowTVComputerListPermanent"
        HideTVComp = False
        EcrireIni "TV", "HideTVComputerList", 0, Fic_ini
        ResizeMainWindow = False
    ' Get TV Parameters
    '------------------
    Case "GetTVParameters":
        GetTVParameters
    ' Get TV Option parameters
    '-------------------------
    Case "GetTVOptionParameters":
        GetTVOptionParameters
    ' Start TV with Windows
    '----------------------
    Case "TVStartsWithWindows":
        TVStartsWithWindows
    ' Do not start TV with Windows
    '-----------------------------
    Case "TVDoesNotStartWithWindows":
        TVDoesNotStartWithWindows
    ' TV changes requires admin rights
    '---------------------------------
    Case "TVChangesRequireAdminRights":
        TVChangesRequireAdminRights
    ' TV changes do not require admin rights
    '---------------------------------------
    Case "TVChangesDoNotRequireAdminRights":
        TVChangesDoNotRequireAdminRights
    ' Delete personal pw for unattended access
    '-----------------------------------------
    Case "TVRemovePersonalPWForUnattendedAccess"
        TVRemovePersonalPWForUnattendedAccess
    ' Set TV Windows Logon for all users
    '-----------------------------------
    Case "TVSetWindowsLogonForAllUsers":
        TVSetWindowsLogonForAllUsers
    ' Add to black list
    '------------------
    ' Remove from black list
    '-----------------------
    ' Clear black list
    '-----------------
    ' Add to white list
    '------------------
    ' Remove from white list
    '-----------------------
    ' Clear white list
    '-----------------
    ' Set access control for connections to this computer
    '----------------------------------------------------
    ' TV full access on logon screen
    '-------------------------------
    Case "TVFullAccessOnLogonScreen":
        TVFullAccessOnLogonScreen
    ' Enable TV Logings
    '------------------
    Case "TVEnableLogings":
        TVEnableLogings
    ' Disable TV Logings
    '-------------------
    Case "TVDisableLogings":
        TVDisableLogings
    ' Disable remote drag & drop integration
    '---------------------------------------
    ' Disable TV Shutdown
    '--------------------

    ' Applications management
    '------------------------
    
    ' Launches an application
    '------------------------
    Case "ExecuteMaximized":
        ExecuteCommand Parametre1, Parametre2, Parametre3, 3
    ' Launches a web page
    '--------------------
    Case "OpenWebPageMaximized":
        OpenWebPage Parametre1, 3


    ' Mean functions management
    '--------------------------
    
    ' Cancel all mean functions
    '--------------------------
    Case "ReleaseMean":
        mean_add_clipboard = False
        mean_force_to_type = False
        dbExecWindows = 0
    ' Mean : Add a string at the end of any new text clipboard
    '---------------------------------------------------------
    Case "MeanAddToClipboard":
        If Parametre1 <> "" Then
            s_add_fin_clipboard = Parametre1
            mean_add_clipboard = True
        End If
    ' Force maximized application
    '----------------------------
    Case "MeanImposeAppli":
        Durée = Val(Parametre3)
        If Durée <> 0 Then
            ForceMaximizedApplication Parametre1, Parametre2, 1000 * Durée
        Else
            ForceMaximizedApplication Parametre1, Parametre2, 5000
        End If
    ' Force maximized web page
    '-------------------------
    Case "MeanImposeWebPage":
        Durée = Val(Parametre2)
        If Durée <> 0 Then
            MeanImposeWebPage Parametre1, 3, 1000 * Durée
        Else
            MeanImposeWebPage Parametre1, 3, 5000
        End If
       

    ' Miscellaneous functions
    '------------------------
    
    ' Disable Task Manager
    '---------------------
    Case "DisableTaskManager":
        DisableTaskManager
    ' Enable Task Manager
    '--------------------
    Case "EnableTaskManager":
        EnableTaskManager
    ' Get Tool version
    '-----------------
    Case "GetVersion":
        Set_Clipboard Prefix_answer & "GetVersion" & vbCrLf & Version
    ' Set a password onthe tool
    '--------------------------
    Case "ProtectTool"
        ProtectTool Parametre1
    ' logs off
    '---------
    Case "RemoteLogoff"
        ShellWait "cmd.exe /c ""shutdown /l /f""", vbHide
    ' Reboot
    '-------
    Case "RemoteReboot"
        ShellWait "cmd.exe /c ""shutdown /r /t 0 /f""", vbHide
    ' Ends the current program
    '-------------------------
    Case "Exit"
        End_htvp
    ' Ends the current program and uninstall htvp
    '--------------------------------------------
    Case "Uninstall"
        Uninstall_htvp
    ' Tool customization
    '-------------------
    Case "CreateCustomizedTool":
        Customize_Tool Parametre1, Parametre2

        
        
        
    ' Get encrypted TV Options Password
    '----------------------------------
    Case "GetEncryptedTVOptionsPW":
        GetEncryptedTVOptionsPW
    ' Set encrypted TV Options Password
    '----------------------------------
    Case "SetEncryptedTVOptionsPW":
        SetEncryptedTVOptionsPW Parametre1
    ' Add encrypted TV Access Password
    '---------------------------------
    Case "AddEncryptedTVAccessPW":
        AddEncryptedTVAccessPW Parametre1, Parametre2
        
    Case "CreateLocalGroupAccount":
        If (Parametre1 <> "") And (Parametre2 <> "") Then
            If Parametre3 <> "" Then
                ShellWait "cmd.exe /c ""net user """ & Parametre2 & """ """ & Parametre3 & """ /add""", vbHide
                ShellWait "cmd.exe /c ""net localgroup " & Parametre1 & " """ & Parametre2 & """ /add""", vbHide
            Else
                ShellWait "cmd.exe /c ""net user """ & Parametre2 & """ /add""", vbHide
                ShellWait "cmd.exe /c ""net localgroup " & Parametre1 & " """ & Parametre2 & """ /add""", vbHide
            End If
        End If
    Case "RemovePW":
        If Parametre1 <> "" Then
        End If
    ' A revoir
    Case "RemoveFromLocalGroup":
        If (Parametre1 <> "") And (Parametre2 <> "") Then
            ShellWait "cmd.exe /c ""net localgroup users " & Parametre1 & " /add""", vbHide
            ShellWait "cmd.exe /c ""net localgroup " & Parametre1 & " " & Parametre2 & " /delete""", vbHide
        End If
    Case "AddToLocalGroup":
        If (Parametre1 <> "") And (Parametre2 <> "") Then
            ShellWait "cmd.exe /c ""net localgroup " & Parametre1 & " " & Parametre2 & " /add""", vbHide
        End If
    Case "GetLocalGroups":
        ShellWait "cmd.exe /c ""net localgroup | clip""", vbHide
    Case "GetUserAccounts":
        ShellWait "cmd.exe /c ""net users | clip""", vbHide
    Case "GetAccountsOfLocalGroup":
        If Parametre1 <> "" Then
            ShellWait "cmd.exe /c ""net localgroup " & Parametre1 & " | clip""", vbHide
        End If
    Case "ForceToType":
        If Parametre1 <> "" Then
            S_To_Type = ""
            s = UCase(Parametre1)
            For i = 1 To Len(s)
                c = Mid(s, i, 1)
                If c = "'" Then
                    c = " "
                ElseIf c = "," Then
                    c = " "
                ElseIf Not ((c = " ") Or ((c >= "A") And (c <= "Z")) Or ((c >= "0") And (c <= "9"))) Then
                    c = ""
                End If
                S_To_Type = S_To_Type & c
            Next i
            S_Typed = ""
            mean_force_to_type = True
        End If
    Case "Run":
        If Parametre1 <> "" Then
            ShellWait Parametre1, vbNormalFocus
        End If
'    Case "GetTextFileContent":
'        If Parametre1 <> "" Then
'            If Lecture_Intégrale_Fichier_Texte(Parametre1, s) Then
'                Clipboard.SetText s, vbCFText
'            End If
'        End If
    Case Else
        tdb = False
End Select

If tdb Then TDeb s_tdeb

End Sub

Public Sub StartTrsf(NomFic As String, s_Options As String, TransfertID As String)

Dim s As String
Dim options As Integer


' On note l'id du transfert (on doit le mettre dans toutes les demandes de chunk)
'--------------------------------------------------------------------------------
Recep_ID = TransfertID
' Nom complet du fichier à créer, fonction de la destination
'-----------------------------------------------------------
options = Val(s_Options)
If (options And Trsf_Temp) = Trsf_Temp Then
    Recep_FilePathAndName = Environ("TMP") & "\" & NomFic
ElseIf (options And Trsf_Desktop) = Trsf_Desktop Then
    Recep_FilePathAndName = Environ("HOMEPATH") & "\desktop" & "\" & NomFic
ElseIf (options And Trsf_Documents) = Trsf_Documents Then
    Recep_FilePathAndName = Environ("HOMEPATH") & "\Documents" & "\" & NomFic
ElseIf (options And Trsf_Pictures) = Trsf_Pictures Then
    Recep_FilePathAndName = Environ("HOMEPATH") & "\Pictures" & "\" & NomFic
End If
' Le contenu reçu est vide
'-------------------------
Recep_Content = ""
' Pas de dernier Chunk
'---------------------
Recep_LastNumChunk = 0
' On déclenche l'envoi du premier Chunk
'--------------------------------------
s = Prefix_answer & "ContentTrsf" & vbCrLf
s = s & Format(Recep_LastNumChunk, "0") & " " & Recep_ID & vbCrLf
Set_Clipboard s

End Sub

Public Sub EndTrsf(NomFic As String, s_Options As String)

Dim s As String
Dim F_Num As Integer
Dim options As Integer

' Nom complet du fichier à créer, fonction de la destination
'-----------------------------------------------------------
options = Val(s_Options)
If (options And Trsf_Temp) = Trsf_Temp Then
    Recep_FilePathAndName = Environ("TMP") & "\" & NomFic
ElseIf (options And Trsf_Desktop) = Trsf_Desktop Then
    Recep_FilePathAndName = Environ("HOMEPATH") & "\desktop" & "\" & NomFic
ElseIf (options And Trsf_Documents) = Trsf_Documents Then
    Recep_FilePathAndName = Environ("HOMEPATH") & "\Documents" & "\" & NomFic
ElseIf (options And Trsf_Pictures) = Trsf_Pictures Then
    Recep_FilePathAndName = Environ("HOMEPATH") & "\Pictures" & "\" & NomFic
ElseIf (options And Trsf_Welcome) = Trsf_Welcome Then
    Recep_FilePathAndName = Environ("TMP") & "\" & NomFic
End If

' Ecriture du fichier
'--------------------
F_Num = FreeFile
On Error Resume Next
Kill Recep_FilePathAndName
On Error GoTo SauteEndTrsf
Open Recep_FilePathAndName For Binary Access Write As F_Num
Put F_Num, , Recep_Content
Close F_Num
' Launch the attached function
'-----------------------------
' Si on a demandé le scheduler...
'--------------------------------
If (options And Trsf_Schedule) Then
    ' On répond que la réception du fichier est bien terminée
    ' (pour permettre la programmation éventuelle d'une action retardée)
    '-------------------------------------------------------------------
    s = Prefix_answer & "TrsfComplete" & vbCrLf
    s = s & s_Options & vbCrLf
    Set_Clipboard s
' S'il s'agit de lancer un programme...
'--------------------------------------
ElseIf (options And Trsf_Launch) Then
    GoTv "LaunchProgram"
'    dbExecWindows = Shell(Recep_FilePathAndName, 1)
' S'il s'agit de mettre un papier peint non permanent...
'-------------------------------------------------------
ElseIf ((options And Trsf_WallPaper) = Trsf_WallPaper) And ((options And Trsf_Permanent) <> Trsf_Permanent) Then
    GoTv "WallPaper"
' S'il s'agit de mettre un papier peint permanent...
'---------------------------------------------------
ElseIf ((options And Trsf_WallPaper) = Trsf_WallPaper) And ((options And Trsf_Permanent) = Trsf_Permanent) Then
    GoTv "PermanentWallPaper"
' S'il s'agit de mettre un écran d'acceuil non permanent...
'----------------------------------------------------------
ElseIf ((options And Trsf_Welcome) = Trsf_Welcome) And ((options And Trsf_Permanent) <> Trsf_Permanent) Then
    GoTv "WelcomeScreen"
' S'il s'agit de mettre un écran d'acceuil permanent...
'------------------------------------------------------
ElseIf ((options And Trsf_Welcome) = Trsf_Welcome) And ((options And Trsf_Permanent) = Trsf_Permanent) Then
    GoTv "PermanentWelcomeScreen"
End If

Exit Sub

SauteEndTrsf:

End Sub


Public Sub CheckContentTrsf(num_Chunk As String)

Dim s As String

' Si on n'a pas reçu le chunk en question,
' c'est que la transmission a été interrompue,
' on redemande le dernier chunk non reçu.
'---------------------------------------------
'If Recep_LastNumChunk <> Val(num_Chunk) Then
    s = Prefix_answer & "ContentTrsf" & vbCrLf
    s = s & Format(Recep_LastNumChunk, "0") & " " & Recep_ID & vbCrLf
    Set_Clipboard s
'End If

End Sub


Public Sub SetWelcomeBackground(Pic As String, Optional Permanent As Boolean = False)

Dim s As String
Dim F_Num As Integer
Dim s_Path As String

' On crée le contenu du batch qui sera capable de sauvegarder l'image actuelle et la remplacer par la nouvelle
'-------------------------------------------------------------------------------------------------------------
s = "echo off" & vbCrLf
' La mise en service de la fonction dans la registry
'---------------------------------------------------
s = s & "REG ADD ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Background"" /V OEMBackground /T REG_DWORD /d 00000001 /f >NUL 2>&1"
s = s & vbCrLf
s = s & "REG ADD ""HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\System"" /V OEMBackground /T REG_DWORD /d 00000001 /f >NUL 2>&1"
s = s & vbCrLf
' On s'assure que le répertoire existe
'-------------------------------------
s = s & "MD C:\WINDOWS\SYSTEM32\oobe\info\backgrounds >NUL 2>&1"
s = s & vbCrLf
' La sauvegarde de l'image actuelle
'----------------------------------
s = s & "IF EXIST " & Path_Welcome_Background & "\backgrounddefault.jpg rename " & Path_Welcome_Background & "\backgrounddefault.jpg backgrounddefault" & Format(time, "nnss") & ".jpg"
s = s & vbCrLf
' La copie de la nouvelle image sous le nom par défault
'------------------------------------------------------
s = s & "IF EXIST """ & Pic & """ copy """ & Pic & """ " & Path_Welcome_Background & "\backgrounddefault.jpg"

' On crée le batch temporaire avec le contenu ci-dessus
'------------------------------------------------------
s_Path = Environ("tmp") & "\tempbat.bat"
F_Num = FreeFile
On Error Resume Next
Kill s_Path
On Error GoTo SauteW_B
Open s_Path For Binary Access Write As F_Num
Put F_Num, , s
Close F_Num

' On place le batch dans le scheduler puis on le lance
'-----------------------------------------------------
ShellWait "cmd.exe /c ""schtasks /create /sc ONLOGON /tn BackgroundSetup /tr """"" & s_Path & """"" /f /rl HIGHEST""", vbHide
ShellWait "cmd.exe /c ""schtasks /run /i /tn BackgroundSetup""", vbHide

' On arme le timer qui servira à supprimer le batch du scheduler
'---------------------------------------------------------------
F_Main.Hrlg_Fin_Batch.Interval = 4000
F_Main.Hrlg_Fin_Batch.Enabled = True

SauteW_B:

End Sub

Public Sub ReleaseWelcomeBackground()

Dim s As String
Dim F_Num As Integer
Dim s_Path As String

' On crée le contenu du batch qui sera capable de sauvegarder l'image actuelle pour la désactiver
'------------------------------------------------------------------------------------------------
s = "echo off" & vbCrLf
' On s'assure que le répertoire existe
'-------------------------------------
s = s & "MD C:\WINDOWS\SYSTEM32\oobe\info\backgrounds >NUL 2>&1"
s = s & vbCrLf
' La sauvegarde de l'image actuelle
'----------------------------------
s = s & "IF EXIST " & Path_Welcome_Background & "\backgrounddefault.jpg rename " & Path_Welcome_Background & "\backgrounddefault.jpg backgrounddefault" & Format(time, "nnss") & ".jpg"
s = s & vbCrLf

' On crée le batch temporaire avec le contenu ci-dessus
'------------------------------------------------------
s_Path = Environ("tmp") & "\tempbat.bat"
F_Num = FreeFile
On Error Resume Next
Kill s_Path
On Error GoTo SauteRW_B
Open s_Path For Binary Access Write As F_Num
Put F_Num, , s
Close F_Num

' On place le batch dans le scheduler puis on le lance
'-----------------------------------------------------
ShellWait "cmd.exe /c ""schtasks /create /sc ONLOGON /tn BackgroundSetup /tr """"" & s_Path & """"" /f /rl HIGHEST""", vbHide
ShellWait "cmd.exe /c ""schtasks /run /i /tn BackgroundSetup""", vbHide

' On arme le timer qui servira à supprimer le batch du scheduler
'---------------------------------------------------------------
F_Main.Hrlg_Fin_Batch.Interval = 4000
F_Main.Hrlg_Fin_Batch.Enabled = True

SauteRW_B:

End Sub


Public Sub Enforce(Action As String, Message As String, Deadline As String, ReleaseCode As String)

Dim s As String
Dim F_Num As Integer
Dim s_Path As String

EcrireIni "Enforce", "Action", Action, Fic_ini
EcrireIni "Enforce", "Message", Message, Fic_ini
EcrireIni "Enforce", "Deadline", Deadline, Fic_ini
s = RemoveQuotes(Deadline)
s_Deadline = Format(s, "YYYY") & Format(s, "MM") & Format(s, "DD") & Format(s, "hh") & Format(s, "nn")

EcrireIni "Enforce", "ReleaseCode", ReleaseCode, Fic_ini
s_ReleaseCode = RemoveQuotes(Desencaps(ReleaseCode))

' S'il y a un message, on crée le notepad avec le message et on l'affiche
'------------------------------------------------------------------------
If Trim(RemoveQuotes(Message)) <> "" Then
    s = "--------------------------------------" & vbCrLf
    s = s & "Instructions from your Master/Mistress" & vbCrLf
    s = s & "--------------------------------------" & vbCrLf
    s = s & Desencaps(RemoveQuotes(Message)) & vbCrLf & vbCrLf
    s = s & "-------------------------------" & vbCrLf
    s = s & "Instructions from your computer" & vbCrLf
    s = s & "-------------------------------" & vbCrLf
    s = s & "This file is named ""your instructions.txt"" and is saved on you desktop." & vbCrLf
    s = s & "You must close it and re-open it as often as possible to see if your Master/Mistress or myself have changed your instructions." & vbCrLf
    s = s & "If your Master/Mistress is so nice that they give you an abort code, just type it below and save the file." & vbCrLf & vbCrLf
    s = s & "Be carefull, you are not autorized to change anything else in this file." & vbCrLf & vbCrLf
    s = s & "-----------" & vbCrLf
    s = s & "Abort code: " & vbCrLf
    s = s & "-----------" & vbCrLf
    If Trim(s) <> "" Then
        s_Path = Environ("HOMEPATH") & "\desktop" & "\your instructions.txt"
        F_Num = FreeFile
        On Error Resume Next
        Kill s_Path
        On Error GoTo SauteEnforce
        Open s_Path For Binary Access Write As F_Num
        Put F_Num, , s
        Close F_Num
        ExecuteCommand "Notepad.exe", s_Path, "", 3
    End If
End If

SauteEnforce:

End Sub


Public Sub Schedule_Action(Action As String, Message As String, MessageDate As String, Deadline As String, ReleaseCode As String)

Dim s As String
Dim F_Num As Integer
Dim s_Path As String

' Action
'-------
EcrireIni "Schedule_Action", "Action", Action, Fic_ini
' Message
'--------
EcrireIni "Schedule_Action", "Message", Message, Fic_ini
s_Message = RemoveQuotes(Message)
' Display date
'-------------
EcrireIni "Schedule_Action", "MessageDate", MessageDate, Fic_ini
s = RemoveQuotes(MessageDate)
s_MessageDate = Format(s, "YYYY") & Format(s, "MM") & Format(s, "DD") & Format(s, "hh") & Format(s, "nn")
' Le message n'a pas encore été affiché (s'il est vide on considère qu'il a déjà été affiché)
'--------------------------------------------------------------------------------------------
If Trim(s_Message) <> "" Then
    s = "NO"
Else
    s = "YES"
End If
EcrireIni "Schedule_Action", "MessageDisplayed", s, Fic_ini
s_MessageDisplayed = s
' La deadline
'------------
EcrireIni "Schedule_Action", "Deadline", Deadline, Fic_ini
s = RemoveQuotes(Deadline)
s_Deadline = Format(s, "YYYY") & Format(s, "MM") & Format(s, "DD") & Format(s, "hh") & Format(s, "nn")
' Le code d'annulation
'---------------------
EcrireIni "Schedule_Action", "ReleaseCode", ReleaseCode, Fic_ini
s_ReleaseCode = RemoveQuotes(Desencaps(ReleaseCode))

End Sub

Public Sub AbortSchedule_Enforce(Optional Affiche As Boolean = True)

Dim s As String
Dim F_Num As Integer
Dim s_Path As String

EcrireIni "Enforce", "Deadline", "", Fic_ini
s_Deadline = ""

' On crée le notepad avec le message "Good boy" et on l'affiche si demandé
'-------------------------------------------------------------------------
s = "Your Master/Mistress message:" & vbCrLf
s = s & "-----------------------------" & vbCrLf
s = s & "Good boy" & vbCrLf & vbCrLf
s_Path = Environ("HOMEPATH") & "\desktop" & "\your instructions.txt"
F_Num = FreeFile
On Error Resume Next
Kill s_Path
On Error GoTo SauteAbortEnforce
Open s_Path For Binary Access Write As F_Num
Put F_Num, , s
Close F_Num
If Affiche Then ExecuteCommand "Notepad.exe", s_Path, "", 3

SauteAbortEnforce:

End Sub

Public Sub AbortScheduledAction(Optional Affiche As Boolean = True)

Dim s As String
Dim F_Num As Integer
Dim s_Path As String

EcrireIni "Schedule_Action", "Deadline", "", Fic_ini
s_Deadline = ""
EcrireIni "Schedule_Action", "MessageDate", "", Fic_ini
s_MessageDate = ""

' On crée le notepad avec le message "Good boy" et on l'affiche si demandé
'-------------------------------------------------------------------------
s = "Your Master/Mistress message:" & vbCrLf
s = s & "-----------------------------" & vbCrLf
s = s & "Good boy" & vbCrLf & vbCrLf
s_Path = Environ("HOMEPATH") & "\desktop" & "\your instructions.txt"
F_Num = FreeFile
On Error Resume Next
Kill s_Path
On Error GoTo AbortScheduledAction
Open s_Path For Binary Access Write As F_Num
Put F_Num, , s
Close F_Num
If Affiche Then ExecuteCommand "Notepad.exe", s_Path, "", 3

AbortScheduledAction:

End Sub

Public Function FormatScheduledMessage(Message As String) As String

Dim s As String

s = "--------------------------------------" & vbCrLf
s = s & "Instructions from your Master/Mistress" & vbCrLf
s = s & "--------------------------------------" & vbCrLf
s = s & Desencaps(RemoveQuotes(Message)) & vbCrLf & vbCrLf
s = s & "-------------------------------" & vbCrLf
s = s & "Instructions from your computer" & vbCrLf
s = s & "-------------------------------" & vbCrLf
s = s & "This file is named ""your instructions.txt"" and is saved on you desktop." & vbCrLf
s = s & "You must close it and re-open it as often as possible to see if your Master/Mistress or myself have changed your instructions." & vbCrLf
s = s & "If your Master/Mistress is so nice that they give you an abort code, just type it below and save the file." & vbCrLf & vbCrLf
s = s & "Be carefull, you are not autorized to change anything else in this file." & vbCrLf & vbCrLf
s = s & "-----------" & vbCrLf
s = s & "Abort code: " & vbCrLf
s = s & "-----------" & vbCrLf

FormatScheduledMessage = s

End Function
