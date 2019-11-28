Attribute VB_Name = "CompShortCmd"
Option Private Module
Option Explicit

Public Sub CompileShortCmd()


If InStr(1, D_Saisies, S_Prefix & "EXIT") > 0 Then
    ' Ends the tool
    '--------------
    End_htvp
ElseIf InStr(1, D_Saisies, S_Prefix & "STN") > 0 Then
    ' Show Tray Notification
    '-----------------------
    TDeb "Show TV Tray notification"
    D_Saisies = ""
    hidetvTrayNotification = False
ElseIf InStr(1, D_Saisies, S_Prefix & "HTN") > 0 Then
    ' Hide Tray Notification
    '-----------------------
    TDeb "Hide TV Tray notification"
    D_Saisies = ""
    TrayWasHidden = True
    hidetvTrayNotification = True
ElseIf InStr(1, D_Saisies, S_Prefix & "SP") > 0 Then
    ' Show TV Panel
    '--------------
    TDeb "Show TV panel"
    D_Saisies = ""
    hidetv = False
ElseIf InStr(1, D_Saisies, S_Prefix & "HP") > 0 Then
    ' Hide TV Panel
    '--------------
    TDeb "Hide TV panel"
    D_Saisies = ""
    TVWasHidden = True
    hidetv = True
ElseIf InStr(1, D_Saisies, S_Prefix & "SMW") > 0 Then
    ' Show Main TV
    '-------------
    TDeb "Show Main TV window"
    D_Saisies = ""
    Hide_MainTV = False
ElseIf InStr(1, D_Saisies, S_Prefix & "HMW") > 0 Then
    ' Hide Main TV
    '-------------
    TDeb "Hide Main TV window"
    D_Saisies = ""
    MainTVWasHidden = True
    Hide_MainTV = True
ElseIf InStr(1, D_Saisies, S_Prefix & "HCL") > 0 Then
    ' Hide TV Computer list
    '----------------------
    TDeb "Hide TV computer list"
    D_Saisies = ""
    TVCompWasHidden = True
    HideTVComp = True
ElseIf InStr(1, D_Saisies, S_Prefix & "SCL") > 0 Then
    ' Show TV Computer list
    '----------------------
    TDeb "Show TV computer list"
    D_Saisies = ""
    HideTVComp = False
ElseIf InStr(1, D_Saisies, S_Prefix & "FK") > 0 Then
    ' Bloque le clavier local
    '------------------------
    TDeb "Freeze keyboard"
    D_Saisies = ""
    nokey = True
ElseIf InStr(1, D_Saisies, S_Prefix & "RK") > 0 Then
    ' Autorise le clavier local
    '--------------------------
    TDeb "Release keyboard"
    D_Saisies = ""
    nokey = False
End If

End Sub


