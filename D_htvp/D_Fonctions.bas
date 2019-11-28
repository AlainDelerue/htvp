Attribute VB_Name = "D_Fonctions"
Option Private Module
Option Explicit

Public Function Prefix_cmd()

Dim i As Integer
Dim M_Version As Integer

' Le préfixe des commandes dépend de la version du distant,
' demandée précédemment avec le plus ancien format
'---------------------------------------------------------
If Remote_Version = "" Then
    Prefix_cmd = Trim(L_Prefix & " " & PC_Name & " " & F_D_Main.Saisie_ToolPW.Text) & vbCrLf
Else
    i = InStr(1, Remote_Version, ".")
    If i > 1 Then
        M_Version = Val(Mid(Remote_Version, 2, i - 2))
        If M_Version >= 8 Then
            Prefix_cmd = Trim(This_PC_Name & " " & L_Prefix & " " & PC_Name & " " & F_D_Main.Saisie_ToolPW.Text) & vbCrLf
        Else
            Prefix_cmd = Trim(L_Prefix & " " & PC_Name & " " & F_D_Main.Saisie_ToolPW.Text) & vbCrLf
        End If
    Else
        Prefix_cmd = Trim(L_Prefix & " " & PC_Name & " " & F_D_Main.Saisie_ToolPW.Text) & vbCrLf
    End If
End If

End Function
