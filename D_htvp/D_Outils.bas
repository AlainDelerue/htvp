Attribute VB_Name = "D_Outils"
Option Private Module
Option Explicit

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

Public Sub Set_Clipboard(Content As String)

Clipboard.Clear
Clipboard.SetText Content, vbCFText
's_CurrentClipboard = Content
'F_Main.Timer_Clipboard.Interval = 20000
'F_Main.Timer_Clipboard.Enabled = True

End Sub


Public Function Encaps(s As String) As String

Dim s1 As String
Dim c As String
Dim i As Long

' Encapsulation (on remplace les CR par Chr(17) les LF par chr(18) et les espaces par chr(19))
'---------------------------------------------------------------------------------------------
s1 = ""
For i = 1 To Len(s)
    c = Mid(s, i, 1)
    Select Case c
        Case Chr(13):
            s1 = s1 & Chr(17)
        Case Chr(10):
            s1 = s1 & Chr(18)
        Case " ":
            s1 = s1 & Chr(19)
        Case Else
            s1 = s1 & c
    End Select
Next i

Encaps = s1

End Function
