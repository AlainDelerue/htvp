VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tri_Index_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_Valeurs() As String
Dim m_Nb_Valeurs As Long
Dim m_i_Valeurs() As Long
Dim m_Trié As Boolean

Public Sub Clear()

m_Nb_Valeurs = 0
m_Trié = False

End Sub

Private Sub Class_Initialize()

Clear

End Sub

Public Sub Add(Valeur As String)

m_Nb_Valeurs = m_Nb_Valeurs + 1
ReDim Preserve m_Valeurs(m_Nb_Valeurs)
ReDim Preserve m_i_Valeurs(m_Nb_Valeurs)
m_Valeurs(m_Nb_Valeurs) = Valeur
m_i_Valeurs(m_Nb_Valeurs) = m_Nb_Valeurs
m_Trié = False

End Sub

Public Property Get Nb_Valeurs() As Variant

Nb_Valeurs = m_Nb_Valeurs

End Property

Public Property Get Is_Sorted() As Variant

Is_Sorted = m_Trié

End Property

Public Property Get i_de_Valeur(Valeur As String) As Variant

' Si pas trié, on trie
'---------------------
If Not m_Trié Then Me.Sort
' On recherche par dichotomie
' Si non trouvé on retourne 0
'----------------------------
i_de_Valeur = Recherche_Dicho(Valeur)

End Property

Public Property Get i(index As Long) As Variant

i = m_i_Valeurs(index)

End Property

Public Sub Sort()

If m_Nb_Valeurs > 0 Then QuickSort_Valeurs_i_Valeurs 1, m_Nb_Valeurs
m_Trié = True

End Sub

Private Sub QuickSort_Valeurs_i_Valeurs(inLow As Long, inHi As Long)
      
'----------------------------------------------------------------------
' Tri Les valeurs et leurs index dans l'ordre croissant des valeurs,
' entre les indices inLow et inHi
' procédure à utiliser avec inLow = 1 et inHi = max du tableau
'----------------------------------------------------------------------

Dim pivot   As Variant
Dim tmpSwap As Variant

Dim tmpLow  As Long
Dim tmpHi   As Long
 
tmpLow = inLow
tmpHi = inHi

pivot = m_Valeurs((inLow + inHi) \ 2)

While (tmpLow <= tmpHi)

   While (StrComp(m_Valeurs(tmpLow), pivot, vbTextCompare) < 0) And (tmpLow < inHi)
      tmpLow = tmpLow + 1
   Wend
   
   While (StrComp(pivot, m_Valeurs(tmpHi), vbTextCompare) < 0) And (tmpHi > inLow)
      tmpHi = tmpHi - 1
   Wend

   If (tmpLow <= tmpHi) Then
   
      tmpSwap = m_Valeurs(tmpLow)
      m_Valeurs(tmpLow) = m_Valeurs(tmpHi)
      m_Valeurs(tmpHi) = tmpSwap
      
      tmpSwap = m_i_Valeurs(tmpLow)
      m_i_Valeurs(tmpLow) = m_i_Valeurs(tmpHi)
      m_i_Valeurs(tmpHi) = tmpSwap
      
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
   End If

Wend

If (inLow < tmpHi) Then QuickSort_Valeurs_i_Valeurs inLow, tmpHi
 
If (tmpLow < inHi) Then QuickSort_Valeurs_i_Valeurs tmpLow, inHi

End Sub


Public Property Get Valeur_Triée(index As Long) As Variant

Valeur_Triée = m_Valeurs(index)

End Property

Private Function Recherche_Dicho(Valeur As String)

Dim low As Long
Dim hi As Long
Dim i As Long

' Protection
'-----------
If m_Nb_Valeurs = 0 Then
    Recherche_Dicho = 0
    Exit Function
End If

' Recherche par dichotomie
'-------------------------
low = 1
hi = m_Nb_Valeurs
i = (low + hi) \ 2
While ((low + 1) < hi) And (m_Valeurs(i) <> Valeur)
    If StrComp(m_Valeurs(i), Valeur, vbTextCompare) < 0 Then
        low = i
    Else
        hi = i
    End If
    i = (low + hi) \ 2
Wend
If m_Valeurs(i) = Valeur Then
    Recherche_Dicho = i
ElseIf m_Valeurs(hi) = Valeur Then
    Recherche_Dicho = hi
Else
    Recherche_Dicho = 0
End If

End Function
