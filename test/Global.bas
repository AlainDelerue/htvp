Attribute VB_Name = "Global"
Public Const Master = False

Public Type FocusPointType
    X As Single
    Y As Single
End Type
Public t_FocusPoint() As FocusPointType
Public nb_t_FocusPoint As Integer

Declare Function GetCursorPos Lib "user32" _
      (lpPoint As POINTAPI) As Long
      ' Access the GetCursorPos function in user32.dll
      Declare Function SetCursorPos Lib "user32" _
      (ByVal X As Long, ByVal Y As Long) As Long

      ' GetCursorPos requires a variable declared as a custom data type
      ' that will hold two integers, one for x value and one for y value
      Type POINTAPI
         X_Pos As Long
         Y_Pos As Long
      End Type
      
Public Const PI = 3.141593

Public Picfile As String
 
Public BytesI() As Byte
Public BytesF() As Byte

Public F_Width As Single
Public F_Height As Single

Public Current_Sentence As String
Public Current_Focus As Integer
Public Click_Allowed As Boolean
Public Input_Allowed As Boolean
Public Load_In_Progress As Boolean
Public Len_Input_Slave As Integer
Public X_curs_prison As Single
Public Y_curs_prison As Single

Public Lol As Boolean
