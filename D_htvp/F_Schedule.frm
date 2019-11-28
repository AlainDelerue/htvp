VERSION 5.00
Begin VB.Form F_Schedule 
   Caption         =   "htvp - Monitoring - Scheduler"
   ClientHeight    =   8580
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Bt_CloseScheduler 
      Caption         =   "Close scheduler"
      Height          =   495
      Left            =   3360
      TabIndex        =   32
      Top             =   7560
      Width           =   1215
   End
   Begin VB.PictureBox PBoxScheduler 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   240
      ScaleHeight     =   3855
      ScaleWidth      =   6855
      TabIndex        =   15
      Top             =   3480
      Width           =   6855
      Begin VB.TextBox Saisie_Schedule_Message 
         Height          =   1935
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   1440
         Width           =   6855
      End
      Begin VB.TextBox Saisie_Schedule_Code 
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CommandButton Bt_add_one_ 
         Caption         =   "Now"
         Height          =   315
         Index           =   7
         Left            =   3840
         TabIndex        =   20
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Bt_add_one_ 
         Caption         =   "+ 1d"
         Height          =   315
         Index           =   6
         Left            =   6000
         TabIndex        =   19
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Bt_add_one_ 
         Caption         =   "+ 1h"
         Height          =   315
         Index           =   5
         Left            =   5280
         TabIndex        =   18
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Bt_add_one_ 
         Caption         =   "+ 1'"
         Height          =   315
         Index           =   4
         Left            =   4560
         TabIndex        =   17
         Top             =   0
         Width           =   615
      End
      Begin VB.TextBox Saisie_Schedule_DisplayDate 
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Lbl_Schedule 
         Caption         =   "Message:"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   27
         Top             =   1200
         Width           =   4935
      End
      Begin VB.Label Lbl_Schedule 
         Caption         =   "Abort code:"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   26
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Lbl_Schedule 
         Caption         =   "(Slave's timezone)"
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Lbl_Schedule 
         Caption         =   "Display date:"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   24
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Lbl_Schedule 
         Caption         =   "MM/DD/YYYY hh:mm (24h)"
         Height          =   255
         Index           =   5
         Left            =   1710
         TabIndex        =   23
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.CommandButton Bt_add_one_ 
      Caption         =   "Now"
      Height          =   315
      Index           =   3
      Left            =   4080
      TabIndex        =   13
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Bt_Display_ListActions 
      Caption         =   "Actions..."
      Height          =   435
      Left            =   5280
      TabIndex        =   12
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Bt_add_one_ 
      Caption         =   "+ 1d"
      Height          =   315
      Index           =   2
      Left            =   6240
      TabIndex        =   9
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Bt_add_one_ 
      Caption         =   "+ 1h"
      Height          =   315
      Index           =   1
      Left            =   5520
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Bt_add_one_ 
      Caption         =   "+ 1'"
      Height          =   315
      Index           =   0
      Left            =   4800
      TabIndex        =   7
      Top             =   1680
      Width           =   615
   End
   Begin VB.OptionButton Opt_Schedule 
      Caption         =   "Actions + Message"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.OptionButton Opt_Schedule 
      Caption         =   "Actions only"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   240
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton Bt_Schedule 
      Caption         =   "Schedule"
      Height          =   495
      Index           =   0
      Left            =   6000
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Bt_Release_Schedule 
      Caption         =   "Delete existing schedule"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Bt_Schedule 
      Caption         =   "Schedule"
      Height          =   495
      Index           =   1
      Left            =   6000
      TabIndex        =   2
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox Saisie_Schedule_Deadline 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Aff_Instructions 
      Alignment       =   2  'Center
      Caption         =   "This schedule will replace the existing schedule on the remote PC"
      Height          =   615
      Left            =   240
      TabIndex        =   31
      Top             =   7560
      Width           =   2655
   End
   Begin VB.Label Aff_Timelag 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   30
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Aff_Timelag 
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   29
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Line Schedule_line 
      BorderColor     =   &H80000006&
      BorderStyle     =   6  'Inside Solid
      X1              =   0
      X2              =   7320
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Label Lbl_Schedule_Message 
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000006&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   7320
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Lbl_Schedule 
      Caption         =   "(Slave's timezone)"
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Actions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000006&
      BorderStyle     =   6  'Inside Solid
      Index           =   0
      X1              =   0
      X2              =   7320
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Lbl_Schedule 
      Caption         =   "MM/DD/YYYY hh:mm (24h)"
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   10
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Lbl_Schedule 
      Caption         =   "Execution date:"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "F_Schedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bt_add_one__Click(Index As Integer)

Dim d As Date
Dim s_dl As String
Dim s_dd As String

If Trim(Saisie_Schedule_Deadline.Text) = "" Then
    s_dl = Now
Else
    s_dl = Saisie_Schedule_Deadline.Text
End If
If Trim(Saisie_Schedule_DisplayDate.Text) = "" Then
    s_dd = Now
Else
    s_dd = Saisie_Schedule_DisplayDate.Text
End If
Select Case Index
    Case 0:
        d = s_dl
        d = d + TimeSerial(0, 1, 0)
        Saisie_Schedule_Deadline.Text = Format_DateTime(d)
    Case 1:
        d = s_dl
        d = d + TimeSerial(1, 0, 0)
        Saisie_Schedule_Deadline.Text = Format_DateTime(d)
    Case 2:
        d = s_dl
        d = d + 1
        Saisie_Schedule_Deadline.Text = Format_DateTime(d)
    Case 3:
        If Not TimeLagRecieved Then
            F_D_Main.Ask_Remote_Date
        Else
            d = Now + Local_minus_Remote_Time
            Saisie_Schedule_Deadline.Text = Format_DateTime(d)
        End If
    Case 4:
        d = s_dd
        d = d + TimeSerial(0, 1, 0)
        Saisie_Schedule_DisplayDate.Text = Format_DateTime(d)
    Case 5:
        d = s_dd
        d = d + TimeSerial(1, 0, 0)
        Saisie_Schedule_DisplayDate.Text = Format_DateTime(d)
    Case 6:
        d = s_dd
        d = d + 1
        Saisie_Schedule_DisplayDate.Text = Format_DateTime(d)
    Case 7:
        If Not TimeLagRecieved Then
            F_D_Main.Ask_Remote_Date
        Else
            d = Now + Local_minus_Remote_Time
            Saisie_Schedule_DisplayDate.Text = Format_DateTime(d)
        End If
End Select

End Sub

Private Function Format_DateTime(d As Date) As String

Format_DateTime = Format(d, "mm") & "/" _
                                & Format(d, "dd") & "/" _
                                & Format(d, "yyyy") & " " _
                                & Format(d, "hh") & ":" _
                                & Format(d, "nn")

End Function

Private Sub Bt_Cancel_Click()

Unload Me

End Sub

Private Sub Bt_in_one__Click(Index As Integer)

Dim d As Date

Select Case Index
    Case 0:
        d = Now + TimeSerial(0, 1, 0)
    Case 1:
        d = Now + TimeSerial(1, 0, 0)
    Case 2:
        d = Now + 1
End Select
Saisie_Schedule_Deadline.Text = Format(d, "mm") & "/" _
                                & Format(d, "dd") & "/" _
                                & Format(d, "yyyy") & " " _
                                & Format(d, "hh") & ":" _
                                & Format(d, "nn")

End Sub

Private Sub Bt_CloseScheduler_Click()

Unload F_Schedule
Unload F_Actions
ScheduleActions = ""
ScheduleOn = False

End Sub

Private Sub Bt_Display_ListActions_Click()

F_Schedule.Visible = False
F_Actions.Show
F_Actions.Saisie_ActionListToBeScheduled.Text = ScheduleActions

End Sub

Private Sub Bt_Release_Schedule_Click()

Dim s As String

s = Prefix_cmd & "AbortScheduledAction" & vbCrLf
Set_Clipboard s

Unload Me

End Sub

Private Sub Bt_Schedule_Click(Index As Integer)

Dim s As String
Dim c As String
Dim s1 As String
Dim I As Long
Dim s2 As String
Dim s3 As String
Dim s4 As String

If ScheduleActions = "" Then
    MsgBox "No action selected.", vbOKOnly, Me.Caption
    Exit Sub
End If
s1 = Encaps(ScheduleActions)

If Index = 0 Then
    s2 = ""
    s1 = s1 & " """ & s2 & """ """ & s2 & """ """ & Me.Saisie_Schedule_Deadline & """ """ & s2 & """" & vbCrLf
Else
    s1 = s1 & " """ & Encaps(Me.Saisie_Schedule_Message) & """ """ & Me.Saisie_Schedule_DisplayDate & """ """ & Me.Saisie_Schedule_Deadline & """ """ & Me.Saisie_Schedule_Code & """" & vbCrLf
End If
s = Prefix_cmd & "ScheduleAction " & s1
Set_Clipboard s

Unload Me

End Sub

Private Sub Bt_ScheduleExample_Click(Index As Integer)

Dim d As Date

If Index = 0 Then
    Saisie_Schedule_Message.Text = "You did great yesterday slave. Such a submissiveness is a turn on. Keep going  :-)" & vbCrLf _
                            & "This being said, I have good news: Starting now, I can use your computer to schedule your tributes and to trigger punishments in case of delay. A lot easier for me. Let's try it:" & vbCrLf & vbCrLf _
                            & "YOU HAVE UNTIL TOMORRW 6 PM YOUR TIME TO SEND ME $30." & vbCrLf & vbCrLf _
                            & "You are so fucked, asshole!" & vbCrLf _
                            & "Be perfect and even if i love to make you suffer, I COULD decide to be nicer with you :DDD"
    Saisie_Schedule_Code.Text = "GoodPig"
    d = Format(Now + 1, "mm/dd/yyyy") & " 18:00"
    Saisie_Schedule_Deadline.Text = Format(d, "mm") & "/" _
                                & Format(d, "dd") & "/" _
                                & Format(d, "yyyy") & " " _
                                & Format(d, "hh") & ":" _
                                & Format(d, "nn")
ElseIf Index = 1 Then
    Saisie_Schedule_Message.Text = "Hello asshole." & vbCrLf & vbCrLf _
                            & "I want a pic of you naked in my mailbox bye tomorrow 9 am your time otherwize your PC will be totally blocked until you pay me a tax of 50 Euros." & vbCrLf & vbCrLf _
                            & "Be carefull, any complain would increase the tax..." & vbCrLf _
                            & "Lol, Im sure you love it. Anyway, I'm serious and I won't change my mind." & vbCrLf _
                            & "Your Mistress"
    Saisie_Schedule_Code.Text = "Pathetic slave"
    d = Format(Now + 1, "mm/dd/yyyy") & " 09:00"
    Saisie_Schedule_Deadline.Text = Format(d, "mm") & "/" _
                                & Format(d, "dd") & "/" _
                                & Format(d, "yyyy") & " " _
                                & Format(d, "hh") & ":" _
                                & Format(d, "nn")
ElseIf Index = 2 Then
    Saisie_Schedule_Message.Text = "You've disobeyed me." & vbCrLf & vbCrLf _
                            & "1/ I want you to send me a long letter of apology" & vbCrLf _
                            & "2/ Your PC is going to lock you out at midnight" & vbCrLf _
                            & "3/ Depending on the content of your letter I could decide to release your PC or to keep it locked for weeks" & vbCrLf & vbCrLf _
                            & "Your upset Mistress"
    Saisie_Schedule_Code.Text = "Asshole"
    d = Format(Now + 1, "mm/dd/yyyy") & " 00:00"
    Saisie_Schedule_Deadline.Text = Format(d, "mm") & "/" _
                                & Format(d, "dd") & "/" _
                                & Format(d, "yyyy") & " " _
                                & Format(d, "hh") & ":" _
                                & Format(d, "nn")
ElseIf Index = 3 Then
    Saisie_Schedule_Message.Text = "You have 3 days to tribute me asshole."
    Saisie_Schedule_Code.Text = "ThankYouSlut"
    d = Now + 3
    Saisie_Schedule_Deadline.Text = Format(d, "mm") & "/" _
                                & Format(d, "dd") & "/" _
                                & Format(d, "yyyy") & " " _
                                & Format(d, "hh") & ":" _
                                & Format(d, "nn")
End If

End Sub


Private Sub Form_Load()

Dim s As String
Dim d As Date

Aff_Timelag(0).Caption = ""
Aff_Timelag(1).Caption = ""


s = "Example:" & vbCrLf & vbCrLf
s = s & "You have until tomorrow 6pm your time to send me a new pic of you naked." & vbCrLf & vbCrLf
s = s & "Otherwise you'll be locked out of your computer for weeks and you'll be in trouble to make me change my mind!" & vbCrLf & vbCrLf
Me.Saisie_Schedule_Message.Text = s
Me.Saisie_Schedule_Message.ForeColor = RGB(192, 192, 192)


F_D_Main.Ask_Remote_Date

Mode_Schedule_Simple

Me.Top = F_D_Main.Top
Me.Left = (Screen.Width - F_Schedule.Width) / 2

End Sub

Public Sub Mode_Schedule_Simple()

'Dim i As Integer

Me.Height = 4050
Bt_Schedule(0).Visible = True
'Bt_Cancel.Top = Bt_Schedule(0).Top
'Bt_Cancel.Left = Bt_Schedule(0).Left - Bt_Cancel.Width - 120
Bt_Release_Schedule.Top = Bt_Schedule(0).Top
Bt_CloseScheduler.Top = Bt_Release_Schedule.Top
Aff_Instructions.Top = Bt_CloseScheduler.Top
Bt_Release_Schedule.Left = Bt_Schedule(0).Left - Bt_Release_Schedule.Width - 120

Me.Schedule_line.Visible = False
Me.PBoxScheduler.Visible = False
Me.Lbl_Schedule_Message.Visible = False
Me.Aff_Timelag(1).Visible = False
Me.Bt_Schedule(1).Visible = False
Me.Lbl_Schedule(0).Visible = False
Me.Lbl_Schedule(2).Visible = False
'Me.Lbl_ScheduleDoc.Visible = False
Me.Saisie_Schedule_Code.Visible = False
Me.Saisie_Schedule_Message.Visible = False
'For i = 0 To 3
'    Bt_ScheduleExample(i).Visible = False
'Next i

End Sub

Public Sub Mode_Enforce()

'Dim i As Integer

Me.Height = 8745
Bt_Schedule(0).Visible = False
'Bt_Cancel.Top = Bt_Schedule(1).Top
'Bt_Cancel.Left = Bt_Schedule(1).Left - Bt_Cancel.Width - 120
Bt_Release_Schedule.Top = Bt_Schedule(1).Top
Bt_CloseScheduler.Top = Bt_Release_Schedule.Top
Aff_Instructions.Top = Bt_CloseScheduler.Top
Bt_Release_Schedule.Left = Bt_Schedule(1).Left - Bt_Release_Schedule.Width - 120

Me.Schedule_line.Visible = True
Me.PBoxScheduler.Visible = True
Me.Lbl_Schedule_Message.Visible = True
Me.Aff_Timelag(1).Visible = True
Me.Bt_Schedule(1).Visible = True
Me.Lbl_Schedule(0).Visible = True
Me.Lbl_Schedule(2).Visible = True
'Me.Lbl_ScheduleDoc.Visible = True
Me.Saisie_Schedule_Code.Visible = True
Me.Saisie_Schedule_Message.Visible = True
'For i = 0 To 3
'    Bt_ScheduleExample(i).Visible = True
'Next i


End Sub

Private Sub Form_Unload(Cancel As Integer)

ScheduleActions = ""
ScheduleOn = False
Unload F_Actions

End Sub

Private Sub Opt_Schedule_Click(Index As Integer)

If Index = 0 Then
    Mode_Schedule_Simple
Else
    Mode_Enforce
End If

End Sub


Private Sub Saisie_Schedule_Message_Click()

If Me.Saisie_Schedule_Message.ForeColor = RGB(192, 192, 192) Then
    Me.Saisie_Schedule_Message.Text = ""
    Me.Saisie_Schedule_Message.ForeColor = 0
End If

End Sub
