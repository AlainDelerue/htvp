VERSION 5.00
Begin VB.Form F_Actions 
   Caption         =   "htvp - Monitoring - Actions to be scheduled"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Bt_Next 
      Caption         =   "Set execution date..."
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Bt_DeleteExistingSchedule 
      Caption         =   "Delete existing schedule"
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Bt_Close 
      Caption         =   "Close scheduler"
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Bt_ListActions_Add 
      Caption         =   "Add action..."
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Bt_ListActions_Clear 
      Caption         =   "Clear actions"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Saisie_ActionListToBeScheduled 
      Height          =   5895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1320
      Width           =   8415
   End
   Begin VB.Label Aff_Instructions 
      Alignment       =   2  'Center
      Caption         =   "Add actions, then click on 'Set execution date'."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   8415
   End
End
Attribute VB_Name = "F_Actions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bt_Close_Click()

Unload F_Schedule
Unload F_Actions
ScheduleActions = ""
ScheduleOn = False

End Sub

Private Sub Bt_DeleteExistingSchedule_Click()

Dim s As String

s = Prefix_cmd & "AbortScheduledAction" & vbCrLf
Set_Clipboard s

End Sub

Private Sub Bt_ListActions_Add_Click()

F_D_Main.SetFocus

End Sub

Private Sub Bt_ListActions_Clear_Click()

ScheduleActions = ""
Saisie_ActionListToBeScheduled.Text = ""

End Sub

Private Sub Bt_Next_Click()

F_Actions.Visible = False
F_Schedule.Show

End Sub

Private Sub Form_Load()

Me.Top = F_D_Main.Top
Me.Left = (Screen.Width - F_Actions.Width) / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)

ScheduleActions = ""
ScheduleOn = False

End Sub
