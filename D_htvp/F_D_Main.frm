VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form F_D_Main 
   Caption         =   "htvp - Monitoring"
   ClientHeight    =   9735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6375
   Icon            =   "F_D_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   6375
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   615
      Left            =   120
      TabIndex        =   199
      Top             =   1380
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   1085
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Bt_GethtvpStatus 
      Caption         =   "Get htvp status"
      Height          =   255
      Left            =   3480
      TabIndex        =   195
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Bt_Scheduler 
      Caption         =   "Scheduler..."
      Height          =   255
      Left            =   4920
      TabIndex        =   193
      Top             =   1080
      Width           =   1335
   End
   Begin VB.PictureBox picCombined 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   6480
      ScaleHeight     =   1635
      ScaleWidth      =   1635
      TabIndex        =   157
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Bt_CheckVersion 
      Caption         =   "Check -->"
      Height          =   255
      Left            =   3480
      TabIndex        =   74
      Top             =   720
      Width           =   855
   End
   Begin VB.Timer Hrlg_Timeout_Send 
      Enabled         =   0   'False
      Left            =   720
      Top             =   120
   End
   Begin VB.TextBox Saisie_ToolPW 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   30
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox Saisie_PC 
      Height          =   315
      ItemData        =   "F_D_Main.frx":1CCA
      Left            =   1920
      List            =   "F_D_Main.frx":1CCC
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin VB.Timer hrlg_Clipboard 
      Interval        =   100
      Left            =   240
      Top             =   120
   End
   Begin VB.Frame Frame_Aff 
      Height          =   7860
      Index           =   13
      Left            =   8520
      TabIndex        =   17
      Top             =   5640
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton Bt_ClearLogs 
         Caption         =   "Clear"
         Height          =   495
         Left            =   9960
         TabIndex        =   28
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton Bt_SaveLogs 
         Caption         =   "Save"
         Height          =   495
         Left            =   9960
         TabIndex        =   27
         Top             =   5040
         Width           =   1935
      End
      Begin VB.CommandButton Bt_LoadLogs 
         Caption         =   "Load"
         Height          =   495
         Left            =   9960
         TabIndex        =   26
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox Aff_Logs 
         BackColor       =   &H00F8F8F8&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3840
         Width           =   9495
      End
      Begin VB.ComboBox Saisie_Action_Command 
         Height          =   315
         ItemData        =   "F_D_Main.frx":1CCE
         Left            =   1320
         List            =   "F_D_Main.frx":1CD0
         TabIndex        =   2
         Top             =   1560
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox Saisie_Action_Param 
         Height          =   285
         Index           =   1
         Left            =   4800
         TabIndex        =   3
         Top             =   1560
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Saisie_Action_Param 
         Height          =   285
         Index           =   2
         Left            =   7080
         TabIndex        =   4
         Top             =   1560
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Saisie_Action_Param 
         Height          =   285
         Index           =   3
         Left            =   9360
         TabIndex        =   5
         Top             =   1560
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Bt_Go_SettingFunction 
         Caption         =   "Go"
         Height          =   495
         Left            =   1320
         TabIndex        =   6
         Top             =   2040
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.ComboBox Saisie_SettingCategory 
         Height          =   315
         ItemData        =   "F_D_Main.frx":1CD2
         Left            =   4080
         List            =   "F_D_Main.frx":1CD4
         TabIndex        =   1
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   "Logs"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Lbl_Command 
         Caption         =   "Command"
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   1320
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Aff_Action_Help 
         Height          =   735
         Left            =   4800
         TabIndex        =   22
         Top             =   2160
         Width           =   6615
      End
      Begin VB.Label Lbl_Setting_Functions_Param 
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   21
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Lbl_Setting_Functions_Param 
         Height          =   255
         Index           =   2
         Left            =   7200
         TabIndex        =   20
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Lbl_Setting_Functions_Param 
         Height          =   255
         Index           =   3
         Left            =   9480
         TabIndex        =   19
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Function category:"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   630
         Width           =   1335
      End
   End
   Begin VB.Frame Frame_Aff 
      Height          =   7860
      Index           =   9
      Left            =   120
      TabIndex        =   114
      Top             =   1920
      Width           =   6135
      Begin VB.Frame Frame17 
         Caption         =   "Lock/Unlock the Task Manager "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   196
         Top             =   600
         Width           =   5655
         Begin VB.CommandButton Bt_ReleaseTskMgr 
            Caption         =   "Release Task Manager"
            Height          =   495
            Left            =   3000
            TabIndex        =   198
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton Bt_LockTskMgr 
            Caption         =   "Lock Task Manager"
            Height          =   495
            Left            =   720
            TabIndex        =   197
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Lock/Unlock the keyboard "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   137
         Top             =   6360
         Visible         =   0   'False
         Width           =   5655
         Begin VB.CommandButton Bt_LockKeyboard 
            Caption         =   "Lock Keyboard"
            Height          =   495
            Left            =   720
            TabIndex        =   139
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton Bt_ReleaseKeyboard 
            Caption         =   "Release Keyboard"
            Height          =   495
            Left            =   3000
            TabIndex        =   138
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Lock/Unlock remote files "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   240
         TabIndex        =   115
         Top             =   2040
         Width           =   5655
         Begin VB.CommandButton Bt_GoLockFiles 
            Caption         =   "Unlock"
            Height          =   495
            Index           =   1
            Left            =   3120
            TabIndex        =   122
            Top             =   2880
            Width           =   1095
         End
         Begin VB.CommandButton Bt_GoLockFiles 
            Caption         =   "Lock"
            Height          =   495
            Index           =   0
            Left            =   4320
            TabIndex        =   121
            Top             =   2880
            Width           =   1095
         End
         Begin VB.TextBox Saisie_LockFiles 
            Height          =   285
            Left            =   240
            TabIndex        =   120
            Top             =   600
            Width           =   5085
         End
         Begin VB.CheckBox Chk_Prefix 
            Caption         =   "Add prefix to file names when locked:"
            Height          =   195
            Left            =   240
            TabIndex        =   119
            Top             =   1200
            Value           =   1  'Checked
            Width           =   3375
         End
         Begin VB.TextBox Saisie_Prefix 
            Height          =   285
            Left            =   240
            TabIndex        =   118
            Text            =   "Locked by Mistress - "
            Top             =   1440
            Width           =   3525
         End
         Begin VB.TextBox Saisie_Lock_PW 
            Height          =   285
            Left            =   240
            TabIndex        =   117
            Text            =   "Such a naughty boy"
            Top             =   2280
            Width           =   3525
         End
         Begin VB.CommandButton Bt_Help_Lockfiles 
            Caption         =   "Read me!"
            Height          =   495
            Left            =   240
            TabIndex        =   116
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Lbl_LockFileOrFolder 
            Caption         =   "Remote file or folder:"
            Height          =   255
            Left            =   240
            TabIndex        =   124
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Ciphering password:"
            Height          =   255
            Left            =   240
            TabIndex        =   123
            Top             =   2040
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame_Aff 
      Height          =   7860
      Index           =   2
      Left            =   240
      TabIndex        =   32
      Top             =   2040
      Width           =   6735
      Begin VB.CommandButton Bt_MainTVParam 
         Caption         =   "Set new values"
         Height          =   735
         Left            =   4800
         TabIndex        =   148
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox Chk_TVChangesRequireAdmin 
         Caption         =   "Changes require administrative rights"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1680
         TabIndex        =   147
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CheckBox Chk_Lock_TV_Options 
         Caption         =   "TeamViewer options protected"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1680
         TabIndex        =   146
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CheckBox Chk_TVStartsWithWindows 
         Caption         =   "Start TeamViewer with Windows"
         Height          =   255
         Left            =   1680
         TabIndex        =   145
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CommandButton Bt_RefreshCurrentTVValues 
         Caption         =   "Get current values"
         Height          =   735
         Left            =   240
         TabIndex        =   144
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Bt_TV_Reboot 
         Caption         =   "Reboot"
         Height          =   495
         Left            =   2520
         TabIndex        =   141
         Top             =   6360
         Width           =   1095
      End
      Begin VB.CommandButton Bt_HideShowTVWindows 
         Caption         =   "Execute"
         Height          =   375
         Left            =   4560
         TabIndex        =   98
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CheckBox Chk_HideTVNotifications 
         Caption         =   "Show"
         Height          =   255
         HelpContextID   =   1
         Index           =   1
         Left            =   3000
         TabIndex        =   47
         Top             =   4335
         Width           =   735
      End
      Begin VB.CheckBox Chk_Permanent 
         Caption         =   "Permanently"
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   46
         Top             =   4335
         Width           =   1335
      End
      Begin VB.CheckBox Chk_Permanent 
         Caption         =   "Permanently"
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   45
         Top             =   4095
         Width           =   1215
      End
      Begin VB.CheckBox Chk_Permanent 
         Caption         =   "Permanently"
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   44
         Top             =   3855
         Width           =   1335
      End
      Begin VB.CheckBox Chk_Permanent 
         Caption         =   "Permanently"
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   43
         Top             =   3615
         Width           =   1335
      End
      Begin VB.CheckBox Chk_HideTVNotifications 
         Caption         =   "Hide"
         Height          =   255
         HelpContextID   =   1
         Index           =   0
         Left            =   2160
         TabIndex        =   42
         Top             =   4335
         Width           =   735
      End
      Begin VB.CheckBox Chk_HideMainTVWindow 
         Caption         =   "Show"
         Height          =   255
         HelpContextID   =   1
         Index           =   1
         Left            =   3000
         TabIndex        =   41
         Top             =   4095
         Width           =   855
      End
      Begin VB.CheckBox Chk_HideMainTVWindow 
         Caption         =   "Hide"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   40
         Top             =   4095
         Width           =   735
      End
      Begin VB.CheckBox Chk_HideComputerList 
         Caption         =   "Show"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   39
         Top             =   3855
         Width           =   855
      End
      Begin VB.CheckBox Chk_HideComputerList 
         Caption         =   "Hide"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   38
         Top             =   3855
         Width           =   855
      End
      Begin VB.CheckBox Chk_HideTVPanel 
         Caption         =   "Show"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   37
         Top             =   3615
         Width           =   855
      End
      Begin VB.CheckBox Chk_HideTVPanel 
         Caption         =   "Hide"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   36
         Top             =   3615
         Width           =   855
      End
      Begin VB.Label Lbl_CurrentTVValues 
         Caption         =   "Current values:"
         Height          =   255
         Left            =   1680
         TabIndex        =   149
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "Warning: changes on blue items require a reboot to become effective"
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   3240
         TabIndex        =   140
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label11 
         Caption         =   "Hide or Show TeamViewer windows"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   126
         Top             =   3240
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "TeamViewer parameters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   125
         Top             =   480
         Width           =   2535
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000006&
         BorderStyle     =   6  'Inside Solid
         Index           =   4
         X1              =   120
         X2              =   5880
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000006&
         BorderStyle     =   6  'Inside Solid
         Index           =   3
         X1              =   120
         X2              =   5880
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label10 
         Caption         =   "TeamViewer notifications:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   51
         Top             =   4350
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Main TeamViewer Window:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   50
         Top             =   4110
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Computers && Contacts:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Top             =   3870
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "TeamViewer panel:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   3630
         Width           =   1935
      End
      Begin VB.Label Lbl_Setting_Functions_Param 
         Height          =   255
         Index           =   7
         Left            =   9480
         TabIndex        =   34
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Lbl_Setting_Functions_Param 
         Height          =   255
         Index           =   6
         Left            =   7200
         TabIndex        =   33
         Top             =   1320
         Width           =   2175
      End
   End
   Begin VB.Frame Frame_Aff 
      Height          =   7860
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   8775
      Begin VB.CheckBox Chk_Logoff 
         Caption         =   "Add a logoff to secure the blue changes above"
         Height          =   255
         Left            =   240
         TabIndex        =   156
         Top             =   6480
         Width           =   3735
      End
      Begin VB.CheckBox Chk_Reboot 
         Caption         =   "Add a shutdown to secure the red and blue changes above"
         Height          =   495
         Left            =   240
         TabIndex        =   155
         Top             =   6840
         Width           =   2775
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   240
         ScaleHeight     =   735
         ScaleWidth      =   2055
         TabIndex        =   151
         Top             =   2760
         Width           =   2055
         Begin VB.OptionButton Chk_Current 
            Caption         =   "Keep current type"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   154
            Top             =   0
            Width           =   1935
         End
         Begin VB.OptionButton Chk_Current 
            Caption         =   "Make it standard"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   153
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton Chk_Current 
            Caption         =   "Make it admin"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   152
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.CommandButton Bt_ExecAllAccounts 
         Caption         =   "Execute all"
         Height          =   495
         Left            =   4200
         TabIndex        =   97
         Top             =   6720
         Width           =   1575
      End
      Begin VB.CommandButton Bt_OtherAccounts 
         Caption         =   "Execute"
         Height          =   375
         Left            =   5040
         TabIndex        =   96
         Top             =   5640
         Width           =   735
      End
      Begin VB.CommandButton Bt_CurAccount 
         Caption         =   "Execute"
         Height          =   375
         Left            =   5040
         TabIndex        =   95
         Top             =   3600
         Width           =   735
      End
      Begin VB.CommandButton Bt_DomAdmin 
         Caption         =   "Execute"
         Height          =   375
         Left            =   5040
         TabIndex        =   94
         Top             =   1440
         Width           =   735
      End
      Begin VB.OptionButton Chk_OtherAccounts 
         Caption         =   "Make them standard"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   93
         Top             =   5160
         Width           =   1935
      End
      Begin VB.OptionButton Chk_OtherAccounts 
         Caption         =   "Keep current type"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   92
         Top             =   4920
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox Saisie_Admin_Name 
         Height          =   285
         Left            =   840
         TabIndex        =   86
         Top             =   960
         Width           =   1650
      End
      Begin VB.CheckBox Chk_ChgeAdminPW 
         Caption         =   "Set password:"
         Height          =   255
         Left            =   2760
         TabIndex        =   85
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Saisie_Admin_PW 
         Height          =   285
         Left            =   4200
         TabIndex        =   84
         Top             =   960
         Width           =   1650
      End
      Begin VB.TextBox Saisie_Current_PW 
         Height          =   285
         Left            =   4200
         TabIndex        =   83
         Top             =   3000
         Width           =   1650
      End
      Begin VB.CheckBox Chk_ChgeCurrentPW 
         Caption         =   "Set password:"
         Height          =   255
         Left            =   2760
         TabIndex        =   82
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Saisie_Others_PW 
         Height          =   285
         Left            =   4200
         TabIndex        =   81
         Top             =   5040
         Width           =   1650
      End
      Begin VB.CheckBox Chk_ChgeOthersPW 
         Caption         =   "Set passwords:"
         Height          =   255
         Left            =   2760
         TabIndex        =   80
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000006&
         BorderStyle     =   6  'Inside Solid
         Index           =   5
         X1              =   120
         X2              =   5880
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Label Label14 
         Caption         =   "The red item requires a reboot to become effective"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   143
         Top             =   5640
         Width           =   3735
      End
      Begin VB.Label Label14 
         Caption         =   "Blue items require a logoff to become effective"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   142
         Top             =   3720
         Width           =   3735
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000006&
         BorderStyle     =   6  'Inside Solid
         Index           =   2
         X1              =   120
         X2              =   5880
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000006&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   120
         X2              =   5880
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000006&
         BorderStyle     =   6  'Inside Solid
         Index           =   0
         X1              =   120
         X2              =   5880
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label19 
         Caption         =   "(Created if doesn't exist)"
         Height          =   375
         Left            =   2880
         TabIndex        =   88
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label16 
         Caption         =   "Mistress/Master admin account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label17 
         Caption         =   "Current account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "All accounts other than Mistress/Master and Current ones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   4560
         Width           =   5415
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Name: "
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Frame Frame_Aff 
      Height          =   7860
      Index           =   7
      Left            =   120
      TabIndex        =   78
      Top             =   1800
      Width           =   6135
      Begin VB.Frame Frame10 
         Caption         =   "Release Welcome screen "
         Height          =   855
         Left            =   3240
         TabIndex        =   178
         Top             =   6600
         Width           =   2295
         Begin VB.CommandButton Bt_GoBackground 
            Caption         =   "Recover"
            Height          =   375
            Index           =   3
            Left            =   600
            TabIndex        =   179
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Release Wallpaper "
         Height          =   855
         Left            =   600
         TabIndex        =   175
         Top             =   6600
         Width           =   2295
         Begin VB.CommandButton Bt_GoBackground 
            Caption         =   "Recover"
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   177
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Bt_GoBackground 
            Caption         =   "Remove"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   176
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "1/ Choose the picture "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   600
         TabIndex        =   166
         Top             =   360
         Width           =   4935
         Begin VB.CommandButton Bt_AddBackGround 
            Caption         =   "Take it"
            Height          =   375
            Left            =   2535
            TabIndex        =   173
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton Bt_ClearBackground 
            Caption         =   "Clear"
            Height          =   375
            Left            =   1335
            TabIndex        =   172
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton opt_WelcomeAlignment 
            Caption         =   "Right Justify"
            Height          =   255
            Index           =   2
            Left            =   3255
            TabIndex        =   171
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton opt_WelcomeAlignment 
            Caption         =   "Center"
            Height          =   255
            Index           =   1
            Left            =   2055
            TabIndex        =   170
            Top             =   720
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton opt_WelcomeAlignment 
            Caption         =   "Left Justify"
            Height          =   255
            Index           =   0
            Left            =   735
            TabIndex        =   169
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox Saisie_WelcomeBackground 
            Height          =   285
            Left            =   735
            TabIndex        =   168
            Top             =   360
            Width           =   3330
         End
         Begin VB.CommandButton Bt_BrowseWelcomeBackground 
            Caption         =   "..."
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
            Left            =   4095
            TabIndex        =   167
            Top             =   360
            Width           =   375
         End
         Begin VB.Image Aff_Preview_Background 
            BorderStyle     =   1  'Fixed Single
            Height          =   1815
            Left            =   120
            Stretch         =   -1  'True
            Top             =   1680
            Width           =   4695
         End
         Begin VB.Label Label20 
            Caption         =   "Picture:"
            Height          =   255
            Left            =   135
            TabIndex        =   174
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "2/ Upload and set up "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   600
         TabIndex        =   158
         Top             =   4320
         Width           =   4935
         Begin VB.Frame Frame5 
            Caption         =   "Welcome screen (W7 only) "
            Height          =   1455
            Left            =   2520
            TabIndex        =   163
            Top             =   360
            Width           =   2175
            Begin VB.CheckBox Chk_ForceWelcomeScreen 
               Caption         =   "Make it permanent"
               Height          =   195
               Left            =   240
               TabIndex        =   180
               Top             =   360
               Width           =   1695
            End
            Begin VB.CommandButton Bt_GoBackground 
               Caption         =   "Upload && Set"
               Height          =   375
               Index           =   4
               Left            =   240
               TabIndex        =   164
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label Aff_WelcomeBackgroundProgression 
               Alignment       =   2  'Center
               Height          =   255
               Left            =   120
               TabIndex        =   165
               Top             =   1080
               Width           =   1935
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Wallpaper "
            Height          =   1455
            Left            =   240
            TabIndex        =   159
            Top             =   360
            Width           =   2175
            Begin VB.CheckBox Chk_ForceWallPaper 
               Caption         =   "Make it permanent"
               Height          =   195
               Left            =   240
               TabIndex        =   161
               Top             =   360
               Width           =   1695
            End
            Begin VB.CommandButton Bt_GoBackground 
               Caption         =   "Upload && Set"
               Height          =   375
               Index           =   0
               Left            =   240
               TabIndex        =   160
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label Aff_WallpaperProgression 
               Alignment       =   2  'Center
               Height          =   255
               Left            =   120
               TabIndex        =   162
               Top             =   1080
               Width           =   1935
            End
         End
      End
   End
   Begin VB.Frame Frame_Aff 
      Height          =   7860
      Index           =   6
      Left            =   120
      TabIndex        =   77
      Top             =   2040
      Width           =   6135
      Begin VB.Frame Frame3 
         Caption         =   "Predefined Web pages "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   127
         Top             =   3240
         Width           =   5655
         Begin VB.CommandButton Bt_Paypal 
            Caption         =   "Open Paypal"
            Height          =   495
            Left            =   240
            TabIndex        =   130
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Bt_Amazon 
            Caption         =   "Open Amazon"
            Height          =   495
            Left            =   2040
            TabIndex        =   129
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Bt_AmazonUK 
            Caption         =   "Open Amazon UK"
            Height          =   495
            Left            =   3840
            TabIndex        =   128
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Open a Web page "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   99
         Top             =   960
         Width           =   5655
         Begin VB.TextBox Saisie_URL 
            Height          =   285
            Left            =   120
            TabIndex        =   101
            Top             =   600
            Width           =   5325
         End
         Begin VB.CommandButton Bt_GoWebPage 
            Caption         =   "Open"
            Height          =   375
            Left            =   4560
            TabIndex        =   100
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label22 
            Caption         =   "Web page:"
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame_Aff 
      Height          =   7860
      Index           =   8
      Left            =   120
      TabIndex        =   76
      Top             =   1920
      Width           =   6135
      Begin VB.Frame Frame6 
         Caption         =   "Hide / Show the TeamViewer panel "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   720
         TabIndex        =   131
         Top             =   960
         Width           =   4455
         Begin VB.CommandButton Bt_HidePanel 
            Caption         =   "   Hide TV panel    Current Windows session"
            Height          =   495
            Left            =   240
            TabIndex        =   135
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton Bt_ShowPanel 
            Caption         =   "   Show TV panel    Current Windows session"
            Height          =   495
            Left            =   2280
            TabIndex        =   134
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton Bt_HidePanelPermanent 
            Caption         =   "Hide TV panel Permanently"
            Height          =   495
            Left            =   240
            TabIndex        =   133
            Top             =   960
            Width           =   1935
         End
         Begin VB.CommandButton Bt_ShowPanelPermanent 
            Caption         =   "Show TV panel Permanently"
            Height          =   495
            Left            =   2280
            TabIndex        =   132
            Top             =   960
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame_Aff 
      Height          =   7860
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   6135
      Begin VB.CommandButton Bt_Get_htvp_Status 
         Caption         =   "Get htvp status"
         Height          =   495
         Left            =   120
         TabIndex        =   194
         Top             =   240
         Width           =   1335
      End
      Begin VB.Frame Frame_GetAccountDetails 
         Caption         =   "Get Windows account details"
         Height          =   855
         Left            =   3000
         TabIndex        =   14
         Top             =   480
         Width           =   3015
         Begin VB.TextBox Saisie_Account_For_Details 
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   2055
         End
         Begin VB.CommandButton Bt_GetAccountDetails 
            Caption         =   "Get"
            Height          =   495
            Left            =   2280
            TabIndex        =   11
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   " Account"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CommandButton Bt_GetTVParameters 
         Caption         =   "Get TeamViewer parameters"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Bt_GetAccountsList 
         Caption         =   "Get Windows accounts list"
         Height          =   495
         Left            =   1560
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Aff_Global_functions 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6255
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1440
         Width           =   5895
      End
   End
   Begin VB.Frame Frame_Aff 
      Height          =   7860
      Index           =   11
      Left            =   120
      TabIndex        =   29
      Top             =   1920
      Width           =   6135
      Begin VB.Frame Frame15 
         Caption         =   "Protect the remote tool with a password "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   1200
         TabIndex        =   63
         Top             =   960
         Width           =   3735
         Begin VB.CheckBox Chk_SavePW 
            Caption         =   "Save password"
            Height          =   255
            Left            =   1680
            TabIndex        =   150
            Top             =   1920
            Width           =   1815
         End
         Begin VB.TextBox Saisie_CurrentPW 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1680
            PasswordChar    =   "*"
            TabIndex        =   69
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox Saisie_NewPW 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   1680
            PasswordChar    =   "*"
            TabIndex        =   67
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox Saisie_NewPW 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1680
            PasswordChar    =   "*"
            TabIndex        =   65
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CommandButton Bt_SetPW 
            Caption         =   "Ok"
            Height          =   375
            Left            =   2520
            TabIndex        =   64
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Aff_PC_Protected 
            Alignment       =   2  'Center
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   73
            Top             =   3240
            Width           =   2295
         End
         Begin VB.Label Aff_PC_Protected 
            Alignment       =   2  'Center
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   72
            Top             =   2760
            Width           =   3135
         End
         Begin VB.Label Aff_PC_Protected 
            Alignment       =   2  'Center
            Caption         =   "Protection of:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   71
            Top             =   2400
            Width           =   3135
         End
         Begin VB.Label Label27 
            Caption         =   "Current password:"
            Height          =   375
            Left            =   240
            TabIndex        =   70
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label26 
            Caption         =   "Confirm password:"
            Height          =   375
            Left            =   240
            TabIndex        =   68
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label25 
            Caption         =   "New password:"
            Height          =   375
            Left            =   240
            TabIndex        =   66
            Top             =   1080
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame_Aff 
      Height          =   7860
      Index           =   0
      Left            =   0
      TabIndex        =   60
      Top             =   2040
      Width           =   6135
      Begin VB.CommandButton Bt_OK_Accueil 
         Caption         =   "OK"
         Height          =   495
         Left            =   2400
         TabIndex        =   136
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "Do not forget to select the right PC name and enter the right password if any"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   61
         Top             =   3600
         Width           =   5295
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Caption         =   "Mistress console"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   62
         Top             =   1560
         Width           =   5895
      End
   End
   Begin VB.Frame Frame_Aff 
      Height          =   7860
      Index           =   4
      Left            =   120
      TabIndex        =   79
      Top             =   1920
      Width           =   6135
      Begin VB.Frame Frame11 
         Caption         =   "Upload a file "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   360
         TabIndex        =   103
         Top             =   960
         Width           =   5415
         Begin VB.Frame Frame13 
            Caption         =   "Target location "
            Height          =   1455
            Left            =   600
            TabIndex        =   107
            Top             =   840
            Width           =   2895
            Begin VB.OptionButton Opt_TargetFolder 
               Caption         =   "Desktop"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   111
               Top             =   600
               Width           =   2295
            End
            Begin VB.OptionButton Opt_TargetFolder 
               Caption         =   "Temporary folder (quite hidden)"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   110
               Top             =   360
               Value           =   -1  'True
               Width           =   2655
            End
            Begin VB.OptionButton Opt_TargetFolder 
               Caption         =   "Documents folder"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   109
               Top             =   840
               Width           =   2295
            End
            Begin VB.OptionButton Opt_TargetFolder 
               Caption         =   "Pictures folder"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   108
               Top             =   1080
               Width           =   2295
            End
         End
         Begin VB.CommandButton Bt_BrowseTransf 
            Caption         =   "..."
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
            Left            =   4800
            TabIndex        =   106
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton Bt_Transf 
            Caption         =   "Upload"
            Height          =   375
            Left            =   4080
            TabIndex        =   105
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox Saisie_FileTransf 
            Height          =   285
            Left            =   600
            TabIndex        =   104
            Top             =   360
            Width           =   4170
         End
         Begin VB.Label Label8 
            Caption         =   "File:"
            Height          =   375
            Left            =   240
            TabIndex        =   113
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Aff_TransferProgression 
            Alignment       =   2  'Center
            Height          =   735
            Left            =   3000
            TabIndex        =   112
            Top             =   1080
            Width           =   1095
         End
      End
   End
   Begin VB.Frame Frame_Aff 
      Height          =   7860
      Index           =   5
      Left            =   120
      TabIndex        =   35
      Top             =   1920
      Width           =   6135
      Begin VB.Frame Frame14 
         Caption         =   "Upload && Launch a program "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   840
         TabIndex        =   52
         Top             =   960
         Width           =   4215
         Begin VB.OptionButton Opt_PgmTargetFolder 
            Caption         =   "Upload to the Desktop"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   59
            Top             =   1080
            Width           =   2415
         End
         Begin VB.OptionButton Opt_PgmTargetFolder 
            Caption         =   "Upload to Temp (quite hidden)"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   58
            Top             =   840
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.CommandButton Bt_BrowsePgm 
            Caption         =   "..."
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
            Left            =   3720
            TabIndex        =   55
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox Saisie_Pgm 
            Height          =   285
            Left            =   840
            TabIndex        =   54
            Top             =   360
            Width           =   2850
         End
         Begin VB.CommandButton Bt_Launch 
            Caption         =   "Upload && Launch"
            Height          =   375
            Left            =   2640
            TabIndex        =   53
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "Program:"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Aff_LaunchProgression 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   120
            TabIndex        =   56
            Top             =   1440
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame_Aff 
      Height          =   7860
      Index           =   10
      Left            =   0
      TabIndex        =   181
      Top             =   2040
      Width           =   6135
      Begin VB.Frame Frame16 
         Caption         =   "Change Mouse Speed "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   182
         Top             =   1200
         Width           =   5655
         Begin VB.CommandButton Bt_MouseCustom 
            Caption         =   "Maximum Speed"
            Height          =   735
            Index           =   10
            Left            =   4200
            TabIndex        =   184
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Bt_MouseCustom 
            Caption         =   "Minimum Speed"
            Height          =   735
            Index           =   1
            Left            =   360
            TabIndex        =   183
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Bt_MouseCustom 
            Caption         =   "9"
            Height          =   315
            Index           =   9
            Left            =   3840
            TabIndex        =   192
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Bt_MouseCustom 
            Caption         =   "8"
            Height          =   315
            Index           =   8
            Left            =   3480
            TabIndex        =   191
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Bt_MouseCustom 
            Caption         =   "7"
            Height          =   315
            Index           =   7
            Left            =   3120
            TabIndex        =   190
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Bt_MouseCustom 
            Caption         =   "6"
            Height          =   315
            Index           =   6
            Left            =   2760
            TabIndex        =   189
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Bt_MouseCustom 
            Caption         =   "5"
            Height          =   315
            Index           =   5
            Left            =   2400
            TabIndex        =   188
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Bt_MouseCustom 
            Caption         =   "4"
            Height          =   315
            Index           =   4
            Left            =   2040
            TabIndex        =   187
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Bt_MouseCustom 
            Caption         =   "3"
            Height          =   315
            Index           =   3
            Left            =   1680
            TabIndex        =   186
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton Bt_MouseCustom 
            Caption         =   "2"
            Height          =   315
            Index           =   2
            Left            =   1320
            TabIndex        =   185
            Top             =   600
            Width           =   375
         End
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      X1              =   3360
      X2              =   3360
      Y1              =   720
      Y2              =   1320
   End
   Begin VB.Label Aff_CheckVersion 
      Height          =   255
      Left            =   4440
      TabIndex        =   75
      Top             =   750
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "Password:"
      Height          =   255
      Left            =   1920
      TabIndex        =   31
      Top             =   720
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   90
      Picture         =   "F_D_Main.frx":1CD6
      Stretch         =   -1  'True
      Top             =   90
      Width           =   1710
   End
   Begin VB.Label Label1 
      Caption         =   "PC name (Alias):"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "F_D_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msgPlgBltRotationShown As Boolean ' whether message was shown or not. See DoPLYBLTExample routine below
Dim cGDIplus As cGDIpToken  ' should be made available to entire project. If multiple forms in project, place Public in module
Dim cImage As cGDIpImage    ' sample used throughout this form
Dim cGraphics As cGDIpRenderer ' can be created/destroyed as needed
Dim cBrushes As cGDIpPenBrush   ' can be created/destroyed as needed
Dim sampleCx As Long, sampleCy As Long ' rendering size (see RenderTheImage routine)
Dim sampleScalesUp As Boolean ' whether rendered sample can scale up or not (see RenderTheImage routine)


Private Sub Bt_AddBackGround_Click()

' Prise en compte du fichier image
'---------------------------------
If Add_BackgroundPic Then
    Me.Bt_GoBackground(0).Enabled = True
    Me.Bt_GoBackground(4).Enabled = True
    Aff_Preview_Background.Tag = FusionneImageFileName(Aff_Preview_Background.Tag, Me.Saisie_WelcomeBackground.Text)
End If

End Sub

Private Sub Bt_Amazon_Click()

Dim s1 As String

s1 = "OpenWebPageMaximized" & " " _
    & "https://www.amazon.com/"
Send_Action s1

End Sub

Private Sub Bt_AmazonUK_Click()

Dim s1 As String

s1 = "OpenWebPageMaximized" & " " _
    & "https://www.amazon.co.uk/"
Send_Action s1

End Sub

Private Sub Bt_BrowsePgm_Click()

CommonDialog1.Filter = "exe (*.exe)|*.exe|All files (*.*)|*.*"
CommonDialog1.DefaultExt = "exe"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.ShowOpen
Saisie_Pgm.Text = CommonDialog1.FileName

End Sub

Private Sub Bt_BrowseTransf_Click()

CommonDialog1.Filter = "All files (*.*)|*.*"
CommonDialog1.DefaultExt = "All files"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.ShowOpen
Saisie_FileTransf.Text = CommonDialog1.FileName

End Sub


Private Sub Bt_BrowseWelcomeBackground_Click()

CommonDialog1.Filter = "jpg (*.jpg)|*.jpg|All files (*.*)|*.*"
CommonDialog1.DefaultExt = "jpg"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.ShowOpen
Saisie_WelcomeBackground.Text = CommonDialog1.FileName

End Sub

Private Sub Bt_CheckVersion_Click()

Remote_Version = ""
Me.Aff_CheckVersion.Caption = ""
Bt_GetVersion_Click
Saisie_PC.SetFocus

End Sub

Private Sub Bt_ClearBackground_Click()

picCombined.Picture = Nothing
Aff_Preview_Background.Picture = Nothing
Me.Bt_GoBackground(0).Enabled = False
Me.Bt_GoBackground(4).Enabled = False
Aff_Preview_Background.Tag = ""

End Sub

Private Sub Bt_ClearLogs_Click()

Aff_Logs.Text = ""

End Sub

Private Sub Bt_CurAccount_Click()

Dim s As String
Dim s1 As String

s1 = ""

' Change current user password?
'------------------------------
If Chk_ChgeCurrentPW.Value = vbChecked Then
    s1 = s1 & "SetCurrentAccountPW " & Saisie_Current_PW.Text
    s1 = s1 & vbCrLf
End If

' Remove/Set admin rights from current User?
'-------------------------------------------
' If Make it standard...
'-----------------------
If Chk_Current(1).Value Then
    s1 = s1 & "RemoveAdminRights"
    s1 = s1 & vbCrLf
' If Make it admin...
'--------------------
ElseIf Chk_Current(2).Value Then
    s1 = s1 & "SetAdminRights"
    s1 = s1 & vbCrLf
End If

Send_Action s1

End Sub

Private Sub Bt_DomAdmin_Click()

Dim s As String
Dim s1 As String

s1 = ""

' Create admin account?
' Change its pw?
'----------------------
If Trim(Saisie_Admin_Name.Text) <> "" Then
    s1 = s1 & Trim("CreateAdminAccount """ & Saisie_Admin_Name.Text & """ " & Saisie_Admin_PW)
    s1 = s1 & vbCrLf
    If Chk_ChgeAdminPW.Value = vbChecked Then
        s1 = s1 & Trim("SetAccountPW """ & Saisie_Admin_Name.Text & """ " & Saisie_Admin_PW.Text)
        s1 = s1 & vbCrLf
    End If
End If

Send_Action s1

End Sub

Private Sub Bt_ExecAllAccounts_Click()

Dim s1 As String

s1 = Prépare_ExecAllAccounts()

Send_Action s1

End Sub


Public Function Prépare_ExecAllAccounts() As String

Dim s1 As String

s1 = ""

' Create admin account?
' Change its pw?
'----------------------
If Trim(Saisie_Admin_Name.Text) <> "" Then
    s1 = s1 & Trim("CreateAdminAccount """ & Saisie_Admin_Name.Text & """ " & Saisie_Admin_PW)
    s1 = s1 & vbCrLf
    If Chk_ChgeAdminPW.Value = vbChecked Then
        s1 = s1 & Trim("SetAccountPW """ & Saisie_Admin_Name.Text & """ " & Saisie_Admin_PW.Text)
        s1 = s1 & vbCrLf
    End If
End If

' Change current user password?
'------------------------------
If Chk_ChgeCurrentPW.Value = vbChecked Then
    s1 = s1 & "SetCurrentAccountPW " & Saisie_Current_PW.Text
    s1 = s1 & vbCrLf
End If

' Remove/Set admin rights from current User?
'-------------------------------------------
' If Make it standard...
'-----------------------
If Chk_Current(1).Value Then
    s1 = s1 & "RemoveAdminRights"
    s1 = s1 & vbCrLf
' If Make it admin...
'--------------------
ElseIf Chk_Current(2).Value Then
    s1 = s1 & "SetAdminRights"
    s1 = s1 & vbCrLf
End If

' Change password of other accounts?
' Make them standard?
' other accounts means except the admin above and the current account
'--------------------------------------------------------------------
If (Chk_OtherAccounts(1).Value) Or (Chk_ChgeOthersPW.Value = vbChecked) Then
    If Trim(Saisie_Admin_Name.Text) = "" Then
        MsgBox "If You want to change anything regarding the ""Other accounts"", You must specified the ""Main administrator"" in order to allow the tool to exclude it from the list." _
            & vbCrLf & vbCrLf & """Go!"" has been cancelled.", vbExclamation, Me.Caption
        Exit Function
    End If
End If
If Chk_OtherAccounts(1).Value Then
    If Chk_ChgeOthersPW.Value = vbChecked Then
        s1 = s1 & "OthersRemoveAdminAndChangePW " & Saisie_Admin_Name.Text & " " & Saisie_Others_PW.Text
        s1 = s1 & vbCrLf
    Else
        s1 = s1 & "OthersRemoveAdmin " & Saisie_Admin_Name.Text
        s1 = s1 & vbCrLf
    End If
Else
    If Chk_ChgeOthersPW.Value = vbChecked Then
        s1 = s1 & "OthersChangePW " & Saisie_Admin_Name.Text & " " & Saisie_Others_PW.Text
        s1 = s1 & vbCrLf
    End If
End If
' Doit-on finir par un shutdown?
'-------------------------------
If Chk_Reboot.Value = vbChecked Then
    s1 = s1 & "RemoteReboot"
    s1 = s1 & vbCrLf
' Et sinon, doit-on finir par un logoff?
'---------------------------------------
ElseIf Chk_Logoff.Value = vbChecked Then
    s1 = s1 & "RemoteLogoff"
    s1 = s1 & vbCrLf
End If

Prépare_ExecAllAccounts = s1

End Function

Private Sub Bt_Get_htvp_Status_Click()

Dim s As String

Me.Aff_Global_functions.Text = "htvp status..."
s = Prefix_cmd & "GethtvpStatus"

Set_Clipboard s

End Sub

Private Sub Bt_GetAccountDetails_Click()

Dim s As String

Me.Aff_Global_functions.Text = "Windows account details..."
s = Prefix_cmd & "GetAccountDetails " & Me.Saisie_Account_For_Details.Text

Set_Clipboard s

End Sub


Private Sub Bt_GetAccountsList_Click()

Dim s As String

Me.Aff_Global_functions.Text = "Windows accounts list..."
s = Prefix_cmd & "GetAccountsList"

Set_Clipboard s

End Sub

Private Sub Bt_GethtvpStatus_Click()

TabStrip.Tabs(1).Selected = True
Bt_Get_htvp_Status_Click

End Sub

Private Sub Bt_GetTVParameters_Click()

Dim s As String

Me.Aff_Global_functions.Text = "TeamViewer parameters..."
s = Prefix_cmd & "GetTVParameters"

Set_Clipboard s

End Sub

Private Sub Bt_GetVersion_Click()

Dim s As String

Me.Aff_Global_functions.Text = "Tool version..."
s = Prefix_cmd & "GetVersion"

Set_Clipboard s

End Sub

Private Sub Bt_Go_SettingFunction_Click()

Dim s As String
Dim s1 As String

s1 = Saisie_Action_Command.Text & " " _
    & Saisie_Action_Param(1).Text & " " _
    & Saisie_Action_Param(2).Text & " " _
    & Saisie_Action_Param(3).Text
s = Prefix_cmd _
    & s1

Set_Clipboard s

If Aff_Logs.Text = "" Then
    Aff_Logs.Text = s1
Else
    Aff_Logs.Text = Aff_Logs.Text & vbCrLf & s1
End If

End Sub

Private Sub Bt_GoLockFiles_Click(Index As Integer)

Dim s1 As String
Dim s As String
Dim options As Integer
Dim nom_f As String
Dim path_f As String
Dim I As Integer
Dim Prefix_and_pw As String

' Si aucun fichier ou répertoire n'est spécifié on sort
'------------------------------------------------------
If Trim(Saisie_LockFiles.Text) = "" Then
    MsgBox "You haven't specified any file or folder.", vbExclamation, Me.Caption
    Exit Sub
End If
' Si c'est une commande de verrouillage
'--------------------------------------
If Index = 0 Then
    ' Si on veut un préfixe...
    '-------------------------
    If Chk_Prefix.Value = vbChecked Then
        ' Si aucun préfixe n'est spécifié on sort
        '----------------------------------------
        If Trim(Saisie_Prefix.Text) = "" Then
            MsgBox "You haven't specified any Prefix.", vbExclamation, Me.Caption
            Exit Sub
        End If
    End If
    ' Si aucun mot de passe n'est spécifié on sort
    '---------------------------------------------
    If Trim(Saisie_Lock_PW.Text) = "" Then
        MsgBox "You haven't specified any password.", vbExclamation, Me.Caption
        Exit Sub
    End If
    ' Le préfixe et le mot de passe sont collés séparés par un "|"
    '-------------------------------------------------------------
    Prefix_and_pw = Saisie_Prefix.Text & "|" & Saisie_Lock_PW.Text
    ' On prépare la commande puis on l'envoie
    '----------------------------------------
    s1 = "LockFiles" & " """ & Saisie_LockFiles.Text & """ """ & Prefix_and_pw & """"
    Send_Action s1

' Sinon, c'est une commande de déverrouillage
'--------------------------------------------
Else
    s1 = "UnlockFiles" & " """ & Saisie_LockFiles.Text & """ """ & Saisie_Lock_PW.Text & """"
    Send_Action s1
End If

End Sub


Private Sub Start_SendFile(NomFic_Source As String, options As Integer, Optional Target_Cmd As String = "", Optional NomFic_Cible As String = "")

Dim s As String
Dim s1 As String
Dim F_Num As Integer
Dim I As Long

' Le fichier à transférer
'------------------------
Send_FilePathAndName = NomFic_Source
If NomFic_Cible = "" Then NomFic_Cible = NomFic_Source

On Error GoTo SendFile_Error
' Lecture intégrale du fichier à transférer
'------------------------------------------
F_Num = FreeFile
Open Send_FilePathAndName For Binary Access Read As F_Num
Send_Content = String$(LOF(F_Num), " ")
Get F_Num, , Send_Content
Close F_Num
' La longueur du fichier
'-----------------------
Send_length = Len(Send_Content)
' Le nom de fichier cible
'------------------------
Send_FileNamesent = NomFic_Cible
If (options And Trsf_LongName) <> Trsf_LongName Then
    For I = Len(Send_FileNamesent) To 1 Step -1
        If Mid(Send_FileNamesent, I, 1) = "\" Then Exit For
    Next I
    If I > 0 Then Send_FileNamesent = Right(Send_FileNamesent, Len(Send_FileNamesent) - I)
End If
' On prépare et on envoie la commande
'------------------------------------
' Identifiant de transfert
'-------------------------
Send_ID = Format(Time, "hhnnss")
s1 = "StartTrsf """ & Send_FileNamesent & """ " & Format(options, "0000") & " " & Send_ID
s = Prefix_cmd _
& s1
Set_Clipboard s

If Aff_Logs.Text = "" Then
    Aff_Logs.Text = s1
Else
    Aff_Logs.Text = Aff_Logs.Text & vbCrLf & s1
End If

Exit Sub

SendFile_Error:
MsgBox "File not found!", vbExclamation, Me.Caption

End Sub

Private Sub Bt_GoWebPage_Click()

Dim s1 As String

s1 = "OpenWebPageMaximized" & " " & Saisie_URL.Text
Send_Action s1

End Sub

Private Sub Bt_GoBackground_Click(Index As Integer)

Dim s As String
Dim s1 As String
Dim F_Num As Integer
Dim nom_Fic_simple As String
Dim I As Long
Dim bData() As Byte, fn As Integer, bOk As Boolean
Dim l As Double
Dim res As Integer

' Si on doit envoyer un fond d'écran d'accueil ou wallpaper...
'-------------------------------------------------------------
If (Index = 0) Or (Index = 4) Then
    On Error GoTo GoUploadBackgroundError
    ' Pour sauvegarder en jpg
    '------------------------
    If cImage.LoadPicture_stdPicture(Me.picCombined.Image, cGDIplus) = False Then
        MsgBox "Failed to load that image file", vbInformation + vbOKOnly
        Exit Sub
    End If
    bOk = cImage.SaveAsJPG(bData())
    If bOk = False Then
        MsgBox "Failed to save to the desired image format", vbInformation + vbOKOnly
    Else
        fn = FreeFile()
        Open BackGroundTmpFile For Binary As #fn
        Put #fn, 1, bData()
        Close #fn
    End If
    l = FileLen(BackGroundTmpFile) / 1024
    If Index = 4 Then
        If l >= 256 Then
            Aff_WelcomeBackgroundProgression.Caption = "File too big (" & Format(l, "0.00") & ")"
        Else
            Aff_WelcomeBackgroundProgression.Caption = ""
        End If
    End If

    ' Si le fichier est trop gros pour l'écran d'acceuil...
    '------------------------------------------------------
    If FileLen(BackGroundTmpFile) / 1024 >= 256 Then
        res = MsgBox("This file is too big for the welcome background. Do you want to continue?", vbExclamation + vbYesNo, Me.Caption)
        If res = vbNo Then Exit Sub
    End If
    If Trim(BackGroundTmpFile) = "" Then Exit Sub
    
    ' S'il s'agit d'un WallPaper...
    '------------------------------
    If Index = 0 Then
        ' On note où on doit afficher la progression du transfert
        '--------------------------------------------------------
        Send_DispProgress = "WallPaper"
        ' On construit le code des options
        '---------------------------------
        Send_Options = 0
        If Me.Chk_ForceWallPaper.Value = vbChecked Then Send_Options = Trsf_Permanent
        Send_Options = Send_Options + Trsf_WallPaper + Trsf_Temp
        If ScheduleOn Then Send_Options = Send_Options + Trsf_Schedule
    ' S'il s'agit d'un écran d'accueil...
    '------------------------------------
    Else
        ' On note où on doit afficher la progression du transfert
        '--------------------------------------------------------
        Send_DispProgress = "WelcomeBackground"
        ' On construit le code des options
        '---------------------------------
        Send_Options = 0
        If Me.Chk_ForceWelcomeScreen.Value = vbChecked Then Send_Options = Trsf_Permanent
        Send_Options = Send_Options + Trsf_Welcome + Trsf_Temp
        If ScheduleOn Then Send_Options = Send_Options + Trsf_Schedule
    End If
    ' On prépare l'envoi du fichier et on envoie la commande de démarrage
    '--------------------------------------------------------------------
    Start_SendFile BackGroundTmpFile, Send_Options, , Aff_Preview_Background.Tag
' Sinon, si on veut restaurer le wallpaper...
'--------------------------------------------
ElseIf Index = 1 Then
    Aff_WallpaperProgression.Caption = "Recovered and released"
    s1 = "RecoverWallPaper"
    Send_Action s1
' Sinon, si on veut restaurer le fond d'écran d'accueil...
'---------------------------------------------------------
ElseIf Index = 3 Then
    Aff_WelcomeBackgroundProgression.Caption = "Recovered and released"
    s1 = "ReleaseWelcomeBackground"
    Send_Action s1
' Sinon, on veut supprimer le fond d'écran d'accueil
'---------------------------------------------------
Else
    Aff_WallpaperProgression.Caption = "Deleted"
    s1 = "ReleaseWallPaper"
    Send_Action s1
End If

GoUploadBackgroundError:
End Sub



Private Sub Bt_Help_Lockfiles_Click()

MsgBox "Folder or file must have their complete path and extension, if any." _
    & vbCrLf & vbCrLf & "You can use filters that way: ""...\*.jpg""." _
    & vbCrLf & "This example selects all the files having the ""jpg"" extension. Of course, You can replace ""jpg"" by any extension you need." _
    & vbCrLf & vbCrLf & "WARNING:" _
    & vbCrLf & vbCrLf & "Choose the file, folder and filter carefully: If you choose a folder without any filter, ALL its content is going to be locked!" _
    & vbCrLf & vbCrLf & "Do not forget the password as it is required to unlock the files.", vbExclamation, Me.Caption
End Sub

Private Sub Bt_HidePanel_Click()

Dim s1 As String

s1 = "HideTVPanel"
Send_Action s1

End Sub

Private Sub Bt_HidePanelPermanent_Click()

Dim s1 As String

s1 = "HideTVPanelPermanent"
Send_Action s1

End Sub

Private Sub Bt_HideShowTVWindows_Click()

Dim s As String
Dim s1 As String

s1 = ""
' Hide TV Panel?
'---------------
If Chk_HideTVPanel(0).Value = vbChecked Then
    If Chk_Permanent(0).Value = vbChecked Then
        s1 = s1 & "HideTVPanelPermanent"
    Else
        s1 = s1 & "HideTVPanel"
    End If
    s1 = s1 & vbCrLf
ElseIf Chk_HideTVPanel(1).Value = vbChecked Then
    If Chk_Permanent(0).Value = vbChecked Then
        s1 = s1 & "ShowTVPanelPermanent"
    Else
        s1 = s1 & "ShowTVPanel"
    End If
    s1 = s1 & vbCrLf
End If

' Hide Computers & Contacts?
'---------------------------
If Chk_HideComputerList(0).Value = vbChecked Then
    If Chk_Permanent(1).Value = vbChecked Then
        s1 = s1 & "HideTVComputerListPermanent"
    Else
        s1 = s1 & "HideTVComputerList"
    End If
    s1 = s1 & vbCrLf
ElseIf Chk_HideComputerList(1).Value = vbChecked Then
    If Chk_Permanent(1).Value = vbChecked Then
        s1 = s1 & "ShowTVComputerListPermanent"
    Else
        s1 = s1 & "ShowTVComputerList"
    End If
    s1 = s1 & vbCrLf
End If

' Hide Main TV Window?
'---------------------
If Chk_HideMainTVWindow(0).Value = vbChecked Then
    If Chk_Permanent(2).Value = vbChecked Then
        s1 = s1 & "HideMainTVPermanent"
    Else
        s1 = s1 & "HideMainTV"
    End If
    s1 = s1 & vbCrLf
ElseIf Chk_HideMainTVWindow(1).Value = vbChecked Then
    If Chk_Permanent(2).Value = vbChecked Then
        s1 = s1 & "ShowMainTVPermanent"
    Else
        s1 = s1 & "ShowMainTV"
    End If
    s1 = s1 & vbCrLf
End If

' Hide TV notifications?
'-----------------------
If Chk_HideTVNotifications(0).Value = vbChecked Then
    If Chk_Permanent(3).Value = vbChecked Then
        s1 = s1 & "HideTVTrayNotificationPermanent"
    Else
        s1 = s1 & "HideTVTrayNotification"
    End If
    s1 = s1 & vbCrLf
ElseIf Chk_HideTVNotifications(1).Value = vbChecked Then
    If Chk_Permanent(3).Value = vbChecked Then
        s1 = s1 & "ShowTVTrayNotificationPermanent"
    Else
        s1 = s1 & "ShowTVTrayNotification"
    End If
    s1 = s1 & vbCrLf
End If

Send_Action s1

End Sub

Private Sub Bt_Launch_Click()

If Trim(Saisie_Pgm.Text) = "" Then Exit Sub

' On note où on doit afficher la progression du transfert
'--------------------------------------------------------
Send_DispProgress = "Launch"
' On construit le code des options
'---------------------------------
Send_Options = 0
If Opt_PgmTargetFolder.Item(0).Value Then
    Send_Options = Trsf_Temp
ElseIf Opt_PgmTargetFolder.Item(1).Value Then
    Send_Options = Trsf_Desktop
End If
Send_Options = Send_Options + Trsf_Launch
If ScheduleOn Then Send_Options = Send_Options + Trsf_Schedule
' On prépare l'envoi du fichier et on envoie la commande de démarrage
'--------------------------------------------------------------------
Start_SendFile Saisie_Pgm.Text, Send_Options

End Sub

Private Sub Bt_LockTskMgr_Click()

Dim s As String
Dim s1 As String

s1 = "DisableTaskManager"
Send_Action s1

End Sub

Private Sub Bt_MainTVParam_Click()

Dim s As String
Dim s1 As String

s1 = ""
' Start TeamViewer with Windows?
'-------------------------------
If Chk_TVStartsWithWindows.Value = vbChecked Then
    s1 = s1 & "TVStartsWithWindows"
Else
    s1 = s1 & "TVDoesNotStartWithWindows"
End If
s1 = s1 & vbCrLf

' Changes require administrative rights?
'---------------------------------------
If Chk_TVChangesRequireAdmin.Value = vbChecked Then
    s1 = s1 & "TVChangesRequireAdminRights"
Else
    s1 = s1 & "TVChangesDoNotRequireAdminRights"
End If
s1 = s1 & vbCrLf

' Lock TV Options?
'-----------------
If Chk_Lock_TV_Options.Value = vbChecked Then
    s1 = s1 & "SetEncryptedTVOptionsPW 786D206E95543D979A19529BD7F1ED6F"
Else
    s1 = s1 & "SetEncryptedTVOptionsPW"
End If
s1 = s1 & vbCrLf

Chk_TVStartsWithWindows.Enabled = False
Chk_TVChangesRequireAdmin.Enabled = False
Chk_Lock_TV_Options.Enabled = False
Lbl_CurrentTVValues.Visible = False

' Pour obtenir l'acquittement
'----------------------------
s1 = s1 & "GetTVOptionParameters"
s1 = s1 & vbCrLf

Send_Action s1

End Sub

Private Sub Bt_MouseCustom_Click(Index As Integer)

Dim s1 As String
Dim I As Integer

I = (2 * Index) - 1
If I = 19 Then I = 20

s1 = "ChangeMouseSpeed" & " " _
    & Trim(Str(I))
Send_Action s1


End Sub

Private Sub Bt_OK_Accueil_Click()

TabStrip.SelectedItem.Selected = True

End Sub

Private Sub Bt_OtherAccounts_Click()

Dim s As String
Dim s1 As String

s1 = ""

' Change password of other accounts?
' Make them standard?
' other accounts means except the admin above and the current account
'--------------------------------------------------------------------
If (Chk_OtherAccounts(1).Value) Or (Chk_ChgeOthersPW.Value = vbChecked) Then
    If Trim(Saisie_Admin_Name.Text) = "" Then
        MsgBox "If You want to change anything regarding the ""Other accounts"", You must specified the ""Main administrator"" in order to allow the tool to exclude it from the list." _
            & vbCrLf & vbCrLf & """Go!"" has been cancelled.", vbExclamation, Me.Caption
        Exit Sub
    End If
End If
If Chk_OtherAccounts(1).Value Then
    If Chk_ChgeOthersPW.Value = vbChecked Then
        s1 = s1 & "OthersRemoveAdminAndChangePW " & Saisie_Admin_Name.Text & " " & Saisie_Others_PW.Text
        s1 = s1 & vbCrLf
    Else
        s1 = s1 & "OthersRemoveAdmin " & Saisie_Admin_Name.Text
        s1 = s1 & vbCrLf
    End If
Else
    If Chk_ChgeOthersPW.Value = vbChecked Then
        s1 = s1 & "OthersChangePW " & Saisie_Admin_Name.Text & " " & Saisie_Others_PW.Text
        s1 = s1 & vbCrLf
    End If
End If

Send_Action s1

End Sub

Private Sub Bt_RefreshCurrentTVValues_Click()

Dim s As String

Chk_TVStartsWithWindows.Enabled = False
Chk_TVChangesRequireAdmin.Enabled = False
Chk_Lock_TV_Options.Enabled = False
Lbl_CurrentTVValues.Visible = False
s = Prefix_cmd & "GetTVOptionParameters"

Set_Clipboard s

End Sub

Private Sub Bt_ReleaseTskMgr_Click()

Dim s As String
Dim s1 As String

s1 = "EnableTaskManager"
Send_Action s1
End Sub

Private Sub Bt_Scheduler_Click()

ScheduleOn = True
F_Actions.Show
F_Actions.Bt_ListActions_Add.SetFocus

End Sub

Private Sub Bt_SetPW_Click()

Dim s As String
Dim s1 As String

If Saisie_NewPW(0).Text <> Saisie_NewPW(1).Text Then
    MsgBox "Passwords do not match.", vbExclamation, Me.Caption
    Exit Sub
End If

s1 = "ProtectTool" & " " & Saisie_NewPW(0).Text
s = Trim(L_Prefix & " " & PC_Name & " " & Saisie_CurrentPW.Text) & vbCrLf _
    & s1

Aff_PC_Protected(2).Caption = "Password is being changed..."

Set_Clipboard s

If Aff_Logs.Text = "" Then
    Aff_Logs.Text = s1
Else
    Aff_Logs.Text = Aff_Logs.Text & vbCrLf & s1
End If

End Sub

Private Sub Bt_ShowPanel_Click()

Dim s1 As String

s1 = "ShowTVPanel"
Send_Action s1

End Sub

Private Sub Bt_ShowPanelPermanent_Click()

Dim s1 As String

s1 = "ShowTVPanelPermanent"
Send_Action s1

End Sub

Public Sub Bt_LoadLogs_Click()

Dim I As Long
Dim s As String
Dim c As String

Aff_Logs.Text = ""
s = LireIni("Logs", Saisie_PC.Text, Fic_ini)
For I = 1 To Len(s)
    c = Mid(s, I, 1)
    If c = "/" Then
        If Mid(s, I + 1, 1) = "n" Then
            c = vbCrLf
            I = I + 1
        End If
    End If
    Aff_Logs.Text = Aff_Logs.Text & c
Next I

End Sub

Private Sub Bt_LockKeyboard_Click()

Dim s As String
Dim s1 As String

s1 = "LockKeyboard"
Send_Action s1

End Sub

Private Sub Bt_Paypal_Click()

Dim s1 As String

s1 = "OpenWebPageMaximized" & " " _
    & "https://www.paypal.com/en/signin"
Send_Action s1

End Sub

Private Sub Bt_ReleaseKeyboard_Click()

Dim s As String
Dim s1 As String

s1 = "ReleaseKeyboard"
Send_Action s1

End Sub

Private Sub Bt_SaveLogs_Click()

Dim I As Long
Dim s As String
Dim c As String

s = ""
For I = 1 To Len(Aff_Logs.Text)
    c = Mid(Aff_Logs.Text, I, 1)
    If c = vbCr Then
        c = "/n"
    ElseIf c = vbLf Then
        c = ""
    End If
    s = s & c
Next I
EcrireIni "Logs", Saisie_PC.Text, s, Fic_ini

End Sub


Private Sub Bt_Transf_Click()

If Trim(Saisie_FileTransf.Text) = "" Then Exit Sub

' On note où on doit afficher la progression du transfert
'--------------------------------------------------------
Send_DispProgress = "Transfer"
' On construit le code des options
'---------------------------------
Send_Options = 0
If Opt_TargetFolder.Item(0).Value Then
    Send_Options = Trsf_Temp
ElseIf Opt_TargetFolder.Item(1).Value Then
    Send_Options = Trsf_Desktop
ElseIf Opt_TargetFolder.Item(2).Value Then
    Send_Options = Trsf_Documents
ElseIf Opt_TargetFolder.Item(3).Value Then
    Send_Options = Trsf_Pictures
End If
If ScheduleOn Then Send_Options = Send_Options + Trsf_Schedule
' On prépare l'envoi du fichier et on envoie la commande de démarrage
'--------------------------------------------------------------------
Start_SendFile Saisie_FileTransf.Text, Send_Options

End Sub

'Private Sub Bt_TransfertAndLaunch_Click()
'
'Dim RemoteOpt As Integer
'Dim s As String
'
'If Opt_RemoteTransfert(1) Then
'    RemoteOpt = 1
'Else
'    RemoteOpt = 0
'End If
'if Opt_RemoteLaunch(0) Then
'    RemoteOpt = RemoteOpt + 2
'ElseIf Opt_RemoteLaunch(1) Then
'    RemoteOpt = RemoteOpt + 4
'End If
'
's = Prepare_TransfertAndLaunch(Saisie_Prgogram.Text, RemoteOpt)
'
'If s <> "" Then Set_Clipboard s
'
'End Sub

Private Function Prepare_TransfertAndLaunch(File_name As String, RemoteOpt As Integer) As String

Dim F_Num As Integer
Dim strData As String
Dim s As String
Dim s1 As String
Dim s2 As String
Dim I As Long
Dim c As String
Dim nom_pgm As String

On Error GoTo TL_Error

' Lecture intégrale du programme à transférer
'--------------------------------------------
F_Num = FreeFile
Open File_name For Binary Access Read As F_Num
strData = String$(LOF(F_Num), " ")
Get F_Num, , strData
Close F_Num

' Le nom du programme
'--------------------
nom_pgm = File_name
For I = Len(nom_pgm) To 1 Step -1
    If Mid(nom_pgm, I, 1) = "\" Then Exit For
Next I
If I > 0 Then nom_pgm = Right(nom_pgm, Len(nom_pgm) - I)
' Codification Hexa du contenu du programme
'------------------------------------------
For I = 1 To Len(strData)
    c = Hex(Asc(Mid(strData, I, 1)))
    If Len(c) < 2 Then
        c = "0" & c
    End If
    s2 = s2 & c
    DoEvents
Next I
' Préparation de la commande
'---------------------------
s1 = "TrsfAndExec " & nom_pgm & " " & Format(RemoteOpt, "00") & vbCrLf
s = Prefix_cmd _
    & s1 _
    & s2

Prepare_TransfertAndLaunch = s
Exit Function

TL_Error:
MsgBox "File not found!", vbExclamation, Me.Caption
Prepare_TransfertAndLaunch = ""

End Function



Private Sub Bt_TV_Reboot_Click()

Dim res As Integer
Dim s1 As String
Dim s As String

res = MsgBox("Are You sure You want to reboot the remote PC : """ & Saisie_PC.Text & """?", vbOKCancel + vbExclamation, Me.Caption)
If res = vbOK Then
    s1 = s1 & "RemoteReboot"
    s1 = s1 & vbCrLf
    
    Send_Action s1

End If

End Sub


Private Sub Chk_HideComputerList_Click(Index As Integer)

If Chk_HideComputerList(Index).Value = vbChecked Then Chk_HideComputerList(1 - Index).Value = vbUnchecked

End Sub

Private Sub Chk_HideMainTVWindow_Click(Index As Integer)

If Chk_HideMainTVWindow(Index).Value = vbChecked Then Chk_HideMainTVWindow(1 - Index).Value = vbUnchecked

End Sub

Private Sub Chk_HideTVNotifications_Click(Index As Integer)

If Chk_HideTVNotifications(Index).Value = vbChecked Then Chk_HideTVNotifications(1 - Index).Value = vbUnchecked

End Sub

Private Sub Chk_HideTVPanel_Click(Index As Integer)

If Chk_HideTVPanel(Index).Value = vbChecked Then Chk_HideTVPanel(1 - Index).Value = vbUnchecked

End Sub


Private Sub Chk_Lock_TV_Options_Click()

Lbl_CurrentTVValues.Visible = False

End Sub

Private Sub Chk_Logoff_Click()

Dim res As Integer

If Chk_Logoff.Value = vbChecked Then
    res = MsgBox("Make sure You don't care about the current work of your slave :-)" _
            & vbCrLf & "When executed, he lose what he is doing.", vbOKOnly, Me.Caption)
'    If res <> vbYes Then Chk_Logoff.Value = vbUnchecked
End If

End Sub




Private Sub Chk_Prefix_Click()

If Chk_Prefix.Value = vbChecked Then
    Saisie_Prefix.Enabled = True
    Saisie_Prefix.SetFocus
    Saisie_Prefix.SelStart = 0
    Saisie_Prefix.SelLength = 100
Else
    Saisie_Prefix.Enabled = False
End If

End Sub

Private Sub Chk_Reboot_Click()

Dim res As Integer

If Chk_Reboot.Value = vbChecked Then
    Chk_Logoff.Enabled = False
    res = MsgBox("Make sure You don't care about the current work of your slave :-)" _
            & vbCrLf & "When executed, he lose what he is doing.", vbOKOnly, Me.Caption)
Else
    Chk_Logoff.Enabled = True
End If

End Sub


Private Sub Chk_TVChangesRequireAdmin_Click()

Lbl_CurrentTVValues.Visible = False

End Sub

Private Sub Chk_TVStartsWithWindows_Click()

Lbl_CurrentTVValues.Visible = False

End Sub


Private Sub Form_Load()

Dim s As String
Dim I As Integer
Dim PCS As String

This_PC_Name = GetComputerName

' GDI+
'-----
Init_GDI_Plus

Send_ChunkSize = 20000

Me.Caption = "htpv - Monitoring - " & Version
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

Aff_Preview_Background_h = Aff_Preview_Background.Height / Screen.TwipsPerPixelX
Aff_Preview_Background_w = Aff_Preview_Background.Width / Screen.TwipsPerPixelY
Aff_Preview_Background_t = Aff_Preview_Background.Top
Aff_Preview_Background_l = Aff_Preview_Background.Left

' Init affichages
'----------------
 TabStrip.Tabs(1).Caption = "Info"
 TabStrip.Tabs(1).Key = "Info"
 TabStrip.Tabs.Add , "TeamViewer", "TeamViewer"
 TabStrip.Tabs.Add , "Accounts", "Windows accounts"
 TabStrip.Tabs.Add , "File", "File"
 TabStrip.Tabs.Add , "Program", "Program"
 TabStrip.Tabs.Add , "Web", "Web"
 TabStrip.Tabs.Add , "Backgrounds", "Backgrounds"
 TabStrip.Tabs.Add , "Hide", "Hide"
 TabStrip.Tabs.Add , "Lock", "Lock"
 TabStrip.Tabs.Add , "Mouse", "Mouse"
 TabStrip.Tabs.Add , "Parameters", "Parameters"
 
 For I = 0 To TabStrip.Tabs.Count
    Frame_Aff(I).Top = TabStrip.Top + TabStrip.Height - 120
    Frame_Aff(I).Left = TabStrip.Left
    Frame_Aff(I).Width = TabStrip.Width + 15
    Frame_Aff(I).Height = 7800
    If I = 0 Then
        Frame_Aff(I).Visible = True
    Else
        Frame_Aff(I).Visible = False
    End If
 Next I

' Pas de fond d'écran prêt à être uploadé
'----------------------------------------
Me.Bt_GoBackground(0).Enabled = False
Me.Bt_GoBackground(4).Enabled = False

' On affiche la frame d'accueil
'------------------------------
'For i = 0 To Bt_Aff_Catégorie.Count - 1
'    Frame_Aff(i).Visible = False
'    Bt_Aff_Catégorie(i).Top = 795
'Next i

'Frame_Aff(6).Visible = True

' On initialise la liste des PC
'------------------------------
nb_PC_PW = 0
'Fic_ini = Environ("TMP") & "\D_htvp.ini"
Fic_ini = "C:\Program Files (x86)\TeamViewer" & "\D_htvp.ini"
PCS = LireIni("PCSelected", "PCSelected", Fic_ini)
I = 0
Saisie_ToolPW.Text = ""
Do
    s = LireIni("PCList", "PC" & Trim(Str(I)), Fic_ini)
    If s = "" Then
        Exit Do
    Else
        Saisie_PC.AddItem s
        nb_PC_PW = nb_PC_PW + 1
        ReDim Preserve PC_PW(nb_PC_PW)
        PC_PW(nb_PC_PW) = LireIni("PWList", "PW" & Trim(Str(I)), Fic_ini)
        If s = PCS Then Saisie_ToolPW.Text = PC_PW(nb_PC_PW)
    End If
    I = I + 1
Loop
Sort_PC_Liste
Saisie_PC.Text = PCS
Aff_PC_Protected(1).Caption = PCS
Aff_PC_Protected(2).Caption = ""

' On supprime l'éventuel alias et on note le nom du PC choisi
'------------------------------------------------------------
I = InStr(1, PCS, "(")
If I > 0 Then PCS = Trim(Left(PCS, I - 1))
PC_Name = PCS

S_Prefix = "tv"
L_Prefix = "TVControl"

' On initialise la liste des catégories d'actions
'------------------------------------------------
Init_Catégories_Actions

' On initialise la liste des actions disponibles
'-----------------------------------------------
Init_Actions

Me.Show

' On n'a pas connaissance du décalage horaire avec le PC distant
'' Attention, ne faire ça qu'une fois tout initialisé car cela déclanche le
'' chargement de F_Schedule et le demande de date au pc distant.
'-------------------------------------------------------------------------
TimeLagRecieved = False
'F_Schedule.Aff_Timelag(0).Caption = ""
'F_Schedule.Aff_Timelag(1).Caption = ""

' On demande le numéro de version au PC distant
'----------------------------------------------
Remote_Version = ""
Bt_CheckVersion_Click

End Sub



Private Sub Form_Unload(Cancel As Integer)

End


End Sub

Private Sub Hrlg_Timeout_Send_Timer()

Dim s1 As String
Dim s As String

If Send_Timeout_Chunk > 0 Then
    ' On indique au PC distant qu'on a envoyé le chuk numéro Send_Timeout_Chunk
    ' S'il ne l'a pas reçu, il redemandera ce qu'il n'a pas reçu.
    '--------------------------------------------------------------------------
    s1 = "CheckContentTrsf " & Trim(Str(Send_Timeout_Chunk)) & vbCrLf
    s = Prefix_cmd _
        & s1
    Set_Clipboard s
End If

End Sub


Private Sub Saisie_Action_Command_Click()

Dim I As Integer
Dim j As Integer

For I = 1 To 3
    Saisie_Action_Param(I).Text = ""
Next I

Select Case Saisie_SettingCategory.Text

    Case "Windows accounts management"
        For j = 1 To nb_t_Actions_Account
            If t_Actions_Account(j).Command = Saisie_Action_Command.Text Then
                For I = 3 To 1 Step -1
                    If I <= t_Actions_Account(j).nb_t_Param_lbl Then
                        Lbl_Setting_Functions_Param(I).Caption = t_Actions_Account(j).t_Param_lbl(I)
                        Saisie_Action_Param(I).Text = ""
                        Saisie_Action_Param(I).Visible = True
                    Else
                        Lbl_Setting_Functions_Param(I).Caption = ""
                        Saisie_Action_Param(I).Visible = False
                    End If
                Next I
                Aff_Action_Help.Caption = t_Actions_Account(j).Help
            End If
        Next j
        
    Case "TeamViewer management"
        For j = 1 To nb_t_Actions_TV
            If t_Actions_TV(j).Command = Saisie_Action_Command.Text Then
                For I = 1 To 3
                    If I <= t_Actions_TV(j).nb_t_Param_lbl Then
                        Lbl_Setting_Functions_Param(I).Caption = t_Actions_TV(j).t_Param_lbl(I)
                        Saisie_Action_Param(I).Visible = True
                    Else
                        Lbl_Setting_Functions_Param(I).Caption = ""
                        Saisie_Action_Param(I).Visible = False
                    End If
                Next I
                Aff_Action_Help.Caption = t_Actions_TV(j).Help
            End If
        Next j
        
    Case "Application management"
        For j = 1 To nb_t_Actions_App
            If t_Actions_App(j).Command = Saisie_Action_Command.Text Then
                For I = 1 To 3
                    If I <= t_Actions_App(j).nb_t_Param_lbl Then
                        Lbl_Setting_Functions_Param(I).Caption = t_Actions_App(j).t_Param_lbl(I)
                        Saisie_Action_Param(I).Visible = True
                    Else
                        Lbl_Setting_Functions_Param(I).Caption = ""
                        Saisie_Action_Param(I).Visible = False
                    End If
                Next I
                Aff_Action_Help.Caption = t_Actions_App(j).Help
            End If
        Next j
        
    Case "Mean functions"
        For j = 1 To nb_t_Actions_Mean
            If t_Actions_Mean(j).Command = Saisie_Action_Command.Text Then
                For I = 1 To 3
                    If I <= t_Actions_Mean(j).nb_t_Param_lbl Then
                        Lbl_Setting_Functions_Param(I).Caption = t_Actions_Mean(j).t_Param_lbl(I)
                        Saisie_Action_Param(I).Visible = True
                    Else
                        Lbl_Setting_Functions_Param(I).Caption = ""
                        Saisie_Action_Param(I).Visible = False
                    End If
                Next I
                Aff_Action_Help.Caption = t_Actions_Mean(j).Help
            End If
        Next j
        
    Case "Miscellaneous"
        For j = 1 To nb_t_Actions_Miscellaneous
            If t_Actions_Miscellaneous(j).Command = Saisie_Action_Command.Text Then
                For I = 1 To 3
                    If I <= t_Actions_Miscellaneous(j).nb_t_Param_lbl Then
                        Lbl_Setting_Functions_Param(I).Caption = t_Actions_Miscellaneous(j).t_Param_lbl(I)
                        Saisie_Action_Param(I).Visible = True
                    Else
                        Lbl_Setting_Functions_Param(I).Caption = ""
                        Saisie_Action_Param(I).Visible = False
                    End If
                Next I
                Aff_Action_Help.Caption = t_Actions_Miscellaneous(j).Help
            End If
        Next j

End Select
If Saisie_Action_Param(1).Visible Then
    Saisie_Action_Param(1).SetFocus
Else
    Bt_Go_SettingFunction.SetFocus
End If

End Sub

Private Sub Saisie_CurrentPW_GotFocus()

Aff_PC_Protected(2).Caption = ""

End Sub

Private Sub Saisie_NewPW_GotFocus(Index As Integer)

Aff_PC_Protected(2).Caption = ""

End Sub

Private Sub Saisie_PC_Click()

Dim j As Integer
Dim I As Integer
Dim s As String

If Frame_Aff(0).Visible Then TabStrip.SelectedItem.Selected = True

' On supprime l'éventuel alias et on note le nom du PC choisi
'------------------------------------------------------------
s = Saisie_PC.Text
I = InStr(1, s, "(")
If I > 0 Then s = Trim(Left(s, I - 1))
' S'il y a changement de PC...
'-----------------------------
If PC_Name <> s Then
    ' On reset l'affichage des données reçues
    '----------------------------------------
    Aff_Global_functions.Text = ""
    Me.Aff_CheckVersion.Caption = ""
    ' On resette l'affichage des paramètres TV
    '-----------------------------------------
    Chk_TVStartsWithWindows.Value = vbUnchecked
    Chk_TVChangesRequireAdmin.Value = vbUnchecked
    Chk_Lock_TV_Options.Value = vbUnchecked
    Chk_TVStartsWithWindows.Enabled = False
    Chk_TVChangesRequireAdmin.Enabled = False
    Chk_Lock_TV_Options.Enabled = False
    Lbl_CurrentTVValues.Visible = False
    
    ' On note qu'on ne sait pas le décalage horaire
    '----------------------------------------------
    TimeLagRecieved = False
    F_Schedule.Aff_Timelag(0).Caption = ""
    F_Schedule.Aff_Timelag(1).Caption = ""
    
    ' On réplique le nom du PC courant dans le libellé accompagnant
    ' la fonction de protection par un mot de passe et on efface toutes les saisies de mot de passe.
    '-----------------------------------------------------------------------------------------------
    Aff_PC_Protected(1).Caption = Saisie_PC.Text
    Aff_PC_Protected(2).Caption = ""
    Saisie_CurrentPW.Text = ""
    Saisie_NewPW(0).Text = ""
    Saisie_NewPW(1).Text = ""
    Saisie_ToolPW.Text = PC_PW(Saisie_PC.ListIndex + 1)
End If
PC_Name = s

' On efface les logs
'-------------------
Aff_Logs.Text = ""

' Si le nouveau choix est dans la liste, on le sauvegarde
'--------------------------------------------------------
For j = 0 To Saisie_PC.ListCount - 1
    If Saisie_PC.Text = Saisie_PC.List(j) Then
'        Fic_ini = Environ("TMP") & "\D_htvp.ini"
        Fic_ini = "C:\Program Files (x86)\TeamViewer" & "\D_htvp.ini"
        EcrireIni "PCSelected", "PCSelected", Saisie_PC.Text, Fic_ini
        Exit For
    End If
Next j

' On demande le numéro de version au pc distant
'----------------------------------------------
Bt_CheckVersion_Click

End Sub


Private Sub Saisie_PC_KeyDown(KeyCode As Integer, Shift As Integer)

' Si touche suppr
'----------------
If KeyCode = 46 Then
    ' Si le Suppr s'applique sur le nom complet
    '------------------------------------------
    If Saisie_PC.SelLength = Len(Saisie_PC.Text) Then
        PC_To_Delete = Saisie_PC.Text
    Else
        PC_To_Delete = ""
    End If
End If

End Sub

Private Sub Saisie_PC_KeyUp(KeyCode As Integer, Shift As Integer)

Dim I As Integer
Dim s As String
Dim s1 As String
Dim j As Integer
Dim res As Integer

' Si touche CR
'-------------
If KeyCode = 13 Then
    Enter_New_PC
' Si touche suppr
'----------------
ElseIf KeyCode = 46 Then
    ' Si le nom du PC à supprimer n'est pas vide
    '-------------------------------------------
    If PC_To_Delete <> "" Then
        res = MsgBox("Do You really want to delete """ & PC_To_Delete & """?", vbOKCancel)
        If res = vbOK Then
            ' On cherche si le PC à effecer est déjà dans la liste
            '-----------------------------------------------------
            For j = 0 To Saisie_PC.ListCount - 1
                s1 = Saisie_PC.List(j)
'                i = InStr(1, s1, "(")
'                If i > 0 Then s1 = Trim(Left(s1, i - 1))
                If PC_To_Delete = s1 Then Exit For
            Next j
            ' s'il y est...
            '--------------
            If j < Saisie_PC.ListCount Then
                ' On l'efface
                '------------
                Saisie_PC.RemoveItem (j)
                ' On sauvegarde la liste ainsi que les mots de passe et le choix courant
                '-----------------------------------------------------------------------
'                Fic_ini = Environ("TMP") & "\D_htvp.ini"
                Fic_ini = "C:\Program Files (x86)\TeamViewer" & "\D_htvp.ini"
                For j = 0 To Saisie_PC.ListCount - 1
                    EcrireIni "PCList", "PC" & Trim(Str(j)), Saisie_PC.List(j), Fic_ini
                    EcrireIni "PWList", "PW" & Trim(Str(j)), PC_PW(j + 1), Fic_ini
                Next j
                EcrireIni "PCSelected", "PCSelected", Saisie_PC.Text, Fic_ini
            End If
        Else
            Saisie_PC.Text = PC_To_Delete
        End If
    End If
End If

End Sub

Public Sub Init_PC()

End Sub


Private Sub Saisie_PC_LostFocus()

Enter_New_PC

End Sub

Private Sub Saisie_SettingCategory_Click()

Dim I As Integer

Saisie_Action_Command.Clear
For I = 1 To 3
    Lbl_Setting_Functions_Param(I).Caption = ""
    Saisie_Action_Param(I).Visible = False
Next I
Aff_Action_Help.Caption = ""

Select Case Saisie_SettingCategory.Text

    Case "Windows accounts management"
        Init_Saisie_Actions_Account
    Case "TeamViewer management"
        Init_Saisie_TV_Mgt
    Case "Application management"
        Init_Saisie_App_Mgt
    Case "Mean functions"
        Init_Saisie_Mean
    Case "Miscellaneous"
        Init_Saisie_Miscellaneous

End Select
Lbl_Command.Visible = True
Saisie_Action_Command.Visible = True
Bt_Go_SettingFunction.Visible = True

End Sub


Private Sub Saisie_ToolPW_KeyUp(KeyCode As Integer, Shift As Integer)

Dim j As Integer

If KeyCode = 13 Then
'    Fic_ini = Environ("TMP") & "\D_htvp.ini"
    Fic_ini = "C:\Program Files (x86)\TeamViewer" & "\D_htvp.ini"
    For j = 0 To Saisie_PC.ListCount - 1
        If Saisie_PC.List(j) = Saisie_PC.Text Then
            PC_PW(j + 1) = Saisie_ToolPW.Text
            EcrireIni "PWList", "PW" & Trim(Str(j)), PC_PW(j + 1), Fic_ini
            Exit For
        End If
    Next j
End If

End Sub

Private Sub Saisie_Wallpaper_Change()

Aff_WallpaperProgression.Caption = ""

End Sub

Private Sub hrlg_Clipboard_Timer()

TraiteClipboard

End Sub

Public Sub TraiteClipboard()

Dim i_deb As Long
Dim i_fin As Long
Dim ligne As String
Dim s As String
Dim I As Long
Dim c As String
Dim rep_command As String
Dim toto As Variant
Dim lg_prefix_loc As Integer

On Error GoTo Sautec

s = ""
If Clipboard.GetFormat(vbCFText) Then
    ClpBrd = Clipboard.GetText(vbCFText)
    ' Extraction de la première ligne
    '--------------------------------
    i_fin = InStr(1, ClpBrd, vbCr)
    If i_fin > 0 Then
        ligne = Left(ClpBrd, i_fin - 1)
        i_deb = i_fin + 2
        ' La première ligne doit commencer par "Return " & L_Pefix & " " & PC_Name
        ' ou bien par "Return " & L_Pefix & " " & This_PC_Name & " " & PC_Name
        '-------------------------------------------------------------------------
        I = InStr(1, UCase(ligne), UCase("Return " & L_Prefix & " " & This_PC_Name & " " & PC_Name))
        If I = 1 Then
            lg_prefix_loc = Len("Return " & L_Prefix & " " & This_PC_Name & " " & PC_Name)
        Else
            I = InStr(1, UCase(ligne), UCase("Return " & L_Prefix & " " & PC_Name))
            lg_prefix_loc = Len("Return " & L_Prefix & " " & PC_Name)
        End If
        If I = 1 Then
            ' On efface le clipboard
            '-----------------------
            Clipboard.Clear
            ' On extrait la commande dont c'est la réponse
            '---------------------------------------------
            rep_command = Right(ligne, Len(ligne) - lg_prefix_loc - 1)
            ' On traite la réponse en fonction de la commande initiale
            '---------------------------------------------------------
            TraiteRéponse rep_command, Right(ClpBrd, Len(ClpBrd) - i_deb + 1)
        End If
    End If
End If

Sautec:

End Sub

Public Sub TraiteRéponse(Command As String, Réponse As String)

Dim I As Integer
Dim j As Integer
Dim s As String

Me.Aff_CheckVersion.FontBold = False
Me.Aff_CheckVersion.ForeColor = 0
s = "Connection OK"
Select Case Command

    Case "GetInfo"
        Me.Aff_Global_functions.Text = Réponse
    Case "GetAccountsList"
        Me.Aff_Global_functions.Text = Réponse
    Case "GetTVParameters"
        Me.Aff_Global_functions.Text = Réponse
    Case "GethtvpStatus"
        Me.Aff_Global_functions.Text = Réponse
    Case "GetTVOptionParameters"
        I = InStr(1, Réponse, "Starts with Windows: Yes")
        If I > 0 Then
            Chk_TVStartsWithWindows.Value = vbChecked
        Else
            Chk_TVStartsWithWindows.Value = vbUnchecked
        End If
        Chk_TVStartsWithWindows.Enabled = True
        I = InStr(1, Réponse, "Changes require admin rights: Yes")
        If I > 0 Then
            Chk_TVChangesRequireAdmin.Value = vbChecked
        Else
            Chk_TVChangesRequireAdmin.Value = vbUnchecked
        End If
        Chk_TVChangesRequireAdmin.Enabled = True
        I = InStr(1, Réponse, "TV Options password: Yes")
        If I > 0 Then
            Chk_Lock_TV_Options.Value = vbChecked
        Else
            Chk_Lock_TV_Options.Value = vbUnchecked
        End If
        Chk_Lock_TV_Options.Enabled = True
        Lbl_CurrentTVValues.Visible = True
    Case "GetAccountDetails"
        Me.Aff_Global_functions.Text = Réponse
    Case "GetVersion"
        Remote_Version = Réponse
        Me.Aff_Global_functions.Text = Réponse
        s = "Connection OK (" & Réponse & ")"
    Case "ProtectTool"
        Me.Aff_PC_Protected(2).Caption = "Password changed successfully"
        Saisie_ToolPW.Text = Saisie_NewPW(0).Text
'        Fic_ini = Environ("TMP") & "\D_htvp.ini"
        Fic_ini = "C:\Program Files (x86)\TeamViewer" & "\D_htvp.ini"
        For j = 0 To Saisie_PC.ListCount - 1
            If Saisie_PC.List(j) = Saisie_PC.Text Then
                If Chk_SavePW.Value = vbChecked Then
                    PC_PW(j + 1) = Saisie_ToolPW.Text
                Else
                    PC_PW(j + 1) = ""
                End If
                    EcrireIni "PWList", "PW" & Trim(Str(j)), PC_PW(j + 1), Fic_ini
                Exit For
            End If
        Next j
    Case "ContentTrsf"
        I = InStr(1, Réponse, " ")
        If I > 0 Then
            SendNextChunk Val(Left(Réponse, I - 1)), Right(Réponse, Len(Réponse) - I)
        End If
    Case "WrongPW"
        Me.Aff_CheckVersion.FontBold = True
        Me.Aff_CheckVersion.ForeColor = RGB(255, 0, 0)
        s = "Wrong password"
    Case "GetScreenResolution"
        Set_Preview_Welcome_Size Réponse
    Case "TrsfComplete"
        TraiteTrsfComplete Réponse
    Case "GetDate"
        TraiteGetDate Réponse
End Select
Me.Aff_CheckVersion.Caption = s

End Sub

Public Sub Init_Actions()

Dim I As Integer

' Actions sur les comptes Windows
'--------------------------------
I = 0
nb_t_Actions_Account = 8
ReDim t_Actions_Account(nb_t_Actions_Account)

I = I + 1
t_Actions_Account(I).Command = "CreateAdminAccount"
t_Actions_Account(I).nb_t_Param_lbl = 2
ReDim t_Actions_Account(I).t_Param_lbl(2)
t_Actions_Account(I).t_Param_lbl(1) = "Account name*"
t_Actions_Account(I).t_Param_lbl(2) = "Password"
t_Actions_Account(I).Help = "Creates a new Windows account and gives it administrator priviledges."

I = I + 1
t_Actions_Account(I).Command = "CreateStandardAccount"
t_Actions_Account(I).nb_t_Param_lbl = 2
ReDim t_Actions_Account(I).t_Param_lbl(2)
t_Actions_Account(I).t_Param_lbl(1) = "Account name*"
t_Actions_Account(I).t_Param_lbl(2) = "Password"
t_Actions_Account(I).Help = "Creates a new Windows account and gives it standard priviledges only."

I = I + 1
t_Actions_Account(I).Command = "RemoveAdminRights"
t_Actions_Account(I).nb_t_Param_lbl = 1
ReDim t_Actions_Account(I).t_Param_lbl(1)
t_Actions_Account(I).t_Param_lbl(1) = "Account name*"
t_Actions_Account(I).Help = "Removes administrator priviledges."

I = I + 1
t_Actions_Account(I).Command = "SetAdminRights"
t_Actions_Account(I).nb_t_Param_lbl = 1
ReDim t_Actions_Account(I).t_Param_lbl(1)
t_Actions_Account(I).t_Param_lbl(1) = "Account name*"
t_Actions_Account(I).Help = "Gives administrator priviledges."

I = I + 1
t_Actions_Account(I).Command = "SetAccountPW"
t_Actions_Account(I).nb_t_Param_lbl = 2
ReDim t_Actions_Account(I).t_Param_lbl(2)
t_Actions_Account(I).t_Param_lbl(1) = "Account name*"
t_Actions_Account(I).t_Param_lbl(2) = "Password"
t_Actions_Account(I).Help = "Set, change or remove the account password. The password is removed if the field is empty."

I = I + 1
t_Actions_Account(I).Command = "LockAccountPW"
t_Actions_Account(I).nb_t_Param_lbl = 1
ReDim t_Actions_Account(I).t_Param_lbl(1)
t_Actions_Account(I).t_Param_lbl(1) = "Account name"
t_Actions_Account(I).Help = "Prevents the user from changing the account password. Keep the Account name empty to lock the password of the current account."

I = I + 1
t_Actions_Account(I).Command = "ReleaseAccountPW"
t_Actions_Account(I).nb_t_Param_lbl = 1
ReDim t_Actions_Account(I).t_Param_lbl(1)
t_Actions_Account(I).t_Param_lbl(1) = "Account name"
t_Actions_Account(I).Help = "Allows the user to change the account password. Keep the Account name empty to release the password of the current account."

I = I + 1
t_Actions_Account(I).Command = "DeleteAccount"
t_Actions_Account(I).nb_t_Param_lbl = 1
ReDim t_Actions_Account(I).t_Param_lbl(1)
t_Actions_Account(I).t_Param_lbl(1) = "Account name*"
t_Actions_Account(I).Help = "Deletes the account."


' Actions sur les paramètres TV
'------------------------------
I = 0
nb_t_Actions_TV = 10
ReDim t_Actions_TV(nb_t_Actions_TV)

I = I + 1
t_Actions_TV(I).Command = "HideTVPanel"
t_Actions_TV(I).nb_t_Param_lbl = 0
t_Actions_TV(I).Help = "Hides the TeamViewer panel."

I = I + 1
t_Actions_TV(I).Command = "ShowTVPanel"
t_Actions_TV(I).nb_t_Param_lbl = 0
t_Actions_TV(I).Help = "Shows the TeamViewer panel."

I = I + 1
t_Actions_TV(I).Command = "HideTVComputerList"
t_Actions_TV(I).nb_t_Param_lbl = 0
t_Actions_TV(I).Help = "Hides the TeamViewer computer list."

I = I + 1
t_Actions_TV(I).Command = "ShowTVComputerList"
t_Actions_TV(I).nb_t_Param_lbl = 0
t_Actions_TV(I).Help = "Shows the TeamViewer computer list."

I = I + 1
t_Actions_TV(I).Command = "TVStartsWithWindows"
t_Actions_TV(I).nb_t_Param_lbl = 0
t_Actions_TV(I).Help = "Set TeamViewer to start with Windows."

I = I + 1
t_Actions_TV(I).Command = "TVRemovePersonalPWForUnattendedAccess"
t_Actions_TV(I).nb_t_Param_lbl = 0
t_Actions_TV(I).Help = "Remove the personal password for unattended access."

I = I + 1
t_Actions_TV(I).Command = "TVSetWindowsLogonForAllUsers"
t_Actions_TV(I).nb_t_Param_lbl = 0
t_Actions_TV(I).Help = "Set Windows logon allowed for all users."

I = I + 1
t_Actions_TV(I).Command = "TVFullAccessOnLogonScreen"
t_Actions_TV(I).nb_t_Param_lbl = 0
t_Actions_TV(I).Help = "Set TV to allow full access on logon screen."

I = I + 1
t_Actions_TV(I).Command = "TVEnableLogings"
t_Actions_TV(I).nb_t_Param_lbl = 0
t_Actions_TV(I).Help = "Enable the logings."

I = I + 1
t_Actions_TV(I).Command = "TVDisableLogings"
t_Actions_TV(I).nb_t_Param_lbl = 0
t_Actions_TV(I).Help = "Disable the logings."


' Actions sur les applications
'-----------------------------
I = 0
nb_t_Actions_App = 2
ReDim t_Actions_App(nb_t_Actions_App)

I = I + 1
t_Actions_App(I).Command = "ExecuteMaximized"
t_Actions_App(I).nb_t_Param_lbl = 3
ReDim t_Actions_App(I).t_Param_lbl(3)
t_Actions_App(I).t_Param_lbl(1) = "Application path and name*"
t_Actions_App(I).t_Param_lbl(2) = "First application parameter"
t_Actions_App(I).t_Param_lbl(3) = "Second application parameter"
t_Actions_App(I).Help = "Execute an application with its window maximized."

I = I + 1
t_Actions_App(I).Command = "OpenWebPageMaximized "
t_Actions_App(I).nb_t_Param_lbl = 1
ReDim t_Actions_App(I).t_Param_lbl(1)
t_Actions_App(I).t_Param_lbl(1) = "Web page address*"
t_Actions_App(I).Help = "Open a Web Page with its window maximized."


' Mean functions
'---------------
I = 0
nb_t_Actions_Mean = 3
ReDim t_Actions_Mean(nb_t_Actions_Mean)

I = I + 1
t_Actions_Mean(I).Command = "ReleaseMean"
t_Actions_Mean(I).nb_t_Param_lbl = 0
t_Actions_Mean(I).Help = "Ends any active mean function"

I = I + 1
t_Actions_Mean(I).Command = "MeanAddToClipboard"
t_Actions_Mean(I).nb_t_Param_lbl = 1
ReDim t_Actions_Mean(I).t_Param_lbl(1)
t_Actions_Mean(I).t_Param_lbl(1) = "Suffix text*"
t_Actions_Mean(I).Help = "When activated, each time a text is copied to the clipboard, the text <Suffix Text> is added at the end of it. Use /n to insert a line feed in your text."

I = I + 1
t_Actions_Mean(I).Command = "MeanImposeAppli"
t_Actions_Mean(I).nb_t_Param_lbl = 3
ReDim t_Actions_Mean(I).t_Param_lbl(3)
t_Actions_Mean(I).t_Param_lbl(1) = "Application path and name*"
t_Actions_Mean(I).t_Param_lbl(2) = "Application parameter"
t_Actions_Mean(I).t_Param_lbl(3) = "Reactivation timeout"
t_Actions_Mean(I).Help = "Launches the application with an optional parameter and keeps it activated and maximized. The optional [Reactivation timout] is the timout in seconds for the application to come back full screen in case the slave minimizes it or even closes it. The default value is 5 seconds."


' Miscellaneous
'--------------
I = 0
nb_t_Actions_Miscellaneous = 3
ReDim t_Actions_Miscellaneous(nb_t_Actions_Miscellaneous)

I = I + 1
t_Actions_Miscellaneous(I).Command = "ProtectTool"
t_Actions_Miscellaneous(I).nb_t_Param_lbl = 1
ReDim t_Actions_Miscellaneous(I).t_Param_lbl(1)
t_Actions_Miscellaneous(I).t_Param_lbl(1) = "Password"
t_Actions_Miscellaneous(I).Help = "Protects the tool with a password."

I = I + 1
t_Actions_Miscellaneous(I).Command = "Exit"
t_Actions_Miscellaneous(I).nb_t_Param_lbl = 0
t_Actions_Miscellaneous(I).Help = "Ends the tool on the slave's PC. Warning, the next login, relaunches it."

I = I + 1
t_Actions_Miscellaneous(I).Command = "Uninstall"
t_Actions_Miscellaneous(I).nb_t_Param_lbl = 0
t_Actions_Miscellaneous(I).Help = "Ends the tool on the slave's PC and uninstall it."

End Sub

Public Sub Init_Saisie_Actions_Account()

Dim I As Integer

Saisie_Action_Command.Clear
For I = 1 To nb_t_Actions_Account
    Saisie_Action_Command.AddItem t_Actions_Account(I).Command
Next I
For I = 1 To 3
    Lbl_Setting_Functions_Param(I).Caption = ""
    Saisie_Action_Param(I).Visible = False
Next I

End Sub

Public Sub Init_Saisie_TV_Mgt()

Dim I As Integer

Saisie_Action_Command.Clear
For I = 1 To nb_t_Actions_TV
    Saisie_Action_Command.AddItem t_Actions_TV(I).Command
Next I
For I = 1 To 3
    Lbl_Setting_Functions_Param(I).Caption = ""
    Saisie_Action_Param(I).Visible = False
Next I

End Sub

Public Sub Init_Saisie_App_Mgt()

Dim I As Integer

Saisie_Action_Command.Clear
For I = 1 To nb_t_Actions_App
    Saisie_Action_Command.AddItem t_Actions_App(I).Command
Next I
For I = 1 To 3
    Lbl_Setting_Functions_Param(I).Caption = ""
    Saisie_Action_Param(I).Visible = False
Next I

End Sub

Public Sub Init_Saisie_Mean()

Dim I As Integer

Saisie_Action_Command.Clear
For I = 1 To nb_t_Actions_Mean
    Saisie_Action_Command.AddItem t_Actions_Mean(I).Command
Next I
For I = 1 To 3
    Lbl_Setting_Functions_Param(I).Caption = ""
    Saisie_Action_Param(I).Visible = False
Next I

End Sub

Public Sub Init_Saisie_Miscellaneous()

Dim I As Integer

Saisie_Action_Command.Clear
For I = 1 To nb_t_Actions_Miscellaneous
    Saisie_Action_Command.AddItem t_Actions_Miscellaneous(I).Command
Next I
For I = 1 To 3
    Lbl_Setting_Functions_Param(I).Caption = ""
    Saisie_Action_Param(I).Visible = False
Next I

End Sub

Public Sub Init_Catégories_Actions()

Saisie_SettingCategory.Clear
Saisie_SettingCategory.AddItem "Windows accounts management"
Saisie_SettingCategory.AddItem "TeamViewer management"
Saisie_SettingCategory.AddItem "Application management"
Saisie_SettingCategory.AddItem "Mean functions"
Saisie_SettingCategory.AddItem "Miscellaneous"

End Sub

Public Sub SendNextChunk(LastChunk As Double, TransferID As String)

Dim strData As String
Dim s1 As String
Dim s2 As String
Dim s As String
Dim c As String
Dim I As Long
Dim SendProgression As Long

' Si l'id de transfert n'est pas le bon, on n'est pas concerné
'-------------------------------------------------------------
If Left(Send_ID, 6) <> Left(TransferID, 6) Then Exit Sub

' On désarme le timeout de réponse
'---------------------------------
Send_Timeout_Chunk = 0
Hrlg_Timeout_Send.Enabled = False
Hrlg_Timeout_Send.Interval = 0
' Si tout a été envoyé...
'------------------------
If Send_ChunkSize * LastChunk >= Send_length Then
    ' Affichage de la progression
    '----------------------------
    Select Case Send_DispProgress
        Case "WallPaper"
            Aff_WallpaperProgression.Caption = "Done"
        Case "Transfer"
            Aff_TransferProgression.Caption = "Done"
        Case "Launch"
            Aff_LaunchProgression.Caption = "Done"
        Case "WelcomeBackground"
            Aff_WelcomeBackgroundProgression.Caption = "Done"
    End Select
    ' On envoie la commande de fin incluant dans les options l'action à exécuter
    '---------------------------------------------------------------------------
    s1 = "EndTrsf """ & Send_FileNamesent & """ " & Format(Send_Options, "00") & vbCrLf
    s = Prefix_cmd _
    & s1
    Set_Clipboard s
' Sinon, on envoie le chunk suivant
'----------------------------------
Else
    ' Affichage de la progression
    '----------------------------
    SendProgression = (100 * Send_ChunkSize * LastChunk) / Send_length
    Select Case Send_DispProgress
        Case "WallPaper"
            Aff_WallpaperProgression.Caption = "Upload in progress: " & Trim(Str(SendProgression)) & "%"
        Case "Transfer"
            Aff_TransferProgression.Caption = "Upload in progress: " & Trim(Str(SendProgression)) & "%"
        Case "Launch"
            Aff_LaunchProgression.Caption = "Upload in progress: " & Trim(Str(SendProgression)) & "%"
        Case "WelcomeBackground"
            Aff_WelcomeBackgroundProgression.Caption = "Upload in progress: " & Trim(Str(SendProgression)) & "%"
    End Select
    ' Le chunk à envoyer
    '-------------------
    strData = Mid(Send_Content, (Send_ChunkSize * LastChunk) + 1, Send_ChunkSize)
    ' Codification Hexa du chunk
    '---------------------------
    For I = 1 To Len(strData)
        c = Hex(Asc(Mid(strData, I, 1)))
        If Len(c) < 2 Then
            c = "0" & c
        End If
        s2 = s2 & c
        DoEvents
    Next I
    ' Préparation de la commande
    '---------------------------
    s1 = "ContentTrsf " & Trim(Str(LastChunk + 1)) & vbCrLf
    s = Prefix_cmd _
        & s1 _
        & s2
    ' Juste avant envoi, on arme le timeout de réponse
    '-------------------------------------------------
    Send_Timeout_Chunk = LastChunk + 1
    Hrlg_Timeout_Send.Interval = 5000
    Hrlg_Timeout_Send.Enabled = True
    ' Envoi de la commande
    '---------------------
    Set_Clipboard s
End If

End Sub





Private Sub TabStrip_Click()

Dim I As Integer

For I = 0 To TabStrip.Tabs.Count
    If TabStrip.SelectedItem.Index = I Then
        Frame_Aff(I).Visible = True
    Else
        Frame_Aff(I).Visible = False
    End If
Next I
If TabStrip.SelectedItem.Caption = "TeamViewer" Then Bt_RefreshCurrentTVValues_Click
If TabStrip.SelectedItem.Caption = "Backgrounds" Then GetScreenResolution

End Sub

Public Sub GetScreenResolution()

Dim s1 As String
Dim s As String

s1 = "GetScreenResolution" & vbCrLf
s = Prefix_cmd _
        & s1
Set_Clipboard s

End Sub

Public Sub Set_Preview_Welcome_Size(Réponse As String)

Dim I As Integer
Dim w As Long
Dim h As Long
Dim maxw As Long
Dim maxh As Long
Dim wf As Long
Dim hf As Long
Dim H_sur_W As Double

I = InStr(1, Réponse, "|")
If I < 2 Then Exit Sub

w = Val(Left(Réponse, I - 1))
h = Val(Right(Réponse, Len(Réponse) - I))

picCombined.ScaleMode = vbPixels
picCombined.AutoRedraw = True
picCombined.Move picCombined.Left, picCombined.Top, _
       ScaleX(w, vbPixels, Me.ScaleMode), _
       ScaleY(h, vbPixels, Me.ScaleMode)
BackGroundTmpFile = Environ("TMP") & "\WelcomeScreen.jpg"
H_sur_W = h / w
' Si on doit prendre toute la hauteur disponible pour le preview... Aff_Preview_Background_h
'------------------------------------------------------------------
If (h * Aff_Preview_Background_w) > (w * Aff_Preview_Background_h) Then
    Aff_Preview_Background.Height = Aff_Preview_Background_h * Screen.TwipsPerPixelY
    Aff_Preview_Background.Width = ((w * Aff_Preview_Background_h) / h) * Screen.TwipsPerPixelX
    Aff_Preview_Background.Top = Aff_Preview_Background_t
    Aff_Preview_Background.Left = Aff_Preview_Background_l + ((Aff_Preview_Background_w * Screen.TwipsPerPixelX - Aff_Preview_Background.Width) / 2)
Else
    ' On prend toute la largeur disponible
    ' Le rapport de réduction est donc celui des largeurs,
    ' à appliquer à la hauteur
    '-----------------------------------------------------
    Aff_Preview_Background.Width = Aff_Preview_Background_w * Screen.TwipsPerPixelX
    Aff_Preview_Background.Height = ((h * Aff_Preview_Background_w) / w) * Screen.TwipsPerPixelY
    Aff_Preview_Background.Left = Aff_Preview_Background_l
    Aff_Preview_Background.Top = Aff_Preview_Background_t + ((Aff_Preview_Background_h * Screen.TwipsPerPixelY - Aff_Preview_Background.Height) / 2)
End If

End Sub


Public Sub Add_Scheduled_Action(Action As String)

If ScheduleActions = "" Then
    ScheduleActions = Action
Else
    ScheduleActions = ScheduleActions & vbCrLf & Action
End If
F_Actions.Show
F_Actions.Saisie_ActionListToBeScheduled.Text = ScheduleActions
F_Actions.SetFocus

End Sub

Public Sub Send_Action(Action As String)

Dim s1 As String
Dim s As String

If ScheduleOn Then
    Add_Scheduled_Action Action
Else
    s = Prefix_cmd & Action
    
    Set_Clipboard s
    
    If Aff_Logs.Text = "" Then
        Aff_Logs.Text = Action
    Else
        Aff_Logs.Text = Aff_Logs.Text & vbCrLf & Action
    End If
End If

End Sub

Public Sub Init_GDI_Plus()

' validate user has GDI+ and setup classes for this sample project
Set cGDIplus = New cGDIpToken
If cGDIplus.Token = 0& Then
    GDIPlusOK = False
'    lstExamples.Enabled = False
'    cboSample.Enabled = False
'    Picture1.Enabled = False
'    MsgBox "Cannot run this project on your computer. GDI+ is required and it could not be loaded.", vbExclamation + vbOKOnly
Else
    GDIPlusOK = True
    Set cImage = New cGDIpImage
    Set cGraphics = New cGDIpRenderer
    Set cBrushes = New cGDIpPenBrush
    cBrushes.AttachTokenClass cGDIplus
    cGraphics.AttachTokenClass cGDIplus
'    cboBkg.Enabled = True
End If

End Sub

Public Function Add_BackgroundPic() As Boolean

Dim l As Double
Dim bData() As Byte, fn As Integer, bOk As Boolean

On Error GoTo Add_BackgroundPicError

' Pour le redimensionnement
'--------------------------
Dim tmpPic As StdPicture
Dim picWidth As Long, picHeight As Long
Dim picW As Long, picH As Long
Dim maxCx As Long, maxCy As Long
picCombined.ScaleMode = vbPixels
picCombined.AutoRedraw = True
picCombined.BorderStyle = 0&
Set tmpPic = LoadPicture(Me.Saisie_WelcomeBackground.Text)
picWidth = ScaleX(tmpPic.Width, vbHimetric, vbPixels)
picHeight = ScaleY(tmpPic.Height, vbHimetric, vbPixels)
' Si c'est la hauteur qui doit être mise au max...
'-------------------------------------------------
If (picHeight * picCombined.Width) > (picWidth * picCombined.Height) Then
    picH = picCombined.Height / Screen.TwipsPerPixelY
    picW = (picWidth * picCombined.Height / Screen.TwipsPerPixelY) / picHeight
    If Me.opt_WelcomeAlignment(0).Value Then
        picCombined.PaintPicture tmpPic, 0&, 0&, picW, picH
    ElseIf Me.opt_WelcomeAlignment(1).Value Then
        picCombined.PaintPicture tmpPic, ((picCombined.Width / Screen.TwipsPerPixelX) - picW) / 2, 0&, picW, picH
    Else
        picCombined.PaintPicture tmpPic, ((picCombined.Width / Screen.TwipsPerPixelX) - picW), 0&, picW, picH
    End If
' Si c'est la largeur qui doit être mise au max...
'-------------------------------------------------
Else
    picH = (picHeight * picCombined.Width / Screen.TwipsPerPixelX) / picWidth
    picW = picCombined.Width / Screen.TwipsPerPixelX
    picCombined.PaintPicture tmpPic, 0&, ((picCombined.Height / Screen.TwipsPerPixelY) - picH) / 2, picW, picH
End If
' Recopie dans le preview
'------------------------
Aff_Preview_Background.Picture = picCombined.Image

Add_BackgroundPic = True
Exit Function

Add_BackgroundPicError:
Add_BackgroundPic = False

End Function

Public Sub TraiteTrsfComplete(s_Options As String)

Dim options As Integer

' Nom complet du fichier à créer, fonction de la destination
'-----------------------------------------------------------
options = Val(s_Options)
' S'il s'agit de lancer un programme...
'--------------------------------------
If (options And Trsf_Launch) Then
    Add_Scheduled_Action "LaunchProgram"
'    dbExecWindows = Shell(Recep_FilePathAndName, 1)
' S'il s'agit de mettre un papier peint non permanent...
'-------------------------------------------------------
ElseIf (options And Trsf_WallPaper) And ((options And Trsf_Permanent) <> Trsf_Permanent) Then
    Add_Scheduled_Action "WallPaper"
' S'il s'agit de mettre un papier peint permanent...
'---------------------------------------------------
ElseIf (options And (Trsf_WallPaper + Trsf_Permanent)) Then
    Add_Scheduled_Action "PermanentWallPaper"
' S'il s'agit de mettre un écran d'acceuil non permanent...
'----------------------------------------------------------
ElseIf (options And Trsf_Welcome) And ((options And Trsf_Permanent) <> Trsf_Permanent) Then
    Add_Scheduled_Action "WelcomeScreen"
' S'il s'agit de mettre un écran d'acceuil permanent...
'------------------------------------------------------
ElseIf (options And (Trsf_Welcome + Trsf_Permanent)) Then
    Add_Scheduled_Action "PermanentWelcomeScreen"
End If

End Sub

Public Sub Ask_Remote_Date()
    
Dim s As String

s = Prefix_cmd & "GetDate"

Set_Clipboard s

End Sub

Public Sub TraiteGetDate(RemoteDate As String)

Dim d As Date
Dim s As String

If Not IsDate(RemoteDate) Then Exit Sub
d = RemoteDate
Local_minus_Remote_Time = d - Now
TimeLagRecieved = True
s = "Timelag: " & Format(Local_minus_Remote_Time, "hh:nn")
F_Schedule.Aff_Timelag(0).Caption = s
F_Schedule.Aff_Timelag(1).Caption = s

End Sub

Public Function FusionneImageFileName(Nom1 As String, Nom2 As String) As String

Dim I As Integer
Dim s1 As String
Dim s2 As String

' On enlève l'éventuelle extension du 1
'--------------------------------------
For I = Len(Nom1) To 1 Step -1
    If Mid(Nom1, I, 1) = "." Then Exit For
Next I
If I > 0 Then
    s1 = Left(Nom1, I - 1)
Else
    s1 = Nom1
End If

' On enlève l'éventuelle extension du 2
'--------------------------------------
For I = Len(Nom2) To 1 Step -1
    If Mid(Nom2, I, 1) = "." Then Exit For
Next I
If I > 0 Then
    s2 = Left(Nom2, I - 1)
Else
    s2 = Nom2
End If

' On enlève l'éventuel chemin du 2
'---------------------------------
For I = Len(s2) To 1 Step -1
    If Mid(s2, I, 1) = "\" Then Exit For
Next I
If I > 0 Then
    s2 = Right(s2, Len(s2) - I)
End If

' On fabrique le nom de sortie
'-----------------------------
FusionneImageFileName = s1 & "_" & s2 & ".jpg"

End Function

Public Sub Enter_New_PC()

Dim I As Integer
Dim s As String
Dim s1 As String
Dim j As Integer
Dim res As Integer

' On supprime l'éventuel alias et on note le nom du PC choisi
'------------------------------------------------------------
s = Saisie_PC.Text
I = InStr(1, s, "(")
If I > 0 Then s = Trim(Left(s, I - 1))
PC_Name = s

' On cherche si le PC est déjà dans la liste
'-------------------------------------------
For j = 0 To Saisie_PC.ListCount - 1
    s1 = Saisie_PC.List(j)
    I = InStr(1, s1, "(")
    If I > 0 Then s1 = Trim(Left(s1, I - 1))
    If s = s1 Then Exit For
Next j
' s'il n'y est pas...
'--------------------
If j >= Saisie_PC.ListCount Then
    ' On l'ajoute ainsi qu'un mot de passe vide par défaut
    '-----------------------------------------------------
    Saisie_PC.AddItem Saisie_PC.Text
    Saisie_PC.SelStart = 0
    Saisie_PC.SelLength = 100
    nb_PC_PW = nb_PC_PW + 1
    ReDim Preserve PC_PW(nb_PC_PW)
    PC_PW(nb_PC_PW) = ""
    Saisie_ToolPW.Text = ""
' Sinon, si l'alias est différent on met à jour
'----------------------------------------------
ElseIf Saisie_PC.Text <> Saisie_PC.List(j) Then
    Saisie_PC.List(j) = Saisie_PC.Text
    Saisie_PC.Text = Saisie_PC.List(j)
    Saisie_PC.SelStart = 0
    Saisie_PC.SelLength = 100
End If

' On initialise les affichages
'-----------------------------
Init_PC

' On sauvegarde la liste ainsi que les mots de passe et le choix courant
'-----------------------------------------------------------------------
'Fic_ini = Environ("TMP") & "\D_htvp.ini"
Fic_ini = "C:\Program Files (x86)\TeamViewer" & "\D_htvp.ini"
For j = 0 To Saisie_PC.ListCount - 1
    EcrireIni "PCList", "PC" & Trim(Str(j)), Saisie_PC.List(j), Fic_ini
    EcrireIni "PWList", "PW" & Trim(Str(j)), PC_PW(j + 1), Fic_ini
Next j
EcrireIni "PCSelected", "PCSelected", Saisie_PC.Text, Fic_ini

End Sub

Public Sub Sort_PC_Liste()

Dim t_i As New Tri_Index_Class
Dim I As Long
Dim j As Long
Dim k As Long
Dim t_l() As String
Dim s As String
Dim nb As Long

nb = Me.Saisie_PC.ListCount
t_i.Clear
ReDim t_l(nb)

For I = 0 To nb - 1
    t_l(I + 1) = Me.Saisie_PC.List(I)
    ' On extrait ce qui est entre parenthèse pour le prendre comme critère de tri
    ' S'il manque la parenthèse droite, ce n'est pas grave et s'il n'y a pas de parenthèses
    ' on prend le nom entier.
    '----------------------------------------------------------------------------
    s = t_l(I + 1)
    j = InStr(1, s, "(")
    If j > 0 Then
        k = InStr(j, s, ")")
        If k > 0 Then
            s = Mid(s, j + 1, k - j - 1)
        Else
            s = Right(s, Len(s) - j)
        End If
    End If
    t_i.Add s
Next I
t_i.Sort
' On efface la liste et on la remplit de nouveau mais dans l'ordre
'-----------------------------------------------------------------
Me.Saisie_PC.Clear
For I = 1 To nb
    Me.Saisie_PC.AddItem t_l(t_i.I(I))
Next I

End Sub
