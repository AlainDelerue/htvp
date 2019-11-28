VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form F_Main 
   Caption         =   "LookAtMe task setup"
   ClientHeight    =   9480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   13950
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture_Consigne 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   4920
      Picture         =   "F_Main.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   1980
      TabIndex        =   99
      Top             =   1440
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox Disp_pic 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3135
      Left            =   4440
      ScaleHeight     =   3075
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.Timer AlertBeamDuration 
         Enabled         =   0   'False
         Left            =   3000
         Top             =   1080
      End
      Begin VB.Label Lbl_Failure 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2175
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   5295
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame_Instructions 
      Height          =   2415
      Left            =   7560
      TabIndex        =   95
      Top             =   3240
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Label Lbl_To_Be_Written 
         Alignment       =   2  'Center
         Caption         =   "I love Your cock"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   98
         Top             =   1680
         Width           =   3495
         WordWrap        =   -1  'True
      End
      Begin VB.Label Lbl_Instruction 
         Caption         =   "- Click on this exact location on the picture and then write this sentence:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   97
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label23 
         Caption         =   "When you see the alert beam, you must:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   96
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.TextBox Input_Slave 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   5160
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.Frame Frame_Mouse 
      Height          =   615
      Left            =   11880
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame_Results 
      Height          =   2415
      Left            =   4800
      TabIndex        =   7
      Top             =   6240
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Label Disp_Unit 
         Caption         =   "hours web exposure"
         Height          =   255
         Index           =   4
         Left            =   900
         TabIndex        =   18
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Disp_Tax 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Disp_Unit 
         Caption         =   "day(s) chastity cage"
         Height          =   255
         Index           =   3
         Left            =   900
         TabIndex        =   16
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Disp_Tax 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Disp_Unit 
         Caption         =   ">5 ==>Keylogger"
         Height          =   255
         Index           =   2
         Left            =   900
         TabIndex        =   14
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Disp_Tax 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Disp_Unit 
         Caption         =   "pic(s) of me naked"
         Height          =   255
         Index           =   1
         Left            =   900
         TabIndex        =   12
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Disp_Tax 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Disp_Unit 
         Caption         =   "$"
         Height          =   255
         Index           =   0
         Left            =   900
         TabIndex        =   10
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Disp_Tax 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Penalties for failure"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.CommandButton Bt_No_Need 
      Caption         =   "No, I've already been fucked by this tool"
      Default         =   -1  'True
      Height          =   855
      Left            =   11520
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Bt_Yes_Need 
      Caption         =   "Yes, I am a beginner"
      Height          =   855
      Left            =   9960
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame__3 
      Height          =   2415
      Left            =   8880
      TabIndex        =   90
      Top             =   4680
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton Bt_S_Load 
         Caption         =   "Load a task file"
         Height          =   375
         Left            =   1320
         TabIndex        =   92
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Lbl_Remaining_Time 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   94
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Lbl_Remaining_Time_Caption 
         AutoSize        =   -1  'True
         Caption         =   "Remaining time (minutes):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3240
         TabIndex        =   93
         Top             =   1200
         Visible         =   0   'False
         Width           =   3150
      End
      Begin VB.Label Lbl_S_Accueil 
         Alignment       =   2  'Center
         Caption         =   "Load your task, slave! Hurry up!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   91
         Top             =   240
         Width           =   4095
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame__1 
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   0
      TabIndex        =   32
      Top             =   -120
      Width           =   4575
      Begin VB.CommandButton Bt_Select_Pic 
         Caption         =   "Chose a picture"
         Height          =   375
         Left            =   480
         TabIndex        =   75
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Bt_Reset_points 
         Caption         =   "Delete focus points"
         Height          =   375
         Left            =   480
         TabIndex        =   74
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Input_Diam_Beam 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   73
         Text            =   "20"
         Top             =   8280
         Width           =   615
      End
      Begin VB.CommandButton Bt_Test_Alert_Beam 
         Caption         =   "Test with the last focus point"
         Height          =   495
         Left            =   2520
         TabIndex        =   72
         Top             =   8400
         Width           =   1335
      End
      Begin VB.TextBox Input_AlertBeam_Duration 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   71
         Text            =   "500"
         Top             =   8520
         Width           =   615
      End
      Begin VB.TextBox Input_Contrast 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   70
         Text            =   "15"
         Top             =   8760
         Width           =   615
      End
      Begin VB.TextBox Input_Time_To_Write 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2640
         TabIndex        =   69
         Text            =   "15"
         Top             =   3600
         Width           =   735
      End
      Begin VB.CommandButton Bt_Display_Points 
         Caption         =   "Show focus points"
         Height          =   375
         Left            =   2160
         TabIndex        =   68
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Bt_Load 
         Caption         =   "Load a task file"
         Height          =   375
         Left            =   2160
         TabIndex        =   67
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Input_Time_To_Click 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2640
         TabIndex        =   66
         Text            =   "5"
         Top             =   2760
         Width           =   735
      End
      Begin VB.Frame Frame_Sentences 
         Height          =   1695
         Left            =   120
         TabIndex        =   53
         Top             =   3960
         Width           =   4095
         Begin VB.TextBox Input_S1 
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   58
            Text            =   "I need to suck"
            Top             =   1320
            Width           =   2625
         End
         Begin VB.TextBox Input_S1 
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   57
            Text            =   "That is my whole life"
            Top             =   1080
            Width           =   2625
         End
         Begin VB.TextBox Input_S1 
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   56
            Text            =   "I am so lucky"
            Top             =   840
            Width           =   2625
         End
         Begin VB.TextBox Input_S1 
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   55
            Text            =   "Thank You so much"
            Top             =   600
            Width           =   2625
         End
         Begin VB.TextBox Input_S1 
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   54
            Text            =   "I love You"
            Top             =   360
            Width           =   2625
         End
         Begin VB.Label Disp_num_sentence 
            Alignment       =   2  'Center
            Caption         =   "5"
            Height          =   255
            Index           =   4
            Left            =   375
            TabIndex        =   65
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Disp_num_sentence 
            Alignment       =   2  'Center
            Caption         =   "4"
            Height          =   255
            Index           =   3
            Left            =   375
            TabIndex        =   64
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Disp_num_sentence 
            Alignment       =   2  'Center
            Caption         =   "3"
            Height          =   255
            Index           =   2
            Left            =   375
            TabIndex        =   63
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Disp_num_sentence 
            Alignment       =   2  'Center
            Caption         =   "2"
            Height          =   255
            Index           =   1
            Left            =   375
            TabIndex        =   62
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Disp_num_sentence 
            Alignment       =   2  'Center
            Caption         =   "1"
            Height          =   255
            Index           =   0
            Left            =   375
            TabIndex        =   61
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label15 
            Caption         =   "Focus point"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Sentence"
            Height          =   255
            Left            =   1560
            TabIndex        =   59
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.CheckBox Chk_Click 
         Caption         =   "The sub must click on the focus point"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   2400
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox Chk_Write 
         Caption         =   "The sub must write a sentence"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   3240
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   120
         TabIndex        =   34
         Top             =   6000
         Width           =   4095
         Begin VB.TextBox Input_nb1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   44
            Text            =   "5"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox Input_Unit1 
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   43
            Text            =   "$"
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Input_nb1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   42
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Input_Unit1 
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   41
            Text            =   "pic(s) of me naked"
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox Input_nb1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   1320
            TabIndex        =   40
            Text            =   "1"
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox Input_Unit1 
            Height          =   285
            Index           =   2
            Left            =   1680
            TabIndex        =   39
            Text            =   ">5 ==>Keylogger"
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox Input_nb1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   1320
            TabIndex        =   38
            Text            =   "1"
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox Input_Unit1 
            Height          =   285
            Index           =   3
            Left            =   1680
            TabIndex        =   37
            Text            =   "day(s) chastity cage"
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox Input_nb1 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   1320
            TabIndex        =   36
            Text            =   "12"
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox Input_Unit1 
            Height          =   285
            Index           =   4
            Left            =   1680
            TabIndex        =   35
            Text            =   "hours web exposure"
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label20 
            Caption         =   "Focus point"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   2655
         End
         Begin VB.Label Disp_num_sentence 
            Alignment       =   2  'Center
            Caption         =   "1"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   49
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Disp_num_sentence 
            Alignment       =   2  'Center
            Caption         =   "2"
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   48
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Disp_num_sentence 
            Alignment       =   2  'Center
            Caption         =   "3"
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   47
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Disp_num_sentence 
            Alignment       =   2  'Center
            Caption         =   "4"
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   46
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Disp_num_sentence 
            Alignment       =   2  'Center
            Caption         =   "5"
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   45
            Top             =   1320
            Width           =   240
         End
      End
      Begin VB.CommandButton Bt_Next 
         Caption         =   "Next"
         Height          =   375
         Left            =   960
         TabIndex        =   33
         Top             =   9120
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "ms"
         Height          =   255
         Left            =   1680
         TabIndex        =   84
         Top             =   8520
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "1 to 127"
         Height          =   255
         Left            =   1680
         TabIndex        =   83
         Top             =   8760
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Contrast:"
         Height          =   375
         Left            =   240
         TabIndex        =   87
         Top             =   8760
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Duration:"
         Height          =   375
         Left            =   240
         TabIndex        =   88
         Top             =   8520
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Diameter:"
         Height          =   375
         Left            =   240
         TabIndex        =   89
         Top             =   8280
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Available time to write a sentence:"
         Height          =   375
         Left            =   120
         TabIndex        =   86
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "sec."
         Height          =   255
         Left            =   3480
         TabIndex        =   85
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Lbl_Accueil 
         Caption         =   "Chose a picture or load an existing task file"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   82
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label14 
         Caption         =   "sec."
         Height          =   255
         Left            =   3480
         TabIndex        =   81
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Alert beam parameters:"
         Height          =   375
         Left            =   240
         TabIndex        =   80
         Top             =   7920
         Width           =   2055
      End
      Begin VB.Label Label19 
         Caption         =   "Actions to be done when the alert beam occurs:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   79
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Label Label16 
         Caption         =   "Available time to click:"
         Height          =   375
         Left            =   960
         TabIndex        =   78
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label21 
         Caption         =   "Penalty in case of failure (either click or sentence):"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   5760
         Width           =   3975
      End
      Begin VB.Label Label22 
         Caption         =   "Click on the picture to create focus points"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   76
         Top             =   1080
         Width           =   4335
      End
   End
   Begin VB.Timer Timer_write 
      Enabled         =   0   'False
      Left            =   6480
      Top             =   3360
   End
   Begin VB.Timer Timer_Mouse 
      Enabled         =   0   'False
      Left            =   4920
      Top             =   3480
   End
   Begin VB.Timer Timer_duration 
      Interval        =   60000
      Left            =   10680
      Top             =   3120
   End
   Begin VB.Timer Timer_Insulte 
      Enabled         =   0   'False
      Left            =   9000
      Top             =   4200
   End
   Begin VB.Timer Timer_React 
      Enabled         =   0   'False
      Left            =   7920
      Top             =   3960
   End
   Begin VB.Timer Timer_Wait 
      Enabled         =   0   'False
      Left            =   6600
      Top             =   4200
   End
   Begin VB.Timer Timer_Remember 
      Enabled         =   0   'False
      Left            =   5760
      Top             =   4200
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame__2 
      Height          =   2415
      Left            =   9120
      TabIndex        =   20
      Top             =   6480
      Width           =   4095
      Begin VB.TextBox Input_PSW 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   30
         Text            =   "Asshole"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Bt_Prev 
         Caption         =   "Back"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Input_Initial_duration 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   24
         Text            =   "15"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Input_Extra_Duration 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   23
         Text            =   "3"
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox Chk_Display_remaining_time 
         Caption         =   "Show the remaining time to Your slave"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.CommandButton Bt_Save 
         Caption         =   "Create task file"
         Height          =   375
         Left            =   2400
         TabIndex        =   21
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Duration"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Extra duration for each failure"
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "min."
         Height          =   255
         Left            =   1080
         TabIndex        =   26
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "min."
         Height          =   375
         Left            =   2400
         TabIndex        =   25
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Label Lbl_test_width 
      AutoSize        =   -1  'True
      Caption         =   "Test Width"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10680
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label Lbl_Remember 
      Caption         =   "It shows me what use of my lips and tongue You have in mind"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   10200
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "F_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
      ' Main routine to dimension variables, retrieve cursor position,
      ' and display coordinates
      Sub Get_Cursor_Pos()


' Dimension the variable that will hold the x and y cursor positions
Dim Hold As POINTAPI

' Place the cursor positions in variable Hold
GetCursorPos Hold

' Display the cursor position coordinates
MsgBox "X Position is : " & Hold.X_Pos & Chr(10) & _
   "Y Position is : " & Hold.Y_Pos
End Sub

' Routine to set cursor position
Sub Set_Cursor_Pos()
Dim X As Long
Dim Y As Long
' Looping routine that positions the cursor
   For X = 1 To 480 Step 20
      SetCursorPos X, X
      For Y = 1 To 40000: Next
   Next X
End Sub

Private Sub Cercle(CentreX As Single, CentreY As Single, Rayon As Single, contrast As Integer)

Dim Angle As Single, X As Single, Y As Single
Dim X0 As Single, Y0 As Single
Dim CR As Double, CG As Double, CB As Double
Dim PTColor As Double
'Dim Rayon As Long
', CentreX As Long, CentreY As Long
Me.Disp_pic.ScaleMode = vbPixels
'Rayon = 20
'CentreX = 150
'CentreY = 110
Angle = 0
Do While Angle < PI * 2
    X = (Cos(Angle) * Rayon) + CentreX
    Y = (Sin(Angle) * Rayon) + CentreY
    If (((X - X0) * (X - X0)) + ((Y - Y0) * (Y - Y0))) > 2 Then
        X0 = X
        Y0 = Y
        'PTColor = Me.Disp_pic.Point(X, Y)
        'CR = (PTColor And 255) + contrast
        'CG = (PTColor And RGB(0, 255, 0) \ 256) + contrast
        'CB = ((PTColor And RGB(0, 0, 255) \ 256) \ 256) + contrast
       ' If CR > 255 Then CR = CR - (2 * contrast)
       ' If CG > 255 Then CR = CG - (2 * contrast)
       ' If CB > 255 Then CR = CB - (2 * contrast)
       ' Me.Disp_pic.PSet (X, Y), RGB(CR, CG, CB)
        Me.Disp_pic.PSet (X, Y), Me.Disp_pic.Point(X, Y) + RGB(contrast, contrast, contrast)
    End If
    'Changer d'angle
    Angle = Angle + 0.01
Loop
DoEvents

End Sub

Private Sub AlertBeamDuration_Timer()

AlertBeamDuration.Enabled = False
AlertBeamDuration.Interval = 0
Me.Disp_pic.Visible = False
Me.Disp_pic.Visible = True
AlertBeamDuration = False

End Sub

Private Sub Bt_Display_Points_Click()

Dim i As Integer

For i = 1 To nb_t_FocusPoint
    Cercle t_FocusPoint(i).X, t_FocusPoint(i).Y, 10, 50
    Me.Disp_pic.CurrentX = t_FocusPoint(i).X - 5
    Me.Disp_pic.CurrentY = t_FocusPoint(i).Y - 13
    Me.Disp_pic.Print Trim(Str(i))
Next i

End Sub

Private Sub Bt_Load_Click()

Dim filenum As Integer
Dim Bytes2() As Byte
Dim s As String
Dim i As Long
Dim lenfic As Long
Dim TaskFile As String
Dim X As Single
Dim Y As Single
Dim strdata As String

Load_In_Progress = True
Me.Disp_pic.ScaleMode = vbPixels

' Load Task file
'---------------
CommonDialog1.Filter = "Task file (*.tsk)|*.tsk|All files (*.*)|*.*"
CommonDialog1.DefaultExt = "tsk"
CommonDialog1.DialogTitle = "Open File"
CommonDialog1.ShowOpen
TaskFile = CommonDialog1.FileName
filenum = FreeFile
Open TaskFile For Binary As filenum
lenfic = LOF(filenum)
strdata = String(lenfic, " ")
Get filenum, , strdata
Close filenum

Me.Bt_Load.Enabled = False

' De-Encryption
'--------------
s = ""
For i = 1 To Len(strdata)
    s = s & (Chr((Val(Asc(Mid(strdata, i, 1))) Xor 145)))
Next i
' Séparation of parameters and pic
'---------------------------------
strdata = Right(s, Len(s) - 2093 - 560 - 11)
s = Left(s, 2093 + 560 + 11)
' Temporary save pic
'-------------------
filenum = FreeFile
Picfile = remove_extension(TaskFile) & ".jpg"
Open Picfile For Binary As filenum
Put filenum, , strdata
Close filenum
If Not Master Then Me.Disp_pic.Left = Me.Frame_Instructions.Width + 4
Me.Disp_pic.Picture = LoadPicture(Picfile)
F_Main.Width = F_Main.Disp_pic.Width + (F_Main.Width - F_Main.ScaleWidth) + F_Main.Disp_pic.Left
If F_Main.Disp_pic.Height + (F_Main.Height - F_Main.ScaleHeight) > F_Main.Height Then
    F_Main.Height = F_Main.Disp_pic.Height + (F_Main.Height - F_Main.ScaleHeight)
End If
'Kill TaskFile & ".jpg"
' Restore parameters
'-------------------
' Diameter
'---------
Input_Diam_Beam.Text = Trim(Mid(s, 1, 10))
' Alert beam duration
'--------------------
Input_AlertBeam_Duration.Text = Trim(Mid(s, 12, 10))
' Alert beam contrast
'--------------------
Input_Contrast.Text = Trim(Mid(s, 23, 10))
' Time to write (0 if no writing)
'--------------------------------
Input_Time_To_Write.Text = Trim(Mid(s, 34, 10))
If Input_Time_To_Write.Text = "0" Then
    Me.Chk_Write.Value = vbUnchecked
Else
    Me.Chk_Write.Value = vbChecked
End If
' 5 times: Text to write, penalty quantity, penalty unit
'-------------------------------------------------------
For i = 0 To 4
    Input_S1(i).Text = Trim(Mid(s, 45 + (i * 413), 300))
    Input_nb1(i).Text = Trim(Mid(s, 346 + (i * 413), 10))
    Input_Unit1(i).Text = Trim(Mid(s, 357 + (i * 413), 100))
    Disp_Unit(i).Caption = Input_Unit1(i).Text
Next i
' Display remaining time
'-----------------------
If Mid(s, 2110, 1) <> 0 Then
    Chk_Display_remaining_time.Value = vbChecked
Else
    Chk_Display_remaining_time.Value = vbUnchecked
End If
' Password
'---------
Input_PSW.Text = Trim(Mid(s, 2111, 100))
' Initial task duration
'----------------------
Input_Initial_duration.Text = Trim(Mid(s, 2212, 10))
' Extra duration in case of failure
'----------------------------------
Input_Extra_Duration.Text = Trim(Mid(s, 2223, 10))
' 10 times: Focus point coordinates
'----------------------------------
nb_t_FocusPoint = 0
For i = 0 To 9
    X = Val(Trim(Mid(s, 2234 + (i * 42), 20)))
    Y = Val(Trim(Mid(s, 2234 + (i * 42) + 21, 20)))
    If X <> 0 Then
        nb_t_FocusPoint = nb_t_FocusPoint + 1
        ReDim Preserve t_FocusPoint(nb_t_FocusPoint)
        t_FocusPoint(nb_t_FocusPoint).X = X
        t_FocusPoint(nb_t_FocusPoint).Y = Y
    Else
        Exit For
    End If
Next i
' Timout for click
'-----------------
Input_Time_To_Click.Text = Trim(Mid(s, 2654, 10))
If Input_Time_To_Click.Text = "0" Then
    Me.Chk_Click.Value = vbUnchecked
Else
    Me.Chk_Click.Value = vbChecked
End If

If Not Master Then
    Me.Lbl_S_Accueil.AutoSize = False
    Me.Lbl_S_Accueil.Width = Me.Disp_pic.Left - 100
    Me.Lbl_S_Accueil.Height = 800
    Me.Lbl_S_Accueil.FontSize = 10
    Me.Lbl_S_Accueil.Caption = "Do you need to learn how to proceed?"
    Me.Bt_S_Load.Visible = False
    Me.Bt_Yes_Need.Visible = True
    Me.Bt_No_Need.Visible = True
    Me.Bt_Yes_Need.Left = 750
    Me.Bt_No_Need.Left = Me.Bt_Yes_Need.Left + Me.Bt_Yes_Need.Width + 100
    Me.Bt_No_Need.SetFocus
    DoEvents
    F_Width = Me.Width
    F_Height = Me.Height
    Me.Width = 4700
    Me.Height = 2500
End If

DoEvents
Load_In_Progress = False

End Sub

Private Sub Bt_Next_Click()

Me.Frame__1.Visible = False
Me.Frame__2.Visible = True

End Sub

Public Sub Bt_No_Need_Click()

Dim i As Integer

Me.Bt_No_Need.Visible = False
Me.Bt_Yes_Need.Visible = False
Me.Lbl_S_Accueil.Height = 600
Me.Lbl_S_Accueil.Visible = False
Me.Input_Slave.Top = Me.Lbl_Remember.Top + Me.Lbl_Remember.Height + 200
Me.Input_Slave.Left = 0
Me.Input_Slave.Visible = True
Me.Width = F_Width
Me.Height = F_Height
Me.Disp_pic.Visible = True
DoEvents

'Me.Lbl_Penalties.Left = 100
'Me.Lbl_Penalties.Top = Me.Input_Slave.Top + Me.Input_Slave.Height + 200
'Me.Lbl_Penalties.Visible = True
'Me.Lbl_Faults.Left = Me.Lbl_Penalties.Left + Me.Lbl_Penalties.Width + 100
'Me.Lbl_Faults.Top = Me.Lbl_Penalties.Top
'Me.Lbl_Faults.Visible = True


'For i = 0 To nb_t_FocusPoint - 1
'    Me.Input_S1(i).BackColor = RGB(248, 248, 248)
'    Me.Input_nb1(i).BackColor = RGB(248, 248, 248)
'    Me.Input_Unit1(i).BackColor = RGB(248, 248, 248)
'Next i
'For i = nb_t_FocusPoint To 4
'    Me.Disp_num_sentence(i).Visible = False
'    Me.Input_S1(i).Visible = False
'    Me.Input_nb1(i).Visible = False
'    Me.Input_Unit1(i).Visible = False
'    Me.Disp_Tax(i).Visible = False
'    Me.Disp_Unit(i).Visible = False
'Next i

Me.Frame_Sentences.Height = Me.Input_S1(nb_t_FocusPoint - 1).Top + Me.Input_S1(nb_t_FocusPoint - 1).Height + 100
Me.Frame_Results.Height = Me.Disp_Tax(nb_t_FocusPoint - 1).Top + Me.Disp_Tax(nb_t_FocusPoint - 1).Height + 100
Me.Frame_Results.Left = 0
Me.Frame_Results.Top = 0
Me.Frame_Results.Visible = True
Me.Frame_Sentences.Visible = False
Me.Frame_Sentences.Enabled = False
Me.Lbl_Remember.Top = Me.Frame_Sentences.Top + Me.Frame_Sentences.Height + 100
Me.Lbl_Remember.Left = 100
Me.Lbl_Remember.Visible = False

Me.Lbl_Remaining_Time_Caption.Left = 30
Me.Lbl_Remaining_Time_Caption.Top = Me.Frame_Results.Top + Me.Frame_Results.Height + 200
Me.Lbl_Remaining_Time_Caption.Visible = True
Me.Lbl_Remaining_Time.Left = Me.Lbl_Remaining_Time_Caption.Left + Me.Lbl_Remaining_Time_Caption.Width + 100
Me.Lbl_Remaining_Time.Top = Me.Lbl_Remaining_Time_Caption.Top
Me.Lbl_Remaining_Time.Caption = Me.Input_Initial_duration.Text
Me.Lbl_Remaining_Time.Visible = True

Me.Input_Slave.Top = Me.Lbl_Remaining_Time.Top + Me.Lbl_Remaining_Time.Height + 200

Me.Frame_Mouse.Left = 0
Me.Frame_Mouse.Top = Me.Input_Slave.Top + Me.Input_Slave.Height + 200
Me.Frame_Mouse.Visible = True

Me.Frame_Instructions.Left = 0
Me.Frame_Instructions.Top = Me.Frame_Mouse.Top + Me.Frame_Mouse.Height + 200
Me.Frame_Instructions.Visible = True

If Not Master Then
    If F_Main.ScaleHeight < Me.Frame_Instructions.Top + Me.Frame_Instructions.Height Then
        F_Main.Height = Me.Frame_Instructions.Top + Me.Frame_Instructions.Height + (F_Main.Height - F_Main.ScaleHeight)
    End If
End If

Set_remember

End Sub

Private Sub Bt_Prev_Click()

Me.Frame__1.Visible = True
Me.Frame__2.Visible = False

End Sub

Private Sub Bt_Reset_points_Click()

nb_t_FocusPoint = 0
Me.Disp_pic.Visible = False
Me.Disp_pic.Visible = True

End Sub

Private Sub Bt_S_Load_Click()

Bt_Load_Click

End Sub

Private Sub Bt_Save_Click()

Dim filenum As Integer
Dim Bytes2() As Byte
Dim s As String
Dim S1 As String
Dim i As Long
Dim lenfic As Long
Dim strdata As String

If Picfile = "" Then Exit Sub

' Load pic file
'--------------
filenum = FreeFile
Open Picfile For Binary As filenum
lenfic = LOF(filenum)
strdata = String(lenfic, " ")
Get filenum, , strdata
Close filenum

' Compilation of parameters
'--------------------------
' Diameter
'---------
s = Pad(Input_Diam_Beam.Text, 10)
' Alert beam duration
'--------------------
s = s & Pad(Input_AlertBeam_Duration.Text, 10)
' Alert beam contrast
'--------------------
s = s & Pad(Input_Contrast.Text, 10)
' Time to write (0 if no writing)
'--------------------------------
If Me.Chk_Write.Value = vbChecked Then
    s = s & Pad(Input_Time_To_Write.Text, 10)
Else
    s = s & Pad("0", 10)
End If
' 5 times: Text to write, penalty quantity, penalty unit
'-------------------------------------------------------
For i = 0 To 4
    s = s & Pad(Input_S1(i).Text, 300)
    s = s & Pad(Input_nb1(i).Text, 10)
    s = s & Pad(Input_Unit1(i).Text, 100)
Next i
' Display remaining time
'-----------------------
If Chk_Display_remaining_time.Value = vbChecked Then
    s = s & "1"
Else
    s = s & "0"
End If
' Password
'---------
s = s & Pad(Input_PSW.Text, 100)
' Initial task duration
'----------------------
s = s & Pad(Input_Initial_duration.Text, 10)
' Extra duration in case of failure
'----------------------------------
s = s & Pad(Input_Extra_Duration.Text, 10)
' 10 times: Focus point coordinates
'----------------------------------
For i = 1 To 10
    If i <= nb_t_FocusPoint Then
        s = s & Pad(Trim(Str(t_FocusPoint(i).X)), 20)
        s = s & Pad(Trim(Str(t_FocusPoint(i).Y)), 20)
    Else
        s = s & Pad(" ", 20)
        s = s & Pad(" ", 20)
    End If
Next i
' Timout for click
'-----------------
If Me.Chk_Click.Value = vbChecked Then
    s = s & Pad(Input_Time_To_Click.Text, 10)
Else
    s = s & Pad("0", 10)
End If
' Concatenation of parameters and pic
'------------------------------------
s = s & strdata
' Encryption
'-----------
S1 = ""
For i = 1 To Len(s)
    S1 = S1 & (Chr((Val(Asc(Mid(s, i, 1))) Xor 145)))
Next i
' Create task file
'-----------------
filenum = FreeFile
Open remove_extension(Picfile) & ".tsk" For Binary As filenum
Put filenum, , S1
Close filenum

End Sub

Private Sub Bt_Select_Pic_Click()

Load_In_Progress = True
CommonDialog1.Filter = "Pic (*.jpg)|*.jpeg|All files (*.*)|*.*"
CommonDialog1.DefaultExt = "jpg"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.ShowOpen
Picfile = CommonDialog1.FileName
F_Main.Disp_pic.Picture = LoadPicture(Picfile)
'F_Main.Width = F_Main.Disp_pic.Width + (F_Main.Width - F_Main.ScaleWidth) + F_Main.Disp_pic.Left
'F_Main.Height = F_Main.Disp_pic.Height + (F_Main.Height - F_Main.ScaleHeight)
F_Main.Width = F_Main.Disp_pic.Width + (F_Main.Width - F_Main.ScaleWidth) + F_Main.Disp_pic.Left
If F_Main.Disp_pic.Height + (F_Main.Height - F_Main.ScaleHeight) > F_Main.Height Then
    F_Main.Height = F_Main.Disp_pic.Height + (F_Main.Height - F_Main.ScaleHeight)
End If
DoEvents
Me.Disp_pic.ScaleMode = vbPixels
Load_In_Progress = False

End Sub

Private Sub Bt_Test_Alert_Beam_Click()

Dim i As Single

Me.Disp_pic.Visible = False
Me.Disp_pic.Visible = True
DoEvents
For i = 1 To Val(Me.Input_Diam_Beam.Text) Step 2
    Cercle t_FocusPoint(nb_t_FocusPoint).X, t_FocusPoint(nb_t_FocusPoint).Y, i, _
                Val(Me.Input_Contrast.Text)
    DoEvents
'    Cercle t_FocusPoint(nb_t_FocusPoint).X, t_FocusPoint(nb_t_FocusPoint).Y, i + 1, _
'                RGB(16 + Val(Me.Input_R.Text), 16 + Val(Me.Input_V.Text), 16 + Val(Me.Input_B.Text))
'    DoEvents
Next i
'Cercle t_FocusPoint(nb_t_FocusPoint).X, t_FocusPoint(nb_t_FocusPoint).Y, 2 * Val(Me.Input_Diam_Beam.Text), _
'            RGB(Val(Me.Input_R.Text), Val(Me.Input_V.Text), Val(Me.Input_B.Text))
AlertBeamDuration.Interval = Val(Input_AlertBeam_Duration.Text)
AlertBeamDuration.Enabled = True

End Sub

Private Sub Bt_Yes_Need_Click()

'F_Help.Show 1
F_Main.Bt_No_Need_Click

End Sub

Private Sub Chk_Click_Click()

If Chk_Click.Value = vbChecked Then
    Input_Time_To_Click.Enabled = True
    Label16.Enabled = True
    Label14.Enabled = True
Else
    If Chk_Write.Value = vbUnchecked Then Chk_Write.Value = vbChecked
    Input_Time_To_Click.Enabled = False
    Label16.Enabled = False
    Label14.Enabled = False
End If

End Sub

Private Sub Chk_Write_Click()

If Chk_Write.Value = vbChecked Then
    Input_Time_To_Write.Enabled = True
    Label8.Enabled = True
    Label9.Enabled = True
    Frame_Sentences.Enabled = True
Else
    If Chk_Click.Value = vbUnchecked Then Chk_Click.Value = vbChecked
    Input_Time_To_Write.Enabled = False
    Label8.Enabled = False
    Label9.Enabled = False
    Frame_Sentences.Enabled = False
End If

End Sub

Private Sub Disp_pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Delta As Single

If Load_In_Progress Then
    Exit Sub
End If
If Master Then
    nb_t_FocusPoint = nb_t_FocusPoint + 1
    ReDim Preserve t_FocusPoint(nb_t_FocusPoint)
    t_FocusPoint(nb_t_FocusPoint).X = X
    t_FocusPoint(nb_t_FocusPoint).Y = Y
    
    Cercle X, Y, 10, 128
    Me.Disp_pic.CurrentX = X - 5
    Me.Disp_pic.CurrentY = Y - 13
    Me.Disp_pic.Print Trim(Str(nb_t_FocusPoint))
Else
    Delta = ((t_FocusPoint(Current_Focus).X - X) ^ 2) + ((t_FocusPoint(Current_Focus).Y - Y) ^ 2)
    If Delta < 100 Then
        Me.Timer_React.Enabled = False
        Me.Timer_React.Interval = 0
        If Me.Chk_Write.Value = vbChecked Then
            Me.Timer_write.Interval = 1000 * Me.Input_Time_To_Write.Text
            Me.Timer_write.Enabled = True
            Me.Input_Slave.SetFocus
            Input_Allowed = True
        Else
            Park_Mouse
            Lbl_Failure.Caption = "Lol!"
            Lbl_Failure.Left = ((Me.Disp_pic.Width / Screen.TwipsPerPixelX) - Me.Lbl_Failure.Width) / 2
            Lbl_Failure.Top = ((Me.Disp_pic.Height / Screen.TwipsPerPixelY) - Me.Lbl_Failure.Height) / 2
            Lbl_Failure.Visible = True
            Lol = True
            Timer_Insulte.Interval = 3000
            Timer_Insulte.Enabled = True
        End If
        Click_Allowed = False
    Else
        Failure "You need to click on the right place, asshole!"
    End If
End If

End Sub

Public Function Pad(s As String, NbChar As Integer) As String

Dim i As Integer

Pad = s
For i = Len(s) To NbChar
    Pad = Pad & " "
Next i

End Function

Private Sub Form_Initialize()

Dim i As Integer

nb_t_FocusPoint = 0
Me.Left = 0
Me.Top = 0
Me.Width = 4560

If Master Then
    Me.Disp_pic.FontItalic = False
    Me.Disp_pic.ForeColor = RGB(0, 0, 0)
    Me.Frame__2.Left = 120
    Me.Frame__2.Top = 500
Else
    Me.Frame__1.Visible = False
    Me.Frame__2.Visible = False
    Me.Frame__3.Visible = True
    Me.Frame__3.Left = Me.Frame__1.Left
    Me.Frame__3.Top = Me.Frame__1.Top
    Me.Frame__3.Width = Me.Frame__1.Width
'    Me.Frame__3.Width = Me.Frame_Instructions.Width
    Me.Frame__3.Height = Me.Frame__1.Width
    
    Me.Disp_pic.FontItalic = True
    Me.Disp_pic.ForeColor = RGB(0, 0, 0)
    Me.Label1.Visible = False
    Me.Label10.Visible = False
    Me.Label11.Visible = False
    Me.Label12.Visible = False
    Me.Label13.Visible = False
    Me.Label14.Visible = False
    Me.Label16.Visible = False
    Me.Input_Time_To_Click.Visible = False
    Me.Label4.Visible = False
    Me.Label2.Visible = False
    Me.Label3.Visible = False
    Me.Label6.Visible = False
    Me.Label7.Visible = False
    Me.Label8.Visible = False
    Me.Label9.Visible = False
    Me.Input_AlertBeam_Duration.Visible = False
    Me.Input_Contrast.Visible = False
    Me.Input_Diam_Beam.Visible = False
    Me.Input_Extra_Duration.Visible = False
    Me.Input_Initial_duration.Visible = False
    Me.Input_PSW.Visible = False
    Me.Frame_Sentences.Visible = False
    Me.Frame_Sentences.Top = -60
'    Me.Label5.Visible = False
'    For i = 0 To 4
'        Me.Input_S1(i).Visible = False
'    Next i
    Me.Input_Time_To_Write.Visible = False
    Me.Bt_Display_Points.Visible = False
    Me.Bt_Reset_points.Visible = False
    Me.Bt_Save.Visible = False
    Me.Bt_Select_Pic.Visible = False
    Me.Bt_Test_Alert_Beam.Visible = False
    Me.Chk_Display_remaining_time.Visible = False
    Me.Disp_pic.Visible = False
    Me.Lbl_S_Accueil.Caption = "Load your task, slave! Hurry up!"
    Me.Lbl_S_Accueil.FontSize = 18
    Me.Lbl_S_Accueil.AutoSize = True
    'Me.Bt_Load.Left = Me.Lbl_Accueil.Left + (Me.Lbl_Accueil.Width - Me.Bt_Load.Width) / 2
    'Me.Bt_Load.Top = Me.Lbl_Accueil.Top + Me.Lbl_Accueil.Height + 100
    Me.Height = Me.Bt_S_Load.Top + Me.Bt_S_Load.Height + 700
    Me.Width = Me.Lbl_S_Accueil.Left + Me.Lbl_S_Accueil.Width + 400
    
End If

End Sub


Public Function Get_Sentence() As String

Dim s() As String
Dim nb_s As Integer
Dim i As Integer

nb_s = 0
For i = 0 To 4
    If Me.Input_S1(i).Text <> "" Then
        nb_s = nb_s + 1
        ReDim Preserve s(nb_s)
        s(nb_s) = Me.Input_S1(i).Text
    End If
Next i

Randomize
i = Int(nb_s * Rnd()) + 1
If i > nb_s Then i = nb_s
If i < 1 Then i = 1
Get_Sentence = s(i)

End Function





Private Sub Input_Slave_Change()

If Me.Input_Slave.Text = "" Then
    Len_Input_Slave = 0
    Exit Sub
End If

If Not Input_Allowed Then
    If Click_Allowed Then
        Failure "You had to click, slave!"
    Else
        Failure "Wait for my command, slave!"
    End If
Else
    If Abs(Len(Me.Input_Slave.Text) - Len_Input_Slave) > 2 Then
        Failure "Do not cheat on me, asshole!"
    Else
        Len_Input_Slave = Len(Me.Input_Slave.Text)
    End If
End If

End Sub

Private Sub Input_Slave_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    If Len(Me.Input_Slave.Text) > 2 Then
        Me.Input_Slave.Text = Left(Me.Input_Slave.Text, Len(Me.Input_Slave.Text) - 2)
    End If
    If Me.Input_Slave.Text = Current_Sentence Then
        Me.Input_Slave.Text = ""
        Me.Input_Slave.SetFocus
        Timer_write.Enabled = False
        Timer_write.Interval = 0
        Input_Allowed = False
        
        Lbl_Failure.Caption = "Lol!"
        Lbl_Failure.Left = ((Me.Disp_pic.Width / Screen.TwipsPerPixelX) - Me.Lbl_Failure.Width) / 2
        Lbl_Failure.Top = ((Me.Disp_pic.Height / Screen.TwipsPerPixelY) - Me.Lbl_Failure.Height) / 2
        Lbl_Failure.Visible = True
        Lol = True
        Timer_Insulte.Interval = 3000
        Timer_Insulte.Enabled = True
    Else
        Failure "Wrong sentence, asshole!"
    End If
End If

Park_Mouse

End Sub





Private Sub Timer_duration_Timer()

Me.Lbl_Remaining_Time.Caption = Trim(Str(Val(Me.Lbl_Remaining_Time.Caption) - 1))
If Val(Me.Lbl_Remaining_Time.Caption) < 1 Then
    CreateReport
    Me.Caption = "Congratulation"
End If

End Sub

Private Sub Timer_Insulte_Timer()

Timer_Insulte.Enabled = False
Timer_Insulte.Interval = 0
Lbl_Failure.Visible = False
If Lol Then
    Lol = False
    If RandomChange Then
        Set_remember
    Else
        Set_Random_Alert
    End If
    Park_Mouse
Else
    Set_remember
    'Set_Random_Alert
    Park_Mouse
End If

End Sub

Private Sub Timer_Mouse_Timer()

Dim Hold As POINTAPI
Dim Delta As Single

' Place the cursor positions in variable Hold
GetCursorPos Hold

Delta = ((Hold.X_Pos - X_curs_prison) ^ 2) + ((Hold.Y_Pos - Y_curs_prison) ^ 2)
' Display the cursor position coordinates

If Delta > 100 Then
    Me.Timer_Mouse.Enabled = False
    Me.Timer_Mouse.Interval = 0
    Failure "Don't move your mouse, asshole!"
End If

End Sub

Private Sub Timer_React_Timer()

Timer_React.Enabled = False
Timer_React.Interval = 0
Failure "Are you sleeping asshole? Keep focused!"
Click_Allowed = False
Input_Allowed = False
Set_Random_Alert

End Sub

Private Sub Timer_Remember_Timer()

Timer_Remember.Enabled = False
Timer_Remember.Interval = 0
Me.Lbl_Remember.Visible = False
Me.Picture_Consigne.Visible = False
Me.Disp_pic.Refresh
Me.Lbl_Accueil.Caption = "Now focus!!!!!!!!!!!!!"
Me.Input_Slave.SetFocus
Set_Random_Alert

End Sub

Public Sub Set_remember()

'Me.Lbl_Accueil.Caption = "Remember slave  :-)" & vbCrLf & "(Sentence + focus point)"

Show_Rnd_Point
'Current_Sentence = Get_Sentence
Current_Sentence = Me.Input_S1(Current_Focus - 1).Text
'Me.Lbl_Accueil.Caption = "This is your thought for now:"
Me.Lbl_Remember.Caption = Current_Sentence
'Me.Lbl_Remember.Visible = True

Timer_Remember.Enabled = True
Timer_Remember.Interval = 4000


End Sub

Public Sub Show_Rnd_Point()

Dim i As Integer
Dim s As String

If nb_t_FocusPoint < 1 Then Exit Sub

Randomize
i = Int(nb_t_FocusPoint * Rnd()) + 1
If i > nb_t_FocusPoint Then i = nb_t_FocusPoint
If i < 1 Then i = 1

Current_Focus = i
Me.Picture_Consigne.Left = Me.Disp_pic.Left + (Screen.TwipsPerPixelX * t_FocusPoint(i).X) - Me.Picture_Consigne.Width
Me.Picture_Consigne.Top = Me.Disp_pic.Top + (Screen.TwipsPerPixelY * t_FocusPoint(i).Y) - Me.Picture_Consigne.Height
Me.Picture_Consigne.Visible = True
'Cercle t_FocusPoint(i).X, t_FocusPoint(i).Y, 10, 127
'Cercle t_FocusPoint(i).X, t_FocusPoint(i).Y, 12, 127
'Cercle t_FocusPoint(i).X, t_FocusPoint(i).Y, 14, 127

Current_Sentence = Me.Input_S1(Current_Focus - 1).Text


If Me.Chk_Click.Value = vbChecked Then
    s = "- Click on this exact location on the picture"
    If Me.Chk_Write.Value = vbChecked Then
        s = s & vbCrLf & "- Then write this sentence:"
        Me.Lbl_Instruction.Caption = s
        Me.Lbl_To_Be_Written.Caption = Current_Sentence
    Else
        Me.Lbl_Instruction.Caption = s
        Me.Lbl_To_Be_Written.Caption = ""
    End If
Else
    If Me.Chk_Write.Value = vbChecked Then
        s = "- Write this sentence:"
        Me.Lbl_Instruction.Caption = s
        Me.Lbl_To_Be_Written.Caption = Current_Sentence
    End If
End If

'If Me.Chk_Click.Value = vbChecked Then
'    s = "Click"
'    If Me.Chk_Write.Value = vbChecked Then
'        s = s & " + """ & Current_Sentence & """"
'    End If
'Else
'    If Me.Chk_Write.Value = vbChecked Then
'        s = """" & Current_Sentence & """"
'    End If
'End If


'Me.Lbl_test_width.Caption = s
'Me.Disp_pic.CurrentX = t_FocusPoint(i).X - (Me.Lbl_test_width.Width / (2 * Screen.TwipsPerPixelX))
'Me.Disp_pic.CurrentY = t_FocusPoint(i).Y + 20
'Me.Disp_pic.Print Me.Lbl_test_width.Caption

Park_Mouse

End Sub

Public Sub Show_Alert(i_Focus_point As Integer)

Dim i As Single

Me.Disp_pic.Visible = False
Me.Disp_pic.Visible = True
DoEvents
For i = 1 To Val(Me.Input_Diam_Beam.Text) Step 2
    Cercle t_FocusPoint(i_Focus_point).X, t_FocusPoint(i_Focus_point).Y, i, _
                Val(Me.Input_Contrast.Text)
    DoEvents
'    Cercle t_FocusPoint(nb_t_FocusPoint).X, t_FocusPoint(nb_t_FocusPoint).Y, i + 1, _
'                RGB(16 + Val(Me.Input_R.Text), 16 + Val(Me.Input_V.Text), 16 + Val(Me.Input_B.Text))
'    DoEvents
Next i
'Cercle t_FocusPoint(nb_t_FocusPoint).X, t_FocusPoint(nb_t_FocusPoint).Y, 2 * Val(Me.Input_Diam_Beam.Text), _
'            RGB(Val(Me.Input_R.Text), Val(Me.Input_V.Text), Val(Me.Input_B.Text))
AlertBeamDuration.Interval = Val(Input_AlertBeam_Duration.Text)
AlertBeamDuration.Enabled = True

End Sub

Public Sub Set_Random_Alert()

Dim i As Single
Randomize
i = CInt((30000 - 2000 + 1) * Rnd()) + 2000
If i < 30000 Then i = 30000
Timer_Wait.Interval = i
Timer_Wait.Enabled = True

End Sub

Private Sub Timer_Wait_Timer()

Timer_Wait.Enabled = False
Timer_Wait.Interval = 0
Show_Alert Current_Focus
If Me.Chk_Click.Value = vbChecked Then
    Click_Allowed = True
    Timer_React.Interval = Val(1000 * Me.Input_Time_To_Click.Text)
    Timer_React.Enabled = True
    Me.Timer_Mouse.Interval = 0
    Me.Timer_Mouse.Enabled = False
ElseIf Me.Chk_Write.Value = vbChecked Then
    Input_Allowed = True
    Timer_write.Interval = Val(1000 * Me.Input_Time_To_Write.Text)
    Timer_write.Enabled = True
End If

End Sub

Public Sub Failure(Text As String)

Me.Timer_Wait.Enabled = False
Me.Timer_Wait.Interval = 0
Me.Timer_React.Enabled = False
Me.Timer_React.Interval = 0
Me.Timer_write.Enabled = False
Me.Timer_write.Interval = 0
Me.Disp_Tax(Current_Focus - 1).Caption = Trim(Str(Val(Me.Disp_Tax(Current_Focus - 1).Caption) + Val(Me.Input_nb1(Current_Focus - 1))))
Me.Input_Slave.Text = ""
Me.Lbl_Remaining_Time.Caption = Trim(Str(Val(Me.Lbl_Remaining_Time.Caption) + Val(Me.Input_Extra_Duration.Text)))
'Lbl_Failure.Caption = Text & vbCrLf & Trim(Str(Me.Lbl_Faults.Caption))
Lbl_Failure.Caption = Text
Lbl_Failure.Left = ((Me.Disp_pic.Width / Screen.TwipsPerPixelX) - Me.Lbl_Failure.Width) / 2
Lbl_Failure.Top = ((Me.Disp_pic.Height / Screen.TwipsPerPixelY) - Me.Lbl_Failure.Height) / 2
Lbl_Failure.Visible = True
Input_Allowed = False
Click_Allowed = False
Timer_Insulte.Interval = 3000
Timer_Insulte.Enabled = True

End Sub

Public Function RandomChange() As Boolean

Dim i As Integer

Randomize
i = CInt((1000 - 1 + 1) * Rnd()) + 1
If i < 600 Then
    RandomChange = True
Else
    RandomChange = False
End If

End Function

Public Sub Park_Mouse()

SetCursorPos 0, 0
X_curs_prison = (Me.Left + Me.Frame_Mouse.Left + (Me.Frame_Mouse.Width / 2) + 50) / Screen.TwipsPerPixelX
Y_curs_prison = (Me.Top + Me.Frame_Mouse.Top + (Me.Frame_Mouse.Height / 2) + 350) / Screen.TwipsPerPixelY
SetCursorPos X_curs_prison, Y_curs_prison
Me.Timer_Mouse.Interval = 200
Me.Timer_Mouse.Enabled = True

End Sub

Private Sub Timer_write_Timer()

Timer_write.Enabled = False
Timer_write.Interval = 0
Failure "Are you sleeping asshole? Keep focused!"
Click_Allowed = False
Input_Allowed = False
Set_Random_Alert

End Sub

Public Function remove_extension(s As String) As String

Dim i As Integer

For i = Len(s) To 1 Step -1
    If Mid(s, i, 1) = "." Then Exit For
Next i
If i > 1 Then
    remove_extension = Left(s, i - 1)
Else
    remove_extension = s
End If

End Function

Public Sub CreateReport()

Dim filenum As Integer
Dim Bytes2() As Byte
Dim s As String
Dim S1 As String
Dim i As Long
Dim lenfic As Long
Dim strdata As String

'If Picfile = "" Then Exit Sub
'
'' Load pic file
''--------------
'filenum = FreeFile
'Open Picfile For Binary As filenum
'lenfic = LOF(filenum)
'strdata = String(lenfic, " ")
'Get filenum, , strdata
'Close filenum

' Compilation of parameters
'--------------------------
's = Pad(Input_Diam_Beam.Text, 10)
's = s & Pad(Input_AlertBeam_Duration.Text, 10)
's = s & Pad(Input_Contrast.Text, 10)
If Me.Chk_Write.Value = vbChecked Then
    s = s & Pad(Input_Time_To_Write.Text, 10)
Else
    s = s & Pad("0", 10)
End If
For i = 0 To 4
    s = s & Pad(Input_S1(i).Text, 300)
    s = s & Pad(Input_nb1(i).Text, 10)
    s = s & Pad(Input_Unit1(i).Text, 100)
Next i
If Chk_Display_remaining_time.Value = vbChecked Then
    s = s & "1"
Else
    s = s & "0"
End If
s = s & Pad(Input_PSW.Text, 100)
s = s & Pad(Input_Initial_duration.Text, 10)
s = s & Pad(Input_Extra_Duration.Text, 10)
For i = 1 To 10
    If i <= nb_t_FocusPoint Then
        s = s & Pad(Trim(Str(t_FocusPoint(i).X)), 20)
        s = s & Pad(Trim(Str(t_FocusPoint(i).Y)), 20)
    Else
        s = s & Pad(" ", 20)
        s = s & Pad(" ", 20)
    End If
Next i
If Me.Chk_Click.Value = vbChecked Then
    s = s & Pad(Input_Time_To_Click.Text, 10)
Else
    s = s & Pad("0", 10)
End If
' Concatenation of parameters and pic
'------------------------------------
s = s & strdata
' Encryption
'-----------
S1 = ""
For i = 1 To Len(s)
    S1 = S1 & (Chr((Val(Asc(Mid(s, i, 1))) Xor 145)))
Next i
' Create task file
'-----------------
filenum = FreeFile
Open remove_extension(Picfile) & ".tsk" For Binary As filenum
Put filenum, , S1
Close filenum

End Sub
