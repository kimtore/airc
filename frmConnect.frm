VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "frmConnect"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Frames 
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   2
      Left            =   2400
      ScaleHeight     =   4335
      ScaleWidth      =   5535
      TabIndex        =   56
      Tag             =   "DCC"
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CheckBox chkDCCSafe 
         Caption         =   "DCC safe mode (use host over IP)"
         Height          =   255
         Left            =   120
         TabIndex        =   134
         ToolTipText     =   "Use recepient's host instead of IP given in DCC CHAT ctcp."
         Top             =   3840
         Width           =   3375
      End
      Begin VB.TextBox txtDCCPorts 
         Height          =   285
         Left            =   3480
         TabIndex        =   115
         Top             =   3480
         Width           =   1935
      End
      Begin VB.CheckBox chkDCCPorts 
         Caption         =   "Use only this port range for DCC"
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   3480
         Width           =   3135
      End
      Begin VB.CheckBox chkPassiveDCC 
         Caption         =   "Use passive DCC chat/send (non-RFC)"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         ToolTipText     =   "Initialize DCCs through firewall"
         Top             =   3120
         Width           =   3735
      End
      Begin VB.TextBox txtIgnoreFiltyper 
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Top             =   1680
         Width           =   3375
      End
      Begin VB.CheckBox chkIgnoreFiltyper 
         Caption         =   "Ignore filetypes:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Which file types to ignore, separated by semicolon"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox chkDCCAutoGet 
         Caption         =   "Automatically accept files"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Recieve every file sent not stopped by the protection"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CheckBox chkPumpDCC 
         Caption         =   "Pump DCC"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Similar to disabling the FTP Nagle algorithm, does not wait for ack before sending next packet"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtSendeBuffer 
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         ToolTipText     =   "How many bytes to keep in the DCC send buffer"
         Top             =   2520
         Width           =   2175
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Left            =   4200
         Max             =   2
         TabIndex        =   18
         Top             =   2520
         Value           =   1
         Width           =   255
      End
      Begin VB.CheckBox chkJoinIgnore 
         Caption         =   "Temporary DCC ignore on join (10 secs)"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Ignore DCC requests in 10 seconds after joining a channel"
         Top             =   960
         Width           =   3855
      End
      Begin VB.CheckBox chkBeskyttVirus 
         Caption         =   "Protect against general type executable viruses"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Ignores potensially dangerous files (shs, vbs, com, vbs, vbe, js, jse, scr, pif, bat)"
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox txtDownDir 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         ToolTipText     =   "Where downloaded files are placed"
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label Label7 
         Caption         =   "bytes"
         Height          =   255
         Left            =   4560
         TabIndex        =   59
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "DCC transfer buffer:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Download folder:"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox Frames 
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   3
      Left            =   2400
      ScaleHeight     =   4335
      ScaleWidth      =   5535
      TabIndex        =   53
      Tag             =   "IP"
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
      Begin VB.OptionButton Option1 
         Caption         =   "Fetch from server"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         ToolTipText     =   "The client will get the IP the server sends instead of the local IP"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.OptionButton optLookupType 
         Caption         =   "Fetch local"
         Height          =   255
         Left            =   480
         TabIndex        =   21
         ToolTipText     =   "The client will get the local IP instead of the one the server sends"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Top             =   120
         Width           =   1575
      End
      Begin VB.CheckBox chkHentIP 
         Caption         =   "Fetch IP on connect"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "If this otion is checked, the client will set the local IP on connect"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Use IP:"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox Frames 
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   5
      Left            =   2400
      ScaleHeight     =   4335
      ScaleWidth      =   5535
      TabIndex        =   60
      Tag             =   "Display"
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CheckBox chkShowNicklist 
         Caption         =   "Show channel nicklist"
         Height          =   255
         Left            =   120
         TabIndex        =   99
         ToolTipText     =   "Show nicklist in channel windows"
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CheckBox chkColorActivity 
         Caption         =   "Color window on activity"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         ToolTipText     =   "Color text in window list on activity"
         Top             =   3360
         Width           =   2415
      End
      Begin VB.CheckBox chkFlashAny 
         Caption         =   "Flash window on any activity"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         ToolTipText     =   "Activate window on any activity"
         Top             =   3000
         Width           =   2775
      End
      Begin VB.CheckBox chkFlashNew 
         Caption         =   "Flash window on new message"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         ToolTipText     =   "Activate window on incoming messages"
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Frame Frame7 
         Height          =   1695
         Left            =   120
         TabIndex        =   62
         Top             =   720
         Width           =   5295
         Begin VB.CheckBox chkStripA 
            Caption         =   "Strip all codes"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            ToolTipText     =   "Remove all attributes from text"
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CheckBox chkStripU 
            Caption         =   "Strip underline codes"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            ToolTipText     =   "Remove underline attributes from text"
            Top             =   840
            Width           =   2175
         End
         Begin VB.CheckBox chkStripB 
            Caption         =   "Strip bold codes"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            ToolTipText     =   "Remove bold attributes from text"
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox chkStripC 
            Caption         =   "Strip color codes"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            ToolTipText     =   "Remove color from text"
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox chkStrip 
            Caption         =   "Strip control codes"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            ToolTipText     =   "Prevent control codes from changing text attributes"
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.TextBox txtTimestamp 
         Height          =   285
         Left            =   2040
         TabIndex        =   29
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   "Timestamp format:"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.PictureBox Frames 
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   6
      Left            =   2400
      ScaleHeight     =   4335
      ScaleWidth      =   5535
      TabIndex        =   64
      Tag             =   "Colors"
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CheckBox chkmIRCColors 
         Caption         =   "Interpret color codes as mIRC colors instead of ANSI"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   3360
         Width           =   4935
      End
      Begin VB.CommandButton cmdColormIRC 
         Caption         =   "Import from mIRC..."
         Height          =   375
         Left            =   3000
         TabIndex        =   40
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdColorLoad 
         Caption         =   "Load..."
         Height          =   375
         Left            =   1560
         TabIndex        =   39
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton cmdColorSave 
         Caption         =   "Save..."
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   3840
         Width           =   1335
      End
      Begin VB.PictureBox picChooseSec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   85
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox picChooseStd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   83
         Top             =   2520
         Width           =   495
      End
      Begin VB.PictureBox picChooseBrand 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   80
         Top             =   2880
         Width           =   495
      End
      Begin VB.PictureBox picChooseURL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   79
         Top             =   2520
         Width           =   495
      End
      Begin VB.CommandButton cmdChooseFont 
         Caption         =   "Choose font..."
         Height          =   375
         Left            =   3480
         TabIndex        =   35
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdResetColors 
         Caption         =   "Reset"
         Height          =   375
         Left            =   3480
         TabIndex        =   36
         Top             =   1440
         Width           =   1935
      End
      Begin VB.PictureBox picColBG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2145
         ScaleWidth      =   3105
         TabIndex        =   65
         Top             =   120
         Width           =   3135
         Begin VB.Label lblColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Notice text"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   1680
            TabIndex        =   82
            Tag             =   "Notice"
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label lblColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Own text"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   1680
            TabIndex        =   76
            Tag             =   "Own"
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Normal text"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   1680
            TabIndex        =   75
            Tag             =   "Normal"
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Topic text"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   1680
            TabIndex        =   74
            Tag             =   "Topic"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Status text"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   73
            Tag             =   "Error"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Action text"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   72
            Tag             =   "Action"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label lblColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Mode text"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   71
            Tag             =   "Mode"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Kick text"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   70
            Tag             =   "Kick"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Nick text"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   69
            Tag             =   "Nick"
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Quit text"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   68
            Tag             =   "Quit"
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Part text"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   67
            Tag             =   "Part"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Join text"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   66
            Tag             =   "Join"
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   120
         X2              =   5400
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   120
         X2              =   5400
         Y1              =   2410
         Y2              =   2410
      End
      Begin VB.Label Label8 
         Caption         =   "Secondary color:"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   86
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Standard color:"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   84
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Brand color:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   81
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "URL color:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   78
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblFontPreview 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3480
         TabIndex        =   77
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.PictureBox Frames 
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   1
      Left            =   2400
      ScaleHeight     =   4335
      ScaleWidth      =   5535
      TabIndex        =   52
      Tag             =   "Cloaking"
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Frame Frame2 
         Caption         =   "Cloak options"
         Height          =   2175
         Left            =   120
         TabIndex        =   106
         Top             =   600
         Width           =   5295
         Begin VB.CommandButton cmdCloakReset 
            Caption         =   "Reset to default"
            Height          =   345
            Left            =   3240
            TabIndex        =   113
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton optCloak 
            Caption         =   "Normal reply"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   112
            Top             =   780
            Width           =   1935
         End
         Begin VB.CheckBox chkCloakHide 
            Caption         =   "Hide request"
            Height          =   255
            Left            =   240
            TabIndex        =   111
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtCloakCustom 
            Height          =   285
            Left            =   2280
            TabIndex        =   110
            Top             =   1680
            Width           =   2775
         End
         Begin VB.OptionButton optCloak 
            Caption         =   "Send 'unavailable'"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   109
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton optCloak 
            Caption         =   "Ignore request"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   108
            Top             =   1365
            Width           =   1575
         End
         Begin VB.OptionButton optCloak 
            Caption         =   "Custom response:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   107
            Top             =   1665
            Width           =   1935
         End
      End
      Begin VB.ComboBox cmbCloak 
         Height          =   315
         ItemData        =   "frmConnect.frx":0000
         Left            =   1440
         List            =   "frmConnect.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   80
         Width           =   3975
      End
      Begin VB.Label lblDummy 
         BackStyle       =   0  'Transparent
         Caption         =   "Select CTCP:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   105
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox Frames 
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   0
      Left            =   2400
      ScaleHeight     =   4335
      ScaleWidth      =   5535
      TabIndex        =   45
      Tag             =   "Connect"
      Top             =   360
      Width           =   5535
      Begin VB.CheckBox chkModeW 
         Caption         =   "Set mode +w (wallops)"
         Height          =   255
         Left            =   2880
         TabIndex        =   132
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CheckBox chkModeI 
         Caption         =   "Set mode +i (invisible)"
         Height          =   255
         Left            =   360
         TabIndex        =   131
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CommandButton cmdConnectNew 
         Caption         =   "Connect new"
         Height          =   375
         Left            =   1920
         TabIndex        =   116
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CheckBox chkAutoConnect 
         Caption         =   "Auto connect to last used server on startup"
         Height          =   255
         Left            =   120
         TabIndex        =   101
         ToolTipText     =   "Automatically connects to the server specified on startup"
         Top             =   3800
         Width           =   4095
      End
      Begin VB.CheckBox chkShowStartup 
         Caption         =   "Show this window on startup"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         ToolTipText     =   "Shows this window on startup"
         Top             =   3480
         Width           =   2775
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete server"
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   4680
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
      Begin VB.ComboBox comboServer 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmConnect.frx":002E
         Left            =   1560
         List            =   "frmConnect.frx":0030
         TabIndex        =   1
         Top             =   120
         Width           =   2535
      End
      Begin VB.TextBox txtNick 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txtAlternative 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox txtIdent 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox txtRealname 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   1560
         Width           =   3855
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CheckBox chkIdent 
         Caption         =   "Enable ident server"
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         ToolTipText     =   "The ident server should always be enabled."
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   195
         Index           =   6
         Left            =   4200
         TabIndex        =   133
         Top             =   150
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   150
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nick:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Top             =   510
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ident:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   47
         Top             =   1230
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   46
         Top             =   1590
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ident server:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   51
         Top             =   1920
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Altnick:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   870
         Width           =   645
      End
   End
   Begin VB.PictureBox Frames 
      BorderStyle     =   0  'None
      Height          =   4095
      Index           =   8
      Left            =   2520
      ScaleHeight     =   4095
      ScaleWidth      =   5295
      TabIndex        =   117
      Tag             =   "Highlighting"
      Top             =   480
      Width           =   5295
      Begin VB.Frame Frame3 
         Height          =   3495
         Left            =   0
         TabIndex        =   118
         Top             =   0
         Width           =   5295
         Begin VB.PictureBox picChooseHighlight 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3240
            ScaleHeight     =   225
            ScaleWidth      =   465
            TabIndex        =   130
            Top             =   2400
            Width           =   495
         End
         Begin VB.CheckBox chkHighlightUnderline 
            Caption         =   "Highlighted words uses underline"
            Height          =   255
            Left            =   240
            TabIndex        =   129
            Top             =   3030
            Width           =   3135
         End
         Begin VB.CheckBox chkHighlightBold 
            Caption         =   "Highlighted words uses bold"
            Height          =   255
            Left            =   240
            TabIndex        =   128
            Top             =   2715
            Width           =   2775
         End
         Begin VB.CheckBox chkHighlightColor 
            Caption         =   "Highlighted words uses color:"
            Height          =   255
            Left            =   240
            TabIndex        =   127
            Top             =   2400
            Width           =   2895
         End
         Begin VB.CommandButton cmdHighlightDelete 
            Caption         =   "Delete"
            Height          =   345
            Left            =   4320
            TabIndex        =   126
            Top             =   1400
            Width           =   735
         End
         Begin VB.CommandButton cmdHighlightAdd 
            Caption         =   "Add"
            Height          =   345
            Left            =   4320
            TabIndex        =   125
            Top             =   980
            Width           =   735
         End
         Begin VB.ListBox lstHighlight 
            Height          =   840
            ItemData        =   "frmConnect.frx":0032
            Left            =   2160
            List            =   "frmConnect.frx":0034
            TabIndex        =   124
            Top             =   1340
            Width           =   2055
         End
         Begin VB.TextBox txtHighlight 
            Height          =   285
            Left            =   2160
            TabIndex        =   123
            Top             =   980
            Width           =   2055
         End
         Begin VB.CheckBox chkHighlightWords 
            Caption         =   "Highlight words:"
            Height          =   255
            Left            =   240
            TabIndex        =   122
            Top             =   980
            Width           =   1695
         End
         Begin VB.CheckBox chkHighlightActive 
            Caption         =   "Show in active window"
            Height          =   255
            Left            =   240
            TabIndex        =   121
            Top             =   670
            Width           =   2295
         End
         Begin VB.CheckBox chkHighlightNick 
            Caption         =   "Highlight own nickname"
            Height          =   255
            Left            =   240
            TabIndex        =   120
            Top             =   360
            Width           =   2415
         End
         Begin VB.CheckBox chkHighlight 
            Caption         =   "Enable highlighting"
            Height          =   255
            Left            =   120
            TabIndex        =   119
            Top             =   0
            Width           =   1935
         End
      End
   End
   Begin VB.PictureBox Frames 
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   7
      Left            =   2400
      ScaleHeight     =   4335
      ScaleWidth      =   5535
      TabIndex        =   87
      Tag             =   "Away"
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CheckBox chkCancelAway 
         Caption         =   "Cancel normal away on keypress"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         ToolTipText     =   "Removes away when you press return"
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Frame Frame1 
         Caption         =   "Autoway"
         Height          =   1575
         Left            =   120
         TabIndex        =   88
         Top             =   120
         Width           =   5295
         Begin VB.TextBox txtAAMsg 
            Height          =   285
            Left            =   1440
            TabIndex        =   93
            Text            =   "Advanced IRC: 10mins auto-away"
            ToolTipText     =   "Away message to set when going auto-away"
            Top             =   1080
            Width           =   3615
         End
         Begin VB.CheckBox chkCancelAAway 
            Caption         =   "Cancel auto-away on keypress"
            Height          =   255
            Left            =   240
            TabIndex        =   92
            ToolTipText     =   "Removes auto-away when you press return"
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtAAMins 
            Height          =   285
            Left            =   2160
            TabIndex        =   90
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox chkAutoAway 
            Caption         =   "Enable autoaway:"
            Height          =   255
            Left            =   240
            TabIndex        =   89
            ToolTipText     =   "Automatically set away"
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label10 
            Caption         =   "Away-msg:"
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "minute(s)"
            Height          =   195
            Left            =   2880
            TabIndex        =   91
            Top             =   390
            Width           =   825
         End
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   6720
      TabIndex        =   103
      Top             =   4920
      Width           =   1215
   End
   Begin VB.PictureBox Frames 
      BorderStyle     =   0  'None
      Height          =   4335
      Index           =   4
      Left            =   2400
      ScaleHeight     =   4335
      ScaleWidth      =   5535
      TabIndex        =   54
      Tag             =   "Logging"
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CheckBox chkLoggDCC 
         Caption         =   "Log DCC chats"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Log the text in the DCC chat windows"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CheckBox chkLoggPrivat 
         Caption         =   "Log queries"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "Log the text in the query windows"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox chkLoggKanaler 
         Caption         =   "Log channels"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Log the text in the channel windows"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CheckBox chkLoggStatus 
         Caption         =   "Log status"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Log the text in the status windows"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkBrukLogg 
         Caption         =   "Enable logging"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtLogDir 
         Height          =   285
         Left            =   1560
         TabIndex        =   24
         ToolTipText     =   "Which folder all the log files are placed"
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "Log to folder:"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   5040
      TabIndex        =   43
      Top             =   4920
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgPicker 
      Left            =   1680
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   42
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   41
      Top             =   4920
      Width           =   1215
   End
   Begin MSComctlLib.TreeView treeOption 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   9551
      _Version        =   393217
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   4
      Appearance      =   0
   End
   Begin VB.Line Line2 
      X1              =   2400
      X2              =   7920
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   7920
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label lblWhat 
      BackStyle       =   0  'Transparent
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   44
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ContextNum As Long

Private Sub chkAutoAway_Click()
    If chkAutoAway Then
        txtAAMins.Enabled = True
        txtAAMins.BackColor = vbWindowBackground
        chkCancelAAway.Enabled = True
        txtAAMsg.Enabled = True
        txtAAMsg.BackColor = vbWindowBackground
    Else
        txtAAMins.Enabled = False
        txtAAMins.BackColor = vbButtonFace
        chkCancelAAway.Enabled = False
        txtAAMsg.Enabled = False
        txtAAMsg.BackColor = vbButtonFace
    End If
End Sub

Private Sub chkAutoConnect_Click()
    If -chkAutoConnect Then chkShowStartup.Value = 0
End Sub

Private Sub chkBrukLogg_Click()
    If chkBrukLogg Then
        txtLogDir.Enabled = True
        txtLogDir.BackColor = vbWindowBackground
        chkLoggStatus.Enabled = True
        chkLoggKanaler.Enabled = True
        chkLoggPrivat.Enabled = True
        chkLoggDCC.Enabled = True
    Else
        txtLogDir.Enabled = False
        txtLogDir.BackColor = vbButtonFace
        chkLoggStatus.Enabled = False
        chkLoggKanaler.Enabled = False
        chkLoggPrivat.Enabled = False
        chkLoggDCC.Enabled = False
    End If
End Sub

Private Sub chkCloakHide_Click()
    CloakApply 'Save current
End Sub

Private Sub chkDCCPorts_Click()
    TestValue -chkDCCPorts.Value, txtDCCPorts
End Sub

Private Sub chkHentIP_Click()
    If chkHentIP Then
        optLookupType.Enabled = True
        Option1.Enabled = True
        txtIP.Enabled = False
        txtIP.BackColor = vbButtonFace
    Else
        optLookupType.Enabled = False
        Option1.Enabled = False
        txtIP.Enabled = True
        txtIP.BackColor = vbWindowBackground
    End If
End Sub

Private Sub chkHighlight_Click()
    TestValue -chkHighlight, chkHighlightNick, chkHighlightActive, chkHighlightWords, txtHighlight, lstHighlight, cmdHighlightAdd, cmdHighlightDelete, chkHighlightColor, chkHighlightBold, chkHighlightUnderline
    If chkHighlightWords.Enabled Then TestValue -chkHighlightWords, txtHighlight, lstHighlight, cmdHighlightAdd, cmdHighlightDelete
End Sub

Private Sub chkHighlightWords_Click()
    TestValue -chkHighlightWords, txtHighlight, lstHighlight, cmdHighlightAdd, cmdHighlightDelete
End Sub

Private Sub chkIgnoreFiltyper_Click()
    If chkIgnoreFiltyper = 0 Then
        txtIgnoreFiltyper.Enabled = False
        txtIgnoreFiltyper.BackColor = vbButtonFace
    Else
        txtIgnoreFiltyper.Enabled = True
        txtIgnoreFiltyper.BackColor = vbWindowBackground
    End If
End Sub

Private Sub chkShowStartup_Click()
    If -chkShowStartup Then chkAutoConnect.Value = 0
End Sub

Private Sub chkStrip_Click()
    If chkStrip = 0 Then
        chkStripC.Enabled = False
        chkStripB.Enabled = False
        chkStripU.Enabled = False
        chkStripA.Enabled = False
    Else
        chkStripC.Enabled = True
        chkStripB.Enabled = True
        chkStripU.Enabled = True
        chkStripA.Enabled = True
        chkStripA_Click
    End If
End Sub

Private Sub chkStripA_Click()
    If chkStripA = 1 Then
        chkStripC.Enabled = False
        chkStripB.Enabled = False
        chkStripU.Enabled = False
    Else
        chkStripC.Enabled = True
        chkStripB.Enabled = True
        chkStripU.Enabled = True
    End If
End Sub

Private Sub cmdAvbryt_Click()
    Unload Me
End Sub

Private Sub cmbCloak_Click()
    CloakLoad 'Load new
    LastCloakType = cmbCloak.ListIndex
End Sub

Private Sub cmdApply_Click()
    ClickOK True
End Sub

Private Sub cmdChooseFont_Click()
    Dim pFont As StdFont
    Dim C As Long
    On Error Resume Next
    With dlgPicker
        .flags = cdlCFBoth
        .FontName = lblFontPreview.FontName
        .FontSize = lblFontPreview.FontSize
        .FontUnderline = lblFontPreview.FontUnderline
        .FontBold = lblFontPreview.FontBold
        .FontItalic = lblFontPreview.FontItalic
        .ShowFont
        If Err <> 0 Then Exit Sub
        On Error GoTo 0
        Set pFont = New StdFont
        pFont.Bold = .FontBold
        pFont.Italic = .FontItalic
        pFont.name = .FontName
        pFont.Size = .FontSize
        pFont.Strikethrough = .FontStrikethru
        pFont.Underline = .FontUnderline
        For C = 0 To 11
            Set lblColor(C).Font = pFont
        Next
        Set lblFontPreview.Font = pFont
        lblFontPreview.Caption = pFont.name
    End With
End Sub

Sub ColorLoad()
    If ApplyColorPath = "" Then Exit Sub
    cmdColorLoad_Click
    ApplyColorPath = ""
End Sub

Private Sub CloakApply()
    Dim T As TypeCloak
    Dim C As Long
    T = TCloak(cmbCloak.Text)
    With T
        .CloakType = 0
        With optCloak
        For C = 0 To 3
            T.CloakType = T.CloakType + (Abs(.Item(C).Value) * C)
        Next
        End With
        .CustomReply = txtCloakCustom
        .HideRequest = -chkCloakHide
    End With
    STCloak cmbCloak.Text, T
End Sub

Private Sub CloakLoad()
    Dim T As TypeCloak
    T = TCloak(cmbCloak)
    With T
        txtCloakCustom = .CustomReply
        chkCloakHide.Value = -.HideRequest
        optCloak(.CloakType).Value = True
    End With
End Sub

Private Sub cmdCloakReset_Click()
    optCloak(0).Value = True
    chkCloakHide.Value = 0
End Sub

Private Sub cmdCloakUndo_Click()
    CloakLoad
End Sub

Private Sub cmdColorLoad_Click()
    Dim s As String
    Dim T As String
    Dim V As Variant
    Dim C As Long
    If ApplyColorPath = "" Then
        s = FileLoad("Advanced IRC exported colors|*.aic|Configuration settings|*.ini")
    Else
        s = ApplyColorPath
    End If
    If s = "" Then Exit Sub
    With frmMain.INIAccess
        On Error Resume Next
        T = .INIFile
        .INIFile = s
        V = SysGetColorInfo
        .INIFile = T
        If ((Err <> 0) Or (V(0) = "")) Then
            On Error GoTo 0
            MsgBox "This file does not contain Advanced IRC color information!", vbExclamation, "Error"
            Exit Sub
        End If
        On Error GoTo 0
        'Fil OK, fortsett
    End With
    For C = 0 To 11
        lblColor(C).ForeColor = V(C)
    Next
    picColBG.BackColor = V(12)
    picChooseURL.BackColor = V(13)
    picChooseBrand.BackColor = V(14)
    picChooseStd.BackColor = V(15)
    picChooseSec.BackColor = V(16)
    chkmIRCColors = -V(17)
    With lblFontPreview
        With .Font
            On Error Resume Next
            .name = V(18)
            .Size = ToVal(V(19))
            .Bold = -ToVal(V(20))
            .Underline = -ToVal(V(21))
            .Italic = -ToVal(V(22))
            On Error GoTo 0
        End With
        .Caption = V(18)
    End With
End Sub

Private Sub cmdColormIRC_Click()
    Dim s As String
    Dim T As String
    Dim CSet As String
    Dim V As Variant 'Orginal
    Dim V2() As Variant '(Falsk)
    Dim C As Long
    s = FileLoad("mIRC configuration file|mirc.ini")
    If s = "" Then Exit Sub
    With frmMain.INIAccess
        T = .INIFile
        .INIFile = s
        .INIEntry = "colours"
        CSet = .INIGetSetting("n0")
        .INIFile = T
    End With
    If CSet = "" Then
        MsgBox "Color information is missing!", vbExclamation, "Error"
        Exit Sub
        Exit Sub
    End If
    V = Split(CSet, ",")
    If UBound(V) <> 25 Then  'Feil format
        If MsgBox("This is not a mIRC v5.91 settings file, or corrupt file!" & vbCrLf & vbCrLf & _
        "Importing these color settings may cause corrupt colors or make Advanced IRC unstable." & vbCrLf & _
        "Proceed with color importing?", vbExclamation + vbYesNo, "Error") = vbNo Then Exit Sub
    End If
    ReDim Preserve V2(0 To 12)
    V2(0) = V(7)    'Join text
    V2(1) = V(16)   'Part text
    V2(2) = V(17)   'Quit text
    V2(3) = V(10)   'Nick text
    V2(4) = V(8)    'Kick text
    V2(5) = V(9)    'Mode text
    V2(6) = V(1)    'Action text
    V2(7) = V(4)    'Error text
    V2(8) = V(18)   'Topic text
    V2(9) = V(11)   'Normal text
    V2(10) = V(15)  'Own text
    V2(11) = V(12)  'Notice text
    V2(12) = V(0)   'Background color
    'Og en ting til - forbannet være mIRC
    'Hvis du er mIRC-sympatør kan du se å dra deg lukt til helvete
    'If you like mIRC, please do not read this source code
    For C = 0 To 11
        lblColor(C).ForeColor = mIRCColors(V2(C))
    Next
    picColBG.BackColor = mIRCColors(V2(C))
End Sub

Private Sub cmdColorSave_Click()
    Dim s As String
    Dim T As String
    Dim V As Variant
    Dim C As Long
    s = FileSave("Advanced IRC exported colors|*.aic|Configuration settings|*.ini")
    If s = "" Then Exit Sub
    With frmMain.INIAccess
        T = .INIFile
        ReDim V(0 To 22)
        For C = 0 To 11
            V(C) = lblColor(C).ForeColor
        Next
        V(12) = picColBG.BackColor
        V(13) = picChooseURL.BackColor
        V(14) = picChooseBrand.BackColor
        V(15) = picChooseStd.BackColor
        V(16) = picChooseSec.BackColor
        V(17) = -chkmIRCColors
        With lblFontPreview
            V(18) = .FontName
            V(19) = .FontSize
            V(20) = -.FontBold
            V(21) = -.FontUnderline
            V(22) = -.FontItalic
        End With
        .INIFile = s
        SysSaveColorInfo V
        .INIFile = T
    End With
End Sub

Private Sub cmdConnectNew_Click()
    If Not ClickOK Then Exit Sub
    NewStatusWnd IRCInfo.Server, IRCInfo.Port
    'InitConnect
    fActive.LogBox.HardRefresh
End Sub

Private Sub cmdDelete_Click()
    Dim D As Integer
    D = comboServer.ListIndex
    If D = -1 Then Exit Sub
    comboServer.RemoveItem D
    If D > comboServer.ListCount Then D = comboServer.ListCount
    If comboServer.ListCount = -1 Then Exit Sub
    If D > comboServer.ListCount - 1 Then
        comboServer.ListIndex = D - 1
        If comboServer.ListIndex = -1 Then
            txtPort = ""
        Else
            txtPort = IRCInfo.PortLst(D - 1)
        End If
    Else
        comboServer.ListIndex = D
        txtPort = IRCInfo.PortLst(D + 1)
    End If
    SII True
End Sub

Private Sub cmdHelp_Click()
    ShowHelp hwnd, ContextNum
End Sub

Private Sub cmdHighlightAdd_Click()
    If Not txtHighlight.Text = "" Then lstHighlight.AddItem txtHighlight.Text: txtHighlight.Text = ""
End Sub

Private Sub cmdHighlightDelete_Click()
    If Not lstHighlight.ListIndex = -1 Then lstHighlight.RemoveItem lstHighlight.ListIndex
End Sub

Private Sub cmdOK_Click()
    ClickOK
End Sub

Function ClickOK(Optional ByVal ApplyOnly As Boolean = False) As Boolean
    Dim C As Long
    
    If chkDCCPorts.Value = 1 Then
        If Not DCC_Range_Check(txtDCCPorts.Text) Then MsgBox "Please specify a port range in the format" & _
        vbCrLf & vbCrLf & "STARTPORT[-ENDPORT]" & vbCrLf & vbCrLf & "separated by a whitespace for each entry.", _
        vbExclamation, "Error": Exit Function
    End If
    If chkHentIP.Value = 0 Then
        If Not IsValidIP(txtIP) Then
            MsgBox "You'll have to type a valid IP!", vbOKOnly, "Error"
            Exit Function
        End If
    End If
    If chkAutoAway.Value = 1 Then
        If txtAAMins = "" Then txtAAMins = "0"
        If Not IsNumeric(txtAAMins) Then
            MsgBox "You'll have to type a number in the 'Away minutes' box!", vbExclamation, "Error"
            Exit Function
        Else
            txtAAMins = CInt(txtAAMins)
        End If
        If CInt(txtAAMins) <= 0 Then
            MsgBox "You'll have to type a number greater than 0 in the 'Away minutes' box!", vbExclamation, "Error"
            Exit Function
        End If
        If txtAAMsg = "" Then
            MsgBox "You'll have to type an away message in the 'Away message' box!", vbExclamation, "Error"
            Exit Function
        End If
    End If
    
    'No need for more
    SaveCloakInfo
    
    With DCCInfo
        .DownloadDir = txtDownDir
        .ProtectVirus = -chkBeskyttVirus
        .JoinIgnore = -chkJoinIgnore
        .DoIgnoreFiltyper = -chkIgnoreFiltyper
        .IgnoreFiltyper = txtIgnoreFiltyper
        .AutoAccept = -chkDCCAutoGet
        If Not IsNumeric(txtSendeBuffer) Then txtSendeBuffer = "0"
        .SendeBuffer = txtSendeBuffer
        .PumpDCC = -chkPumpDCC
        .PassiveDCC = -chkPassiveDCC
        .UDCCPorts = -chkDCCPorts
        .DCCPortRange = txtDCCPorts
        .SafeMode = -chkDCCSafe
        
        If Not Right(.DownloadDir, 1) = "\" Then .DownloadDir = .DownloadDir & "\"
        On Error Resume Next
        MkDir .DownloadDir
        On Error GoTo 0
    End With
    SaveDCCInfo
    
    With IPInfo
        .IP = txtIP
        .UseCustomIP = -chkHentIP
        .LookupType = -optLookupType.Value
    End With
    If Not -chkHentIP Then DCCIP = txtIP
    SaveIPInfo
    
    With LogInfo
        .LoggDir = txtLogDir
        .LoggStatus = -chkLoggStatus
        .LoggKanaler = -chkLoggKanaler
        .LoggPrivat = -chkLoggPrivat
        .LoggDCC = -chkLoggDCC
        If Not Right(.LoggDir, 1) = "\" Then .LoggDir = .LoggDir & "\"
        On Error Resume Next
        MkDir .LoggDir
        On Error GoTo 0
        If ((-chkBrukLogg) And Not .BrukLogg) Then
            If .LoggStatus Then OpenAllLogs logStatus Else CloseAllLogs logStatus
            If .LoggKanaler Then OpenAllLogs logChannel Else CloseAllLogs logChannel
            If .LoggPrivat Then OpenAllLogs logPrivate Else CloseAllLogs logPrivate
            If .LoggDCC Then OpenAllLogs logDCC Else CloseAllLogs logDCC
        ElseIf ((Not -chkBrukLogg) And .BrukLogg) Then
            CloseAllLogs logAll
        End If
        .BrukLogg = -chkBrukLogg
    End With
    SaveLogInfo
    
    With DisplayInfo
        .Timestamp = txtTimestamp
        .StripCodes = -chkStrip
        .StripC = -chkStripC
        .StripB = -chkStripB
        .StripU = -chkStripU
        .StripA = -chkStripA
        .FlashNew = -chkFlashNew
        .FlashAny = -chkFlashAny
        .ColorActivity = -chkColorActivity
        .ShowNicklist = -chkShowNicklist
    End With
    SaveDisplayInfo
    
    With ColorInfo
        .cJoin = lblColor(0).ForeColor
        .cPart = lblColor(1).ForeColor
        .cQuit = lblColor(2).ForeColor
        .cNick = lblColor(3).ForeColor
        .cKick = lblColor(4).ForeColor
        .cMode = lblColor(5).ForeColor
        .cAction = lblColor(6).ForeColor
        .cStatus = lblColor(7).ForeColor
        .cTopic = lblColor(8).ForeColor
        .cNormal = lblColor(9).ForeColor
        .cOwn = lblColor(10).ForeColor
        .cNotice = lblColor(11).ForeColor
        .cBackColor = picColBG.BackColor
        .cURLColor = picChooseURL.BackColor
        .cBrandColor = picChooseBrand.BackColor
        .cStdColor = picChooseStd.BackColor
        .cSecColor = picChooseSec.BackColor
        If Not .UsemIRCColors = -chkmIRCColors Then
            .UsemIRCColors = -chkmIRCColors
            InitColors True
        End If
        If Not .Font Is lblFontPreview.Font Then
            Set .Font = lblFontPreview.Font
        End If
    End With
    SaveColorInfo
    AC_Code = ColorCode & StdColNum
    
    With AwayInfo
        .AAUse = -chkAutoAway
        .AAMinutes = txtAAMins
        .AACancelAway = -chkCancelAAway
        .AAMsg = txtAAMsg
        .CancelAway = -chkCancelAway
    End With
    SaveAwayInfo
    
    With HighlightInfo
        .UseHighlight = -chkHighlight
        .HiNick = -chkHighlightNick
        .HiActive = -chkHighlightActive
        .HiWords = -chkHighlightWords
        If lstHighlight.ListCount > 0 Then
            ReDim .HiWordList(0 To lstHighlight.ListCount - 1)
            For C = 0 To lstHighlight.ListCount - 1
                .HiWordList(C) = lstHighlight.list(C)
            Next
        End If
        .UseColor = -chkHighlightColor
        .HiColor = picChooseHighlight.BackColor
        .UseBold = -chkHighlightBold
        .UseUnderline = -chkHighlightUnderline
    End With
    SaveHighlightInfo
    
    ColorWindows
    UpdateChannelWindows 'Show/Hide nicklisten
    
    SII '(S)ave (I)RC (I)nformation
    ClickOK = True
    If Not ApplyOnly Then Unload Me
End Function

Private Sub cmdResetColors_Click()
    Dim V As Variant
    Dim C As Long
    
    With ColorInfo
        .cJoin = RGB(0, 127, 127)
        .cPart = RGB(0, 127, 127)
        .cQuit = RGB(255, 0, 0)
        .cNick = RGB(0, 127, 0)
        .cKick = RGB(255, 0, 0)
        .cMode = RGB(0, 0, 255)
        .cAction = RGB(127, 0, 127)
        .cStatus = RGB(0, 0, 127)
        .cTopic = RGB(0, 127, 0)
        .cNormal = RGB(0, 0, 0)
        .cOwn = RGB(0, 0, 0)
        .cNotice = RGB(127, 0, 0)
        .cBackColor = RGB(207, 207, 207)
        .cURLColor = RGB(0, 0, 255)
        .cBrandColor = RGB(0, 0, 255)
        .cStdColor = RGB(0, 0, 255)
        .cSecColor = RGB(127, 127, 127)
        Set .Font = New StdFont
        .Font.name = "Courier New"
        .Font.Size = 9
        Set lblFontPreview.Font = .Font
        lblFontPreview.Caption = .Font.name
    End With
    SaveColorInfo
    
    V = SysGetColorInfo
    For C = 0 To 11
        lblColor(C).ForeColor = V(C)
        Set lblColor(C).Font = ColorInfo.Font
    Next
    picColBG.BackColor = V(12)
    picChooseURL.BackColor = V(13)
    picChooseBrand.BackColor = V(14)
    picChooseStd.BackColor = V(15)
    picChooseSec.BackColor = V(16)

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdConnect_Click()
    If Not ClickOK Then Exit Sub
    InitConnect
    fActive.LogBox.HardRefresh
End Sub

Private Sub comboServer_Change()
    ChkConVal
End Sub

Private Sub comboServer_Click()
    If comboServer.ListIndex = -1 Then Exit Sub
    txtPort = IRCInfo.PortLst(comboServer.ListIndex)
End Sub

Private Sub comboServer_Validate(Cancel As Boolean)
    If comboServer.ListIndex = -1 Then Exit Sub
    txtPort = IRCInfo.PortLst(comboServer.ListIndex)
End Sub

Sub ChkConVal()
    If ((comboServer.Text = "") Or (txtPort = "") Or (txtNick = "") Or (txtAlternative = "") Or (txtIdent = "") Or (txtRealname = "")) Then
        cmdConnect.Enabled = False
    ElseIf Not StatusWnd(ActiveServer).IsConnected Then
        cmdConnect.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim TmpVar As Variant
    Dim C As Long
    'GetIRCInfo
    With IRCInfo
        comboServer.Clear
        For C = 0 To UBound(.SrvLst)
            comboServer.AddItem .SrvLst(C)
        Next
        comboServer = .Server
        txtPort = .Port
        txtNick = .Nick
        txtAlternative = .Alternative
        txtIdent = .Ident
        txtRealname = .Realname
        chkIdent = -.UseIdent
        chkModeI = -.ModeInvisible
        chkModeW = -.ModeWallops
        If .AutoMode = 1 Then
            chkShowStartup.Value = 1
        ElseIf .AutoMode = 2 Then
            chkAutoConnect.Value = 1
        End If
    End With
    
    'Trenger ikke mer
    'GetCloakInfo
    
    'GetDCCInfo
    With DCCInfo
        txtDownDir = .DownloadDir
        chkBeskyttVirus = -.ProtectVirus
        chkJoinIgnore = -.JoinIgnore
        chkIgnoreFiltyper = -.DoIgnoreFiltyper
        txtIgnoreFiltyper = .IgnoreFiltyper
        chkDCCAutoGet = -.AutoAccept
        txtSendeBuffer = .SendeBuffer
        chkPumpDCC = -.PumpDCC
        chkPassiveDCC = -.PassiveDCC
        chkDCCPorts = -.UDCCPorts
        txtDCCPorts = .DCCPortRange
        chkDCCSafe = -.SafeMode
    End With
    chkIgnoreFiltyper_Click
    chkDCCPorts_Click
    
    'GetIPInfo
    With IPInfo
        txtIP = .IP
        chkHentIP = -.UseCustomIP
        If .LookupType Then optLookupType.Value = True Else Option1.Value = True
    End With
    chkHentIP_Click
    
    'GetLogInfo
    With LogInfo
        chkBrukLogg = -.BrukLogg
        txtLogDir = .LoggDir
        chkLoggStatus = -.LoggStatus
        chkLoggKanaler = -.LoggKanaler
        chkLoggPrivat = -.LoggPrivat
        chkLoggDCC = -.LoggDCC
    End With
    chkBrukLogg_Click
    
    'GetDisplayInfo
    With DisplayInfo
        txtTimestamp = .Timestamp
        chkStrip = -.StripCodes
        chkStripC = -.StripC
        chkStripB = -.StripB
        chkStripU = -.StripU
        chkStripA = -.StripA
        chkFlashNew = -.FlashNew
        chkFlashAny = -.FlashAny
        chkColorActivity = -.ColorActivity
        chkShowNicklist = -.ShowNicklist
    End With
    chkStrip_Click
    
    'GetColorInfo
    TmpVar = SysGetColorInfo
    With ColorInfo
        For C = 0 To 11
            lblColor(C).ForeColor = TmpVar(C)
            Set lblColor(C).Font = .Font
        Next
        Set lblFontPreview.Font = .Font
        lblFontPreview.Caption = .Font.name
    End With
    picColBG.BackColor = TmpVar(12)
    picChooseURL.BackColor = TmpVar(13)
    picChooseBrand.BackColor = TmpVar(14)
    picChooseStd.BackColor = TmpVar(15)
    picChooseSec.BackColor = TmpVar(16)
    chkmIRCColors = -ColorInfo.UsemIRCColors
    'Enkleste måten
    
    'GetAwayInfo
    With AwayInfo
        chkAutoAway = -.AAUse
        txtAAMins = .AAMinutes
        chkCancelAAway = -.AACancelAway
        txtAAMsg = .AAMsg
        chkCancelAway = -.CancelAway
    End With
    chkAutoAway_Click
    
    With HighlightInfo
        chkHighlight = -.UseHighlight
        chkHighlightNick = -.HiNick
        chkHighlightActive = -.HiActive
        chkHighlightWords = -.HiWords
        For C = LBound(.HiWordList) To UBound(.HiWordList)
            lstHighlight.AddItem .HiWordList(C)
        Next
        chkHighlightColor = -.UseColor
        picChooseHighlight.BackColor = .HiColor
        chkHighlightBold = -.UseBold
        chkHighlightUnderline = -.UseUnderline
    End With
    chkHighlight_Click
    
    
    If Not StatusWnd(ActiveServer).IsOpen Then cmdConnect.Enabled = False Else cmdConnect.Enabled = True
    ChkConVal
    
    Dim NodX As Node
    On Error GoTo 0
    With treeOption
        .Nodes.Add , , "Root1", "General"
        .Nodes.Add "Root1", tvwChild, "1Connect", "Connect"
        .Nodes.Add "Root1", tvwChild, "1Cloaking", "Cloaking"
        .Nodes.Add "Root1", tvwChild, "1DCC", "DCC"
        .Nodes.Add "Root1", tvwChild, "1IP", "IP"
        .Nodes.Add "Root1", tvwChild, "1Logging", "Logging"
        .Nodes.Add "Root1", tvwChild, "1Away", "Away"
        .Nodes.Add , , "Root2", "Display"
        .Nodes.Add "Root2", tvwChild, "2General", "General"
        .Nodes.Add "Root2", tvwChild, "2Highlighting", "Highlighting"
        .Nodes.Add "Root2", tvwChild, "2Colors", "Colors"
        .Nodes(1).Expanded = True
        .Nodes(8).Expanded = True
    End With
    
    SwitchFrames 0
    cmbCloak.ListIndex = LastCloakType
End Sub

Sub SII(Optional ByVal JustServerList As Boolean = False)
    Dim C As Long
    Dim D As Boolean
    For C = 0 To comboServer.ListCount - 1
        If comboServer.list(C) = comboServer.Text Then D = True
    Next
    If Not D Then comboServer.AddItem comboServer.Text
    With IRCInfo
        If Not IsNumeric(txtPort) Then txtPort = "6667"
        .Server = comboServer
        ReDim Preserve .SrvLst(0 To comboServer.ListCount - 1)
        For C = 0 To comboServer.ListCount - 1
            .SrvLst(C) = comboServer.list(C)
        Next
        .Port = txtPort
        ReDim Preserve .PortLst(0 To comboServer.ListCount - 1)
        For C = 0 To comboServer.ListCount - 1
            If comboServer.list(C) = .Server Then .PortLst(C) = txtPort
        Next
        If Not JustServerList Then
            .Nick = txtNick
            .Alternative = txtAlternative
            .Ident = txtIdent
            .Realname = txtRealname
            .UseIdent = -chkIdent
            .ModeInvisible = -chkModeI
            .ModeWallops = -chkModeW
            .AutoMode = 0
            If -chkShowStartup Then
                .AutoMode = 1
            ElseIf -chkAutoConnect Then
                .AutoMode = 2
            End If
        End If
    End With
    SaveIRCInfo
End Sub

Public Function DCC_Range_Check(ByVal s As String) As Boolean
    Dim V As Variant, V2 As Variant, C As Long
    Dim D As Long, Cm() As Long, CmCount As Long
    If InStr(1, txtDCCPorts.Text, " ") = 0 Then
        V2 = Split(txtDCCPorts.Text & " 0", " ")
        ReDim Preserve V2(0 To 0)
    Else
        V2 = Split(txtDCCPorts.Text, " ")
    End If
    For C = LBound(V2) To UBound(V2)
        V = Split(V2(C), "-")
        If IsArray(V) Then
            If (LBound(V) = 0) And (UBound(V) = 1) Then
                If (Not IsNumeric(V(0))) Or (Not IsNumeric(V(1))) Then Exit Function
                If V(1) < V(0) Then D = V(0): V(0) = V(1): V(1) = D 'Switch
                
                ReDim Preserve Cm(1 To CmCount + (V(1) - V(0)) + 1)
                
                For D = V(0) To V(1)
                    Inc CmCount
                    Cm(CmCount) = D
                Next
            ElseIf (LBound(V) = 0) And (UBound(V) = 0) Then
                If Not IsNumeric(V(0)) Then Exit Function
                Inc CmCount
                ReDim Preserve Cm(1 To CmCount)
                Cm(CmCount) = V(0)
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    Next
    DCC_Range_Check = True
    DCCInfo.DCCPortList = Cm
End Function

Private Sub lblColor_Click(Index As Integer)
    lblColor(Index).ForeColor = ChooseColor(lblColor(Index).ForeColor)
End Sub

Private Sub lstHighlight_GotFocus()
    cmdHighlightDelete.Default = True
End Sub

Private Sub lstHighlight_LostFocus()
    cmdOK.Default = True
End Sub

Private Sub optCloak_Click(Index As Integer)
    TestValue optCloak(3).Value, txtCloakCustom
    CloakApply 'Save current
End Sub

Private Sub picChooseBrand_Click()
    picChooseBrand.BackColor = ChooseColor(picChooseBrand.BackColor)
End Sub

Private Sub picChooseHighlight_Click()
    picChooseHighlight.BackColor = ChooseColor(picChooseHighlight.BackColor, True)
End Sub

Private Sub picChooseSec_Click()
    picChooseSec.BackColor = ChooseColor(picChooseSec.BackColor, True)
End Sub

Private Sub picChooseStd_Click()
    picChooseStd.BackColor = ChooseColor(picChooseStd.BackColor, True)
End Sub

Private Sub picChooseURL_Click()
    picChooseURL.BackColor = ChooseColor(picChooseURL.BackColor, True)
End Sub

Private Sub picColBG_Click()
    picColBG.BackColor = ChooseColor(picColBG.BackColor)
End Sub

Private Sub treeOption_NodeClick(ByVal Node As MSComctlLib.Node)
    Select Case Node.Key
        Case "Root1", "1Connect"
            SwitchFrames 0
        Case "1Cloaking"
            SwitchFrames 1
        Case "1DCC"
            SwitchFrames 2
        Case "1IP"
            SwitchFrames 3
        Case "1Logging"
            SwitchFrames 4
        Case "1Away"
            SwitchFrames 7
        Case "Root2", "2General"
            SwitchFrames 5
        Case "2Highlighting"
            SwitchFrames 8
        Case "2Colors"
            SwitchFrames 6
    End Select
End Sub

Sub SwitchFrames(ByVal FrameNum As Integer)
    Dim C As Long
    If ((FrameNum < Frames.LBound) Or (FrameNum > Frames.UBound)) Then Exit Sub
    If FrameNum = 0 Then
        cmdConnect.Default = True
    Else
        cmdOK.Default = True
    End If
    For C = Frames.LBound To Frames.UBound
        If FrameNum = C Then
            Frames(C).Visible = True
            lblWhat = Frames(C).Tag
        Else
            Frames(C).Visible = False
        End If
    Next
    ContextNum = FrameNum + 2001
End Sub

Private Sub txtAlternative_Change()
    ChkConVal
End Sub

Private Sub txtCloakCustom_Change()
    CloakApply 'Save current
End Sub

Private Sub txtHighlight_GotFocus()
    cmdHighlightAdd.Default = True
End Sub

Private Sub txtHighlight_LostFocus()
    cmdOK.Default = True
End Sub

Private Sub txtIdent_Change()
    ChkConVal
End Sub

Private Sub txtNick_Change()
    ChkConVal
End Sub

Private Sub txtPort_Change()
    ChkConVal
End Sub

Private Sub txtRealname_Change()
    ChkConVal
End Sub

Private Sub VScroll1_Change()
    If txtSendeBuffer = "" Then txtSendeBuffer = 0
    If VScroll1.Value = 0 Then
        txtSendeBuffer = CLng(txtSendeBuffer) + 256
    ElseIf VScroll1.Value = 2 Then
        txtSendeBuffer = CLng(txtSendeBuffer) - 256
    End If
    If txtSendeBuffer < 1 Then txtSendeBuffer = 1
    VScroll1.Value = 1
End Sub

Function ChooseColor(ByVal OldColor As Long, Optional ByVal StrictColors As Boolean = False) As OLE_COLOR
    Dim pColor As OLE_COLOR
    ChooseColor = OldColor
    On Error Resume Next
    With dlgPicker
        .flags = 0
        .Color = OldColor
        .ShowColor
        If Err <> 0 Then Exit Function
        On Error GoTo 0
        pColor = .Color
    End With
    If ((pColor = RGB(128, 0, 128)) And (StrictColors)) Then 'm$ bug workaround
        pColor = RGB(127, 0, 127)
    End If
    If ((StrictColors) And (Not TestColor(pColor))) Then 'Color not approved
        MsgBox "This color is not a standard color and can not be used.", vbOKOnly, "Error"
        Exit Function
    End If
    ChooseColor = pColor
End Function

Function FileLoad(ByVal sFilter As String) As String
    On Error Resume Next
    Dim FF As Integer
    With dlgPicker
        .flags = 0
        .Filter = sFilter
        .FilterIndex = 0
        .Filename = ""
        .ShowOpen
        If Err <> 0 Then Exit Function
        On Error GoTo 0
        If .Filename = "" Then Exit Function
        If Not TrimPath(.Filename) = TrimPath(.Filename, True) Then Exit Function
        On Error Resume Next
        FF = FreeFile
        Open .Filename For Random As FF
        If Err <> 0 Then Exit Function
        On Error GoTo 0
        If LOF(FF) = 0 Then
            Close FF
            Kill .Filename
            Exit Function
        End If
        Close FF
        FileLoad = .Filename
    End With
End Function

Function FileSave(ByVal sFilter As String) As String
    On Error Resume Next
    Dim FF As Integer
    With dlgPicker
        .flags = 0
        .Filter = sFilter
        .FilterIndex = 0
        .Filename = ""
        .ShowSave
        If Err <> 0 Then Exit Function
        On Error GoTo 0
        If .Filename = "" Then Exit Function
        If Not TrimPath(.Filename) = TrimPath(.Filename, True) Then Exit Function
        On Error Resume Next
        FF = FreeFile
        Open .Filename For Random As FF
        If Err <> 0 Then Exit Function
        On Error GoTo 0
        Close FF
        Kill .Filename
        FileSave = .Filename
    End With
End Function

