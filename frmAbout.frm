VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Advanced IRC"
   ClientHeight    =   4920
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5205
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      Height          =   3960
      Left            =   120
      Picture         =   "frmAbout.frx":1CCA
      ScaleHeight     =   3900
      ScaleWidth      =   720
      TabIndex        =   1
      Top             =   240
      Width           =   780
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3840
      TabIndex        =   0
      Top             =   4515
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced IRC is freeware and open sourced."
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   3285
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All the guys at #vbnorge @ EFNet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   1320
      MouseIcon       =   "frmAbout.frx":AF4E
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Tag             =   "http://www.web-amp.com/vbnorge"
      Top             =   3360
      Width           =   2475
   End
   Begin VB.Label lblWebsite 
      BackStyle       =   0  'Transparent
      Caption         =   "Who helps me solve trivial and difficult coding problems, and of course for being good friends."
      Height          =   675
      Index           =   3
      Left            =   1320
      TabIndex        =   10
      Top             =   3600
      Width           =   2685
   End
   Begin VB.Label lblWebsite 
      BackStyle       =   0  'Transparent
      Caption         =   "For being a very good friend and helping me with virtually anything I need help with."
      Height          =   675
      Index           =   2
      Left            =   1320
      TabIndex        =   9
      Top             =   2640
      Width           =   2685
   End
   Begin VB.Label lblWebsite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special thanks to:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1050
      TabIndex        =   8
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://home.no.net/slice/airc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   1680
      MouseIcon       =   "frmAbout.frx":B258
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Tag             =   "http://home.no.net/slice/airc"
      Top             =   1680
      Width           =   2085
   End
   Begin VB.Label lblWebsite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WWW:"
      Height          =   195
      Index           =   0
      Left            =   1050
      TabIndex        =   6
      Top             =   1680
      Width           =   510
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Erlend Sommerfelt Ervik"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   1320
      MouseIcon       =   "frmAbout.frx":B562
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Tag             =   "http://home.c2i.net/sommerfelt_ervik/directirc2"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -30
      X2              =   5218
      Y1              =   4365
      Y2              =   4365
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2000-2002 Kim Tore Jensen"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1050
      TabIndex        =   2
      Top             =   1320
      Width           =   2940
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced IRC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   2850
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   -15
      X2              =   5218
      Y1              =   4380
      Y2              =   4380
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   630
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Running version " & VerStr & vbCrLf & "Compiled " & CompileConst
End Sub

Private Sub lblURL_Click(Index As Integer)
    ShellExecute hwnd, vbNullString, lblURL(Index).Tag, vbNull, vbNullString, 0
End Sub
