VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2120D62E-1B94-47CE-956E-F31CED1DA6C4}#19.3#0"; "aircutils.ocx"
Begin VB.Form frmScrEdit 
   Caption         =   "Script viewer"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScrEdit.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBut1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   6360
      ScaleHeight     =   1095
      ScaleWidth      =   1215
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
      Begin VB.CommandButton cmdSaveExit 
         Caption         =   "Save && exit"
         Enabled         =   0   'False
         Height          =   345
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   345
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   345
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList imgScript 
      Left            =   8280
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScrEdit.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScrEdit.frx":0B24
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAddEvent 
      Caption         =   "Add event"
      Enabled         =   0   'False
      Height          =   345
      Left            =   6360
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddAlias 
      Caption         =   "Add alias"
      Enabled         =   0   'False
      Height          =   345
      Left            =   6360
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtScFunc 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   4935
   End
   Begin VB.TextBox txtScName 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   4935
   End
   Begin VB.TextBox txtScript 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   840
      Width           =   6135
   End
   Begin aircutils.ScList scList1 
      Align           =   4  'Align Right
      Height          =   4905
      Left            =   7665
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   8652
   End
   Begin VB.Label lblDummy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Script function:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label lblDummy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Script name:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   900
   End
End
Attribute VB_Name = "frmScrEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurSc As Integer, IsEd As Boolean

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Activate()
frmMain.WSwitch.ActWnd Me
End Sub

Private Sub Form_Load()
frmMain.WSwitch.AddWnd Me, 0, wndOther
frmMain.WSwitch.ActWnd Me
Dim C As Integer
If ScriptArrayU = 0 Then Exit Sub
ReDim ScTemp(1 To ScriptArrayU)
With scList1
For C = 1 To ScriptArrayU
.AddScr ScriptArray(C).Sc_Name, imgScript.ListImages(1).Picture
ScTemp(C).Sc_Name = ScriptArray(C).Sc_Name
ScTemp(C).Sc_Func = ScriptArray(C).Sc_Func
ScTemp(C).Code = ScriptArray(C).Code
Next
End With
scList1.ActScr 1
CurSc = 1
With ScTemp(CurSc)
txtScName.Text = .Sc_Name
txtScFunc.Text = .Sc_Func
IsEd = True
txtScript.Text = .Code
End With
End Sub

Private Sub Form_Resize()
If WindowState = 1 Then Exit Sub
txtScName.Width = Width - 4185
txtScFunc.Width = txtScName.Width
txtScript.Width = Width - 2985
txtScript.Height = Height - 1305
'cmd buttons
With cmdAddAlias
.Left = Width - 2760
cmdAddEvent.Left = .Left
picBut1.Left = .Left
End With
picBut1.Top = Height - 1560
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.WSwitch.RemWnd Me
End Sub

Private Sub scList1_ScChange(ByVal ScNum As Integer)
With ScTemp(CurSc)
.Sc_Name = txtScName
.Sc_Func = txtScFunc
.Code = txtScript
End With
CurSc = ScNum
With ScTemp(ScNum)
txtScName.Text = .Sc_Name
txtScFunc.Text = .Sc_Func
IsEd = True
txtScript.Text = .Code
End With
End Sub

Private Sub txtScript_Change()
'If Not ScTemp(CurSc).Edited Then
'ScTemp(CurSc).Edited = True
'scList1.SetIcon CurSc, imgScript.ListImages(2).Picture
'End If
If IsEd Then IsEd = False: Exit Sub
scList1.ColScr CurSc, vbBlue
End Sub
