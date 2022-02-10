VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmChanSet 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9135
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "frmChanSet"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Avbryt"
      Height          =   375
      Left            =   6240
      TabIndex        =   51
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7680
      TabIndex        =   50
      Top             =   6360
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Modus"
      TabPicture(0)   =   "frmChanSet.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Lister"
      TabPicture(1)   =   "frmChanSet.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   5415
         Left            =   -74760
         TabIndex        =   2
         Top             =   480
         Width           =   8415
      End
      Begin VB.Frame Frame1 
         Height          =   5415
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   8415
         Begin VB.CheckBox chkCW 
            Caption         =   "+W"
            Height          =   255
            Left            =   7560
            TabIndex        =   49
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox chkCZ 
            Caption         =   "+Z"
            Height          =   255
            Left            =   7560
            TabIndex        =   48
            Top             =   1800
            Width           =   615
         End
         Begin VB.CheckBox chkCQ 
            Caption         =   "+Q"
            Height          =   255
            Left            =   5280
            TabIndex        =   47
            Top             =   3240
            Width           =   615
         End
         Begin VB.CheckBox chkCR 
            Caption         =   "+R"
            Height          =   255
            Left            =   5280
            TabIndex        =   46
            Top             =   3600
            Width           =   615
         End
         Begin VB.CheckBox chkCS 
            Caption         =   "+S"
            Height          =   255
            Left            =   5280
            TabIndex        =   45
            Top             =   3960
            Width           =   615
         End
         Begin VB.CheckBox chkCT 
            Caption         =   "+T"
            Height          =   255
            Left            =   5280
            TabIndex        =   44
            Top             =   4320
            Width           =   615
         End
         Begin VB.CheckBox chkCU 
            Caption         =   "+U"
            Height          =   255
            Left            =   5280
            TabIndex        =   43
            Top             =   4680
            Width           =   615
         End
         Begin VB.CheckBox chkCV 
            Caption         =   "+V"
            Height          =   255
            Left            =   5280
            TabIndex        =   42
            Top             =   5040
            Width           =   615
         End
         Begin VB.CheckBox chkCY 
            Caption         =   "+Y"
            Height          =   255
            Left            =   7560
            TabIndex        =   41
            Top             =   1440
            Width           =   615
         End
         Begin VB.CheckBox chkCX 
            Caption         =   "+X"
            Height          =   255
            Left            =   7560
            TabIndex        =   40
            Top             =   1080
            Width           =   615
         End
         Begin VB.CheckBox chkCJ 
            Caption         =   "+J"
            Height          =   255
            Left            =   5280
            TabIndex        =   39
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox chkCM 
            Caption         =   "+M"
            Height          =   255
            Left            =   5280
            TabIndex        =   38
            Top             =   1800
            Width           =   615
         End
         Begin VB.CheckBox chkCN 
            Caption         =   "+N"
            Height          =   255
            Left            =   5280
            TabIndex        =   37
            Top             =   2160
            Width           =   615
         End
         Begin VB.CheckBox chkCO 
            Caption         =   "+O"
            Height          =   255
            Left            =   5280
            TabIndex        =   36
            Top             =   2520
            Width           =   615
         End
         Begin VB.CheckBox chkCP 
            Caption         =   "+P"
            Height          =   255
            Left            =   5280
            TabIndex        =   35
            Top             =   2880
            Width           =   615
         End
         Begin VB.CheckBox chkCC 
            Caption         =   "+C"
            Height          =   255
            Left            =   3000
            TabIndex        =   34
            Top             =   3960
            Width           =   615
         End
         Begin VB.CheckBox chkCF 
            Caption         =   "+F"
            Height          =   255
            Left            =   3000
            TabIndex        =   33
            Top             =   4320
            Width           =   615
         End
         Begin VB.CheckBox chkCG 
            Caption         =   "+G"
            Height          =   255
            Left            =   3000
            TabIndex        =   32
            Top             =   4680
            Width           =   615
         End
         Begin VB.CheckBox chkCH 
            Caption         =   "+H"
            Height          =   255
            Left            =   3000
            TabIndex        =   31
            Top             =   5040
            Width           =   615
         End
         Begin VB.CheckBox chkCL 
            Caption         =   "+L"
            Height          =   255
            Left            =   5280
            TabIndex        =   30
            Top             =   1440
            Width           =   615
         End
         Begin VB.CheckBox chkCK 
            Caption         =   "+K"
            Height          =   255
            Left            =   5280
            TabIndex        =   29
            Top             =   1080
            Width           =   615
         End
         Begin VB.CheckBox chkCA 
            Caption         =   "+A"
            Height          =   255
            Left            =   3000
            TabIndex        =   28
            Top             =   3600
            Width           =   615
         End
         Begin VB.TextBox txtLimit 
            Height          =   285
            Left            =   1800
            TabIndex        =   27
            Top             =   3600
            Width           =   975
         End
         Begin VB.TextBox txtKey 
            Height          =   285
            Left            =   1800
            TabIndex        =   26
            Top             =   3240
            Width           =   975
         End
         Begin VB.CheckBox chkz 
            Caption         =   "+z"
            Height          =   255
            Left            =   3000
            TabIndex        =   25
            Top             =   3240
            Width           =   615
         End
         Begin VB.CheckBox chki 
            Caption         =   "+i (Invite only)"
            Height          =   255
            Left            =   720
            TabIndex        =   24
            Top             =   2520
            Width           =   1455
         End
         Begin VB.CheckBox chkl 
            Caption         =   "+l (Limit)"
            Height          =   255
            Left            =   720
            TabIndex        =   23
            Top             =   3600
            Width           =   975
         End
         Begin VB.CheckBox chkm 
            Caption         =   "+m (Moderated)"
            Height          =   255
            Left            =   720
            TabIndex        =   22
            Top             =   3960
            Width           =   1575
         End
         Begin VB.CheckBox chkn 
            Caption         =   "+n (No external)"
            Height          =   255
            Left            =   720
            TabIndex        =   21
            Top             =   4320
            Width           =   1575
         End
         Begin VB.CheckBox chkp 
            Caption         =   "+p (Private)"
            Height          =   255
            Left            =   720
            TabIndex        =   20
            Top             =   4680
            Width           =   1215
         End
         Begin VB.CheckBox chkq 
            Caption         =   "+q"
            Height          =   255
            Left            =   720
            TabIndex        =   19
            Top             =   5040
            Width           =   615
         End
         Begin VB.CheckBox chkr 
            Caption         =   "+r"
            Height          =   255
            Left            =   3000
            TabIndex        =   18
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox chks 
            Caption         =   "+s (Secret)"
            Height          =   255
            Left            =   3000
            TabIndex        =   17
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox chkj 
            Caption         =   "+j"
            Height          =   255
            Left            =   720
            TabIndex        =   16
            Top             =   2880
            Width           =   615
         End
         Begin VB.CheckBox chkk 
            Caption         =   "+k (Key)"
            Height          =   255
            Left            =   720
            TabIndex        =   15
            Top             =   3240
            Width           =   975
         End
         Begin VB.CheckBox chky 
            Caption         =   "+y"
            Height          =   255
            Left            =   3000
            TabIndex        =   14
            Top             =   2880
            Width           =   615
         End
         Begin VB.CheckBox chkx 
            Caption         =   "+x"
            Height          =   255
            Left            =   3000
            TabIndex        =   13
            Top             =   2520
            Width           =   615
         End
         Begin VB.CheckBox chkw 
            Caption         =   "+w"
            Height          =   255
            Left            =   3000
            TabIndex        =   12
            Top             =   2160
            Width           =   615
         End
         Begin VB.CheckBox chku 
            Caption         =   "+u"
            Height          =   255
            Left            =   3000
            TabIndex        =   11
            Top             =   1800
            Width           =   615
         End
         Begin VB.CheckBox chkt 
            Caption         =   "+t (Protected topic)"
            Height          =   255
            Left            =   3000
            TabIndex        =   10
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CheckBox chkh 
            Caption         =   "+h"
            Height          =   255
            Left            =   720
            TabIndex        =   9
            Top             =   2160
            Width           =   615
         End
         Begin VB.CheckBox chkg 
            Caption         =   "+g"
            Height          =   255
            Left            =   720
            TabIndex        =   8
            Top             =   1800
            Width           =   615
         End
         Begin VB.CheckBox chkf 
            Caption         =   "+f"
            Height          =   255
            Left            =   720
            TabIndex        =   7
            Top             =   1440
            Width           =   615
         End
         Begin VB.CheckBox chkc 
            Caption         =   "+c"
            Height          =   255
            Left            =   720
            TabIndex        =   6
            Top             =   1080
            Width           =   615
         End
         Begin VB.CheckBox chka 
            Caption         =   "+a"
            Height          =   255
            Left            =   720
            TabIndex        =   5
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtTema 
            Height          =   285
            Left            =   720
            TabIndex        =   4
            Top             =   240
            Width           =   7455
         End
         Begin VB.Label Label1 
            Caption         =   "Tema:"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "frmChanSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Me.Tag = "-"
    If chka.Value = 1 Then Me.Tag = "a" & Me.Tag Else Me.Tag = Me.Tag & "a"
    If chkc.Value = 1 Then Me.Tag = "c" & Me.Tag Else Me.Tag = Me.Tag & "c"
    If chkf.Value = 1 Then Me.Tag = "f" & Me.Tag Else Me.Tag = Me.Tag & "f"
    If chkg.Value = 1 Then Me.Tag = "g" & Me.Tag Else Me.Tag = Me.Tag & "g"
    If chkh.Value = 1 Then Me.Tag = "h" & Me.Tag Else Me.Tag = Me.Tag & "h"
    If chki.Value = 1 Then Me.Tag = "i" & Me.Tag Else Me.Tag = Me.Tag & "i"
    If chkj.Value = 1 Then Me.Tag = "j" & Me.Tag Else Me.Tag = Me.Tag & "j"
    If chkk.Value = 1 Then Me.Tag = "k" & Me.Tag Else Me.Tag = Me.Tag & "k"
    If chkl.Value = 1 Then Me.Tag = "l" & Me.Tag Else Me.Tag = Me.Tag & "l"
    If chkm.Value = 1 Then Me.Tag = "m" & Me.Tag Else Me.Tag = Me.Tag & "m"
    If chkn.Value = 1 Then Me.Tag = "n" & Me.Tag Else Me.Tag = Me.Tag & "n"
    If chkp.Value = 1 Then Me.Tag = "p" & Me.Tag Else Me.Tag = Me.Tag & "p"
    If chkq.Value = 1 Then Me.Tag = "q" & Me.Tag Else Me.Tag = Me.Tag & "q"
    If chkr.Value = 1 Then Me.Tag = "r" & Me.Tag Else Me.Tag = Me.Tag & "r"
    If chks.Value = 1 Then Me.Tag = "s" & Me.Tag Else Me.Tag = Me.Tag & "s"
    If chkt.Value = 1 Then Me.Tag = "t" & Me.Tag Else Me.Tag = Me.Tag & "t"
    If chku.Value = 1 Then Me.Tag = "u" & Me.Tag Else Me.Tag = Me.Tag & "u"
    If chkw.Value = 1 Then Me.Tag = "w" & Me.Tag Else Me.Tag = Me.Tag & "w"
    If chkx.Value = 1 Then Me.Tag = "x" & Me.Tag Else Me.Tag = Me.Tag & "x"
    If chky.Value = 1 Then Me.Tag = "y" & Me.Tag Else Me.Tag = Me.Tag & "y"
    If chkz.Value = 1 Then Me.Tag = "z" & Me.Tag Else Me.Tag = Me.Tag & "z"
    
    If chkCA.Value = 1 Then Me.Tag = "A" & Me.Tag Else Me.Tag = Me.Tag & "A"
    If chkCC.Value = 1 Then Me.Tag = "C" & Me.Tag Else Me.Tag = Me.Tag & "C"
    If chkCF.Value = 1 Then Me.Tag = "F" & Me.Tag Else Me.Tag = Me.Tag & "F"
    If chkCG.Value = 1 Then Me.Tag = "G" & Me.Tag Else Me.Tag = Me.Tag & "G"
    If chkCH.Value = 1 Then Me.Tag = "H" & Me.Tag Else Me.Tag = Me.Tag & "H"
    If chkCJ.Value = 1 Then Me.Tag = "J" & Me.Tag Else Me.Tag = Me.Tag & "J"
    If chkCK.Value = 1 Then Me.Tag = "K" & Me.Tag Else Me.Tag = Me.Tag & "K"
    If chkCL.Value = 1 Then Me.Tag = "L" & Me.Tag Else Me.Tag = Me.Tag & "L"
    If chkCM.Value = 1 Then Me.Tag = "M" & Me.Tag Else Me.Tag = Me.Tag & "M"
    If chkCN.Value = 1 Then Me.Tag = "N" & Me.Tag Else Me.Tag = Me.Tag & "N"
    If chkCO.Value = 1 Then Me.Tag = "O" & Me.Tag Else Me.Tag = Me.Tag & "O"
    If chkCP.Value = 1 Then Me.Tag = "P" & Me.Tag Else Me.Tag = Me.Tag & "P"
    If chkCQ.Value = 1 Then Me.Tag = "Q" & Me.Tag Else Me.Tag = Me.Tag & "Q"
    If chkCR.Value = 1 Then Me.Tag = "R" & Me.Tag Else Me.Tag = Me.Tag & "R"
    If chkCS.Value = 1 Then Me.Tag = "S" & Me.Tag Else Me.Tag = Me.Tag & "S"
    If chkCT.Value = 1 Then Me.Tag = "T" & Me.Tag Else Me.Tag = Me.Tag & "T"
    If chkCU.Value = 1 Then Me.Tag = "U" & Me.Tag Else Me.Tag = Me.Tag & "U"
    If chkCV.Value = 1 Then Me.Tag = "V" & Me.Tag Else Me.Tag = Me.Tag & "V"
    If chkCW.Value = 1 Then Me.Tag = "W" & Me.Tag Else Me.Tag = Me.Tag & "W"
    If chkCX.Value = 1 Then Me.Tag = "X" & Me.Tag Else Me.Tag = Me.Tag & "X"
    If chkCY.Value = 1 Then Me.Tag = "Y" & Me.Tag Else Me.Tag = Me.Tag & "Y"
    If chkCZ.Value = 1 Then Me.Tag = "Z" & Me.Tag Else Me.Tag = Me.Tag & "Z"
    
    Me.Tag = "+" & Me.Tag & " " & txtLimit & " " & txtKey
    Me.Tag = Trim(Me.Tag)
    PutServ "MODE " & TmpChanSetChan & " " & Me.Tag
    Unload Me
End Sub

Private Sub Form_Load()
    If InStr(1, ChannelmodeStr, "a") = 0 Then chka.Enabled = False
    If InStr(1, ChannelmodeStr, "c") = 0 Then chkc.Enabled = False
    If InStr(1, ChannelmodeStr, "f") = 0 Then chkf.Enabled = False
    If InStr(1, ChannelmodeStr, "g") = 0 Then chkg.Enabled = False
    If InStr(1, ChannelmodeStr, "h") = 0 Then chkh.Enabled = False
    If InStr(1, ChannelmodeStr, "i") = 0 Then chki.Enabled = False
    If InStr(1, ChannelmodeStr, "j") = 0 Then chkj.Enabled = False
    If InStr(1, ChannelmodeStr, "k") = 0 Then chkk.Enabled = False
    If InStr(1, ChannelmodeStr, "l") = 0 Then chkl.Enabled = False
    If InStr(1, ChannelmodeStr, "m") = 0 Then chkm.Enabled = False
    If InStr(1, ChannelmodeStr, "n") = 0 Then chkn.Enabled = False
    If InStr(1, ChannelmodeStr, "p") = 0 Then chkp.Enabled = False
    If InStr(1, ChannelmodeStr, "q") = 0 Then chkq.Enabled = False
    If InStr(1, ChannelmodeStr, "r") = 0 Then chkr.Enabled = False
    If InStr(1, ChannelmodeStr, "s") = 0 Then chks.Enabled = False
    If InStr(1, ChannelmodeStr, "t") = 0 Then chkt.Enabled = False
    If InStr(1, ChannelmodeStr, "u") = 0 Then chku.Enabled = False
    If InStr(1, ChannelmodeStr, "w") = 0 Then chkw.Enabled = False
    If InStr(1, ChannelmodeStr, "x") = 0 Then chkx.Enabled = False
    If InStr(1, ChannelmodeStr, "y") = 0 Then chky.Enabled = False
    If InStr(1, ChannelmodeStr, "z") = 0 Then chkz.Enabled = False
    
    If InStr(1, ChannelmodeStr, "A") = 0 Then chkCA.Enabled = False
    If InStr(1, ChannelmodeStr, "C") = 0 Then chkCC.Enabled = False
    If InStr(1, ChannelmodeStr, "F") = 0 Then chkCF.Enabled = False
    If InStr(1, ChannelmodeStr, "G") = 0 Then chkCG.Enabled = False
    If InStr(1, ChannelmodeStr, "H") = 0 Then chkCH.Enabled = False
    If InStr(1, ChannelmodeStr, "J") = 0 Then chkCJ.Enabled = False
    If InStr(1, ChannelmodeStr, "K") = 0 Then chkCK.Enabled = False
    If InStr(1, ChannelmodeStr, "L") = 0 Then chkCL.Enabled = False
    If InStr(1, ChannelmodeStr, "M") = 0 Then chkCM.Enabled = False
    If InStr(1, ChannelmodeStr, "N") = 0 Then chkCN.Enabled = False
    If InStr(1, ChannelmodeStr, "O") = 0 Then chkCO.Enabled = False
    If InStr(1, ChannelmodeStr, "P") = 0 Then chkCP.Enabled = False
    If InStr(1, ChannelmodeStr, "Q") = 0 Then chkCQ.Enabled = False
    If InStr(1, ChannelmodeStr, "R") = 0 Then chkCR.Enabled = False
    If InStr(1, ChannelmodeStr, "S") = 0 Then chkCS.Enabled = False
    If InStr(1, ChannelmodeStr, "T") = 0 Then chkCT.Enabled = False
    If InStr(1, ChannelmodeStr, "U") = 0 Then chkCU.Enabled = False
    If InStr(1, ChannelmodeStr, "V") = 0 Then chkCV.Enabled = False
    If InStr(1, ChannelmodeStr, "W") = 0 Then chkCW.Enabled = False
    If InStr(1, ChannelmodeStr, "X") = 0 Then chkCX.Enabled = False
    If InStr(1, ChannelmodeStr, "Y") = 0 Then chkCY.Enabled = False
    If InStr(1, ChannelmodeStr, "Z") = 0 Then chkCZ.Enabled = False
    
    If Not InStr(1, TmpChanSet, "a") = 0 Then chka.Value = 1
    If Not InStr(1, TmpChanSet, "c") = 0 Then chkc.Value = 1
    If Not InStr(1, TmpChanSet, "f") = 0 Then chkf.Value = 1
    If Not InStr(1, TmpChanSet, "g") = 0 Then chkg.Value = 1
    If Not InStr(1, TmpChanSet, "h") = 0 Then chkh.Value = 1
    If Not InStr(1, TmpChanSet, "i") = 0 Then chki.Value = 1
    If Not InStr(1, TmpChanSet, "j") = 0 Then chkj.Value = 1
    If Not InStr(1, TmpChanSet, "k") = 0 Then chkk.Value = 1
    If Not InStr(1, TmpChanSet, "l") = 0 Then chkl.Value = 1
    If Not InStr(1, TmpChanSet, "m") = 0 Then chkm.Value = 1
    If Not InStr(1, TmpChanSet, "n") = 0 Then chkn.Value = 1
    If Not InStr(1, TmpChanSet, "p") = 0 Then chkp.Value = 1
    If Not InStr(1, TmpChanSet, "q") = 0 Then chkq.Value = 1
    If Not InStr(1, TmpChanSet, "r") = 0 Then chkr.Value = 1
    If Not InStr(1, TmpChanSet, "s") = 0 Then chks.Value = 1
    If Not InStr(1, TmpChanSet, "t") = 0 Then chkt.Value = 1
    If Not InStr(1, TmpChanSet, "u") = 0 Then chku.Value = 1
    If Not InStr(1, TmpChanSet, "w") = 0 Then chkw.Value = 1
    If Not InStr(1, TmpChanSet, "x") = 0 Then chkx.Value = 1
    If Not InStr(1, TmpChanSet, "y") = 0 Then chky.Value = 1
    If Not InStr(1, TmpChanSet, "z") = 0 Then chkz.Value = 1
    
    If Not InStr(1, TmpChanSet, "A") = 0 Then chkCA.Value = 1
    If Not InStr(1, TmpChanSet, "C") = 0 Then chkCC.Value = 1
    If Not InStr(1, TmpChanSet, "F") = 0 Then chkCF.Value = 1
    If Not InStr(1, TmpChanSet, "G") = 0 Then chkCG.Value = 1
    If Not InStr(1, TmpChanSet, "H") = 0 Then chkCH.Value = 1
    If Not InStr(1, TmpChanSet, "J") = 0 Then chkCJ.Value = 1
    If Not InStr(1, TmpChanSet, "K") = 0 Then chkCK.Value = 1
    If Not InStr(1, TmpChanSet, "L") = 0 Then chkCL.Value = 1
    If Not InStr(1, TmpChanSet, "M") = 0 Then chkCM.Value = 1
    If Not InStr(1, TmpChanSet, "N") = 0 Then chkCN.Value = 1
    If Not InStr(1, TmpChanSet, "O") = 0 Then chkCO.Value = 1
    If Not InStr(1, TmpChanSet, "P") = 0 Then chkCP.Value = 1
    If Not InStr(1, TmpChanSet, "Q") = 0 Then chkCQ.Value = 1
    If Not InStr(1, TmpChanSet, "R") = 0 Then chkCR.Value = 1
    If Not InStr(1, TmpChanSet, "S") = 0 Then chkCS.Value = 1
    If Not InStr(1, TmpChanSet, "T") = 0 Then chkCT.Value = 1
    If Not InStr(1, TmpChanSet, "U") = 0 Then chkCU.Value = 1
    If Not InStr(1, TmpChanSet, "V") = 0 Then chkCV.Value = 1
    If Not InStr(1, TmpChanSet, "W") = 0 Then chkCW.Value = 1
    If Not InStr(1, TmpChanSet, "X") = 0 Then chkCX.Value = 1
    If Not InStr(1, TmpChanSet, "Y") = 0 Then chkCY.Value = 1
    If Not InStr(1, TmpChanSet, "Z") = 0 Then chkCZ.Value = 1
End Sub
