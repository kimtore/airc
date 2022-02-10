VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2120D62E-1B94-47CE-956E-F31CED1DA6C4}#18.5#0"; "aircutils.ocx"
Begin VB.Form frmDCCStatus 
   Caption         =   "DCC Transfers"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDCCStatus.frx":0000
   LinkTopic       =   "frmDCCStatus"
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   10230
   Begin MSComctlLib.ListView listTransfers 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   13816530
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "From/To"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Sent"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Received"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Speed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "%"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Time elapsed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Time left"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Direction"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin aircutils.KeyFetch KeyFetch1 
      Height          =   465
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   820
   End
End
Attribute VB_Name = "frmDCCStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    frmMain.WSwitch.ActWnd Me
    WindowState = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then 'alt-knappen
        If ChkWndChange(KeyCode) Then
            KeyCode = 0
            Shift = 0
            Exit Sub
       End If
    End If
End Sub

Private Sub Form_Load()
    dcStat = True
    Me.Left = 0
    Me.Top = 0
    Me.Width = frmMain.ScaleWidth
    Me.Height = frmMain.ScaleHeight
    frmMain.WSwitch.AddWnd Me, ActiveServer
    With listTransfers
        .ColumnHeaders(1).Width = 1800
        .ColumnHeaders(2).Width = 1155
        .ColumnHeaders(3).Width = 990
        .ColumnHeaders(4).Width = 990
        .ColumnHeaders(5).Width = 990
        .ColumnHeaders(6).Width = 990
        .ColumnHeaders(7).Width = 615
        .ColumnHeaders(8).Width = 1395
        .ColumnHeaders(9).Width = 1395
        .ColumnHeaders(10).Width = 810
        .ColumnHeaders(11).Width = 5805
    End With
    Dim C As Long
    With listTransfers.ListItems
        For C = 1 To DCCWndU
            AddDCC C
        Next
    End With
End Sub

Private Sub Form_Resize()
    If Not WindowState = 0 Then Exit Sub
    listTransfers.Width = ScaleWidth
    listTransfers.Height = ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.WSwitch.RemWnd Me
    frmMain.WSwitch.ActWnd fActive
    dcStat = False
End Sub

Private Sub KeyFetch1_ChangeWindow(ByVal WindowNum As Long)
    frmMain.WSwitch.NumWnd WindowNum
End Sub

Private Sub listTransfers_DblClick()
    If listTransfers.SelectedItem Is Nothing Then Exit Sub
    With DCCWnd(listTransfers.SelectedItem.Tag)
        .SetFocus
    End With
End Sub

Private Sub listTransfers_KeyDown(KeyCode As Integer, Shift As Integer)
    ChkFunction KeyCode
End Sub

Private Sub listTransfers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set listTransfers.SelectedItem = listTransfers.HitTest(X, Y)
End Sub
