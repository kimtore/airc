VERSION 5.00
Begin VB.Form frmURLList 
   Caption         =   "URL List"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmURLList.frx":0000
   LinkTopic       =   "frmURLList"
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox listURL 
      BackColor       =   &H00D2D2D2&
      Height          =   6810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "frmURLList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Dim C As Long
    frmMain.WSwitch.ActWnd Me
    listURL.Clear
    For C = 1 To URLCount
        listURL.AddItem "* " & URLList(C)
    Next
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
    Me.Left = 0
    Me.Top = 0
    Me.Width = frmMain.ScaleWidth
    Me.Height = frmMain.ScaleHeight
    frmMain.WSwitch.AddWnd Me, ActiveServer
End Sub

Private Sub Form_Resize()
    If Not WindowState = 0 Then Exit Sub
    listURL.Width = ScaleWidth
    listURL.Height = ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.WSwitch.RemWnd Me
    frmMain.WSwitch.ActWnd fActive
End Sub

Private Sub listURL_DblClick()
    ShellExecute hwnd, vbNullString, Mid(listURL.List(listURL.ListIndex), 3), vbNull, vbNullString, 0
End Sub

Private Sub listURL_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then 'alt-knappen
        If ChkWndChange(KeyCode) Then
            KeyCode = 0
            Shift = 0
            Exit Sub
       End If
    End If
    ChkFunction KeyCode
End Sub
