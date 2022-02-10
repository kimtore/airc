VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLinks 
   Caption         =   "Server links"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLinks.frx":0000
   LinkTopic       =   "frmLinks"
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   9900
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Check"
            Object.ToolTipText     =   "Check server links"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Export"
            Object.ToolTipText     =   "Export to text file"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8160
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinks.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinks.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinks.frx":087E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinks.frx":0C18
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinks.frx":0FB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLinks.frx":134C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwLinks 
      Height          =   4935
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8705
      _Version        =   393217
      Indentation     =   564
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text

Private Type SList
Host As String
Uplink As String
Desc As String
proto As String
End Type

Dim list() As SList, Servers As Long
Public CurServer As String

Dim ExportS As String 'Export string
Dim ExportL As Long 'Export level

Sub AddCon(ByVal NSrv As String, ByVal UpSrv As String, ByVal Desc As String)
    Servers = Servers + 1
    ReDim Preserve list(1 To Servers)
    list(Servers).Host = NSrv
    list(Servers).Uplink = UpSrv
    list(Servers).Desc = VBA.Split(Desc, " ", 2)(1)
    'list(Servers).Proto = Proto
End Sub

Sub MMap()
    If Servers = 0 Then
        Dim x As Node
        Set x = tvwLinks.Nodes.Add(, , , "Failed: network in sandbox mode")
        x.Tag = CurServer & "||"
    Else
        MakeTree CurServer
        Erase list
        Servers = 0
    End If
    Show
    frmLinks.SetFocus
    CurServer = ""
End Sub

Function AddConT(ByVal NSrv As String, ByVal UpSrv As String, ByVal proto As String, ByVal Desc As String) As Node
    Dim x As Node
    Dim C As Long
    If NSrv = UpSrv Then
        tvwLinks.Nodes.Clear
        Set x = tvwLinks.Nodes.Add
    Else
        For C = 1 To tvwLinks.Nodes.Count
            Set x = tvwLinks.Nodes.Item(C)
            If VBA.Split(x.Tag, "|")(0) = UpSrv Then Exit For
        Next
        Set x = tvwLinks.Nodes.Add(x, tvwChild)
    End If
    x.Expanded = True
    x.Text = NSrv & " | " & Desc
    x.Tag = NSrv & "|" & proto & "|" & Desc
    
    x.Image = 3
    If NSrv Like "*service*" Then
    x.Image = 6
    End If
    If NSrv Like "[*]*" Then
    x.Image = 5
    End If
    If NSrv = CurServer Then
    x.Image = 1
    End If
    
    Set AddConT = x
End Function

Sub MakeTree(Srv As String)
    Dim C As Long, Index As Integer, x As Node, IsH As Boolean
    For C = 1 To Servers
        If list(C).Host = Srv Then Index = C: Exit For
    Next
    If Index = 0 Then Exit Sub
    Set x = AddConT(list(Index).Host, list(Index).Uplink, list(Index).proto, list(Index).Desc)
    IsH = False
    For C = 1 To Servers
        If C <> Index Then If list(C).Uplink Like list(Index).Host Then IsH = True: MakeTree list(C).Host
    Next
    If IsH And x.Image = 3 Then x.Image = 4
End Sub

Private Sub Form_Activate()
frmMain.WSwitch.ActWnd Me
End Sub

Private Sub Form_Load()
frmMain.WSwitch.AddWnd Me, 0, wndOther
frmMain.WSwitch.ActWnd Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If (UnloadMode = vbFormControlMenu) And (CurServer <> "") Then Cancel = True
End Sub

Private Sub Form_Resize()
tvwLinks.Move 0, Toolbar1.Height, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.WSwitch.RemWnd Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case LCase(Button.Key)
    Case "check"
        PutServ "LINKS"
    Case "export"
        If tvwLinks.Nodes.Count = 0 Then Exit Sub 'No nodes (fixed, 130-b3)
        ExportL = 0
        ExportS = ""
        Export tvwLinks.Nodes(1)
        Clipboard.Clear
        Clipboard.SetText ExportS, vbCFText
End Select
End Sub


Private Sub Export(ByVal N As Node)
Dim C As Long
Dim M As Node

If ExportL = 0 Then
ExportS = ExportS & " " & N.Text & vbCrLf
Else
ExportS = ExportS & StrFill(ExportL - 1, "  | ") & "  |- " & N.Text & vbCrLf
End If
            
ExportL = ExportL + 1

If N.Children = 0 Then
    ExportS = ExportS & StrFill(ExportL - 1, "  | ") & "  |- " & M.Text & vbCrLf
Else
    Set M = N.Child
    Do
        
        If M.Children > 0 Then
            Export M
            Set M = M.Next
            If M Is Nothing Then Exit Do
        End If
        
        If M Is N.Child.LastSibling Then
            ExportS = ExportS & StrFill(ExportL - 1, "  | ") & "  `- " & M.Text & vbCrLf
        Else
            ExportS = ExportS & StrFill(ExportL - 1, "  | ") & "  |- " & M.Text & vbCrLf
        End If
        
        Set M = M.Next
        If M Is Nothing Then Exit Do
    Loop While Not M Is Nothing
    'If Not M Is Nothing Then ExportS = ExportS & Space(ExportL + 1) & "`- " & M.Text & vbCrLf
End If


ExportL = ExportL - 1
End Sub

Private Function StrFill(N As Long, ByVal s As String) As String
Dim C As Long
For C = 1 To N: StrFill = StrFill & s: Next
End Function
