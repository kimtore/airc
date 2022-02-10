VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2120D62E-1B94-47CE-956E-F31CED1DA6C4}#19.2#0"; "aircutils.ocx"
Begin VB.Form frmScripts 
   Caption         =   "Scripts"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScripts.frx":0000
   LinkTopic       =   "frmScripts"
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   7620
   Begin VB.CommandButton cmdUnloadAll 
      Caption         =   "Unload all"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoadDir 
      Caption         =   "Load dir"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgLoad 
      Left            =   720
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load script"
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdUnload 
      Caption         =   "Unload script"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add script"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin MSComctlLib.ListView listScripts 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   13816530
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Script name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Script function"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Script status"
         Object.Width           =   1940
      EndProperty
   End
   Begin aircutils.KeyFetch KeyFetch1 
      Height          =   465
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   -5160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   820
   End
End
Attribute VB_Name = "frmScripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsHidden As Boolean
Public autoAdd As String
Dim script_used As Integer

Sub DoAdd(ByVal s As String)
    autoAdd = s
    cmdAdd_Click
End Sub

Private Sub cmdAdd_Click()
    Dim FF As Integer, s As String, FName As String
    Dim ScName As Variant, ScFunc As Variant
    On Error Resume Next
    With dlgLoad
        If Not autoAdd = "" Then
            .Filename = autoAdd
            autoAdd = ""
        Else
            .Filter = "VBScript files|*.vbs"
            .CancelError = True
            .InitDir = App.Path
            .DefaultExt = "vbs"
            frmMain.ToggleBlock True
            .ShowOpen
            frmMain.ToggleBlock False
            If Err <> 0 Then Exit Sub
        End If
        On Error GoTo 0
        FF = FreeFile
        If .Filename = TrimPath(.Filename) Then
            If Not Right(App.Path, 1) = "\" Then .Filename = "\" & .Filename
            .Filename = App.Path & .Filename
        End If
        If Not IsValidFile(.Filename) Then KillFile .Filename: Exit Sub
        With listScripts.ListItems
            .Add , , dlgLoad.Filename
            .Item(.Count).SubItems(2) = "loading..."
            listScripts.SelectedItem = .Item(.Count)
        End With
    End With
    If listScripts.SelectedItem Is Nothing Then Exit Sub
    Dim V As classScript
    Set V = New classScript
    Inc script_used
    Load frmMain.airc_Sc(script_used)
    frmMain.airc_Sc(script_used).Reset
    frmMain.airc_Sc(script_used).AddObject "airc", V, True
    FF = FreeFile
    FName = listScripts.SelectedItem
    Open FName For Binary Access Read Shared As #FF
    s = String(LOF(FF), " ")
    Get #FF, , s
    Close #FF
    On Error Resume Next
    With frmMain.airc_Sc(script_used)
        .AddCode s
        .Run "airc_init", ScName, ScFunc
        If Err <> 0 Then
            frmMain.ToggleBlock True
            MsgBox "Script error: cannot start sub 'airc_init', aborting...", vbCritical + vbMsgBoxHelpButton, , "airc.chm", 1001
            frmMain.ToggleBlock False
            .Reset
            Unload frmMain.airc_Sc(script_used)
            Dec script_used
            Set V = Nothing
            listScripts.ListItems.Remove listScripts.SelectedItem.Index
        Else
            With listScripts.SelectedItem
                AddScript V, .Text, CStr(ScName), CStr(ScFunc), frmMain.airc_Sc(script_used), s
                .Text = ScName
                .ListSubItems(1) = ScFunc
                .ListSubItems(2) = "Active"
            End With
        End If
    End With
    On Error GoTo 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdLoadDir_Click()
    Dim D As String
    Dim F As String
    Dim F2() As Variant
    Dim F2C As Integer
    Dim C As Long
    D = App.Path
    If Not Right(D, 1) = "\" Then D = D & "\"
    F = Dir(D)
    ReDim F2(1 To 1)
    Do While Not F = ""
        If Len(F) > 4 Then
            If Right(F, 4) = ".vbs" Then
                Inc F2C
                ReDim Preserve F2(1 To F2C)
                F2(F2C) = F
            End If
        End If
        F = Dir()
    Loop
    If F2C = 0 Then
        frmMain.ToggleBlock True
        MsgBox "There is no scripts in the program directory.", vbExclamation, "Load scripts"
        frmMain.ToggleBlock False
        Exit Sub
    Else
        frmMain.ToggleBlock True
        If MsgBox("Warning: there is " & F2C & " scripts in program directory." & vbCrLf & "Load all scripts?", vbYesNo + vbExclamation, "Load scripts") = vbNo Then frmMain.ToggleBlock False: Exit Sub
        frmMain.ToggleBlock False
        For C = 1 To F2C
            DoAdd F2(C)
        Next
    End If
End Sub

Private Sub cmdUnload_Click()
    If listScripts.SelectedItem Is Nothing Then Exit Sub
    frmMain.ToggleBlock True
    If MsgBox("Unload script '" & listScripts.SelectedItem & "'?", vbExclamation + vbYesNo, "Unload script") = vbNo Then frmMain.ToggleBlock False: Exit Sub
    frmMain.ToggleBlock False
    RemoveScript listScripts.SelectedItem.Index
End Sub

Private Sub cmdUnloadAll_Click()
    If ScriptArrayU = 0 Then
        frmMain.ToggleBlock True
        MsgBox "There is no scripts loaded.", vbExclamation, "Unload scripts"
        frmMain.ToggleBlock False
        Exit Sub
    Else
        frmMain.ToggleBlock True
        If MsgBox("Warning: there is " & ScriptArrayU & " scripts loaded." & vbCrLf & "Unload all loaded scripts?", vbYesNo + vbExclamation, "Unload scripts") = vbNo Then frmMain.ToggleBlock False: Exit Sub
        frmMain.ToggleBlock False
        Do While ScriptArrayU > 0
            RemoveScript 1
        Loop
    End If
End Sub

Private Sub Form_Activate()
    If IsHidden Then
        With frmMain.WSwitch
            .AddWnd Me, 1
            .ActWnd Me
        End With
    End If
    IsHidden = False
    WindowState = 0
End Sub

Private Sub Form_Load()
    IsHidden = True
    Top = 0
    Left = 0
    Width = frmMain.ScaleWidth
    Height = frmMain.ScaleHeight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If ((UnloadMode = 0) Or (UnloadMode = 1)) Then 'cancel + hide
        Cancel = True
        IsHidden = True
        Hide
        With frmMain.WSwitch
            .RemWnd Me
            .ActWnd fActive
        End With
    End If
End Sub

Private Sub Form_Resize()
    Dim C As Long
    If Not WindowState = 0 Then Exit Sub
    C = cmdLoadDir.Width + cmdUnloadAll.Width + 105 + cmdAdd.Width + cmdUnload.Width + cmdClose.Width
    If Width < C Then
        Width = C
        Exit Sub
    End If
    On Error GoTo ErrHndl
    listScripts.Height = Me.ScaleHeight - cmdAdd.Height
    With cmdAdd
        .Top = listScripts.Height
        cmdUnload.Top = .Top
        cmdClose.Top = .Top
        cmdUnloadAll.Top = .Top
        cmdLoadDir.Top = .Top
    End With
    listScripts.Width = Me.ScaleWidth
    cmdClose.Left = Me.ScaleWidth - cmdClose.Width
    cmdUnload.Left = cmdClose.Left - cmdUnload.Width + 10
    cmdAdd.Left = cmdUnload.Left - cmdAdd.Width + 20
    cmdUnloadAll.Left = cmdAdd.Left - 1320
    cmdLoadDir.Left = cmdUnloadAll.Left - cmdLoadDir.Width + 20
    With listScripts
        .ColumnHeaders(2).Width = .Width - (.ColumnHeaders(1).Width + .ColumnHeaders(3).Width + 80)
    End With
ErrHndl:
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    With frmMain.WSwitch
        .RemWnd Me
    End With
End Sub

Sub AddScript(V As classScript, File_Name As String, Sc_Name As String, Sc_Func As String, ScrCtl As ScriptControl, Code As String)
    Inc ScriptArrayU
    ReDim Preserve ScriptArray(1 To ScriptArrayU)
    With ScriptArray(ScriptArrayU)
        Set .V = V
        .File_Name = File_Name
        .Sc_Name = Sc_Name
        .Sc_Func = Sc_Func
        .Code = Code
        Set .ScrCtl = ScrCtl
    End With
End Sub

Sub RemoveScript(ByVal ScriptNumber As Integer)
    Dim C As Long
    If ((ScriptNumber < 1) Or (ScriptNumber > ScriptArrayU)) Then Exit Sub
    On Error Resume Next
    ScriptArray(ScriptNumber).ScrCtl.Run "airc_close"
    If Err <> 0 Then Err.Clear
    On Error GoTo 0
    Unload ScriptArray(ScriptNumber).ScrCtl
    For C = ScriptNumber To ScriptArrayU - 1
        ScriptArray(C) = ScriptArray(C + 1)
    Next
    Dec ScriptArrayU
    If ScriptArrayU < 1 Then
        ScriptArrayU = 0
        Erase ScriptArray
    Else
        ReDim Preserve ScriptArray(1 To ScriptArrayU)
    End If
    listScripts.ListItems.Remove ScriptNumber
End Sub

Private Sub KeyFetch1_ChangeWindow(ByVal WindowNum As Long)
    frmMain.WSwitch.NumWnd WindowNum
End Sub

Private Sub listScripts_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim O As ListItem
    Set O = listScripts.HitTest(x, y)
    If O Is Nothing Then
        Set listScripts.SelectedItem = Nothing
    Else
        Set listScripts.SelectedItem = O
        If Button = 2 Then
            Me.PopupMenu frmMain.mnuscript
        End If
    End If
End Sub

