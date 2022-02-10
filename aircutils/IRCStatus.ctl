VERSION 5.00
Begin VB.UserControl IRCStatus 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3390
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1860
   ScaleWidth      =   3390
   ToolboxBitmap   =   "IRCStatus.ctx":0000
   Windowless      =   -1  'True
End
Attribute VB_Name = "IRCStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Nick As String, Server As String, Away As String, Idle As String, Modes As String, Lag As String

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Refresh
End Sub

Sub ChangeNick(newNick As String)
    Nick = newNick
    Refresh
End Sub

Sub ChangeServer(newServer As String)
    Server = newServer
    Refresh
End Sub

Sub ChangeAway(newAway As String)
    Away = newAway
    Refresh
End Sub

Sub ChangeIdle(newIdle As String)
    Idle = newIdle
    Refresh
End Sub

Sub ChangeModes(newModes As String)
    Modes = newModes
    Refresh
End Sub

Sub ChangeLag(newLag As String)
    Lag = newLag
    Refresh
End Sub

Function GetNick() As String
    GetNick = Nick
End Function

Function GetServer() As String
    GetServer = Server
End Function

Function GetAway() As String
    GetAway = Away
End Function

Function GetIdle() As String
    GetIdle = Idle
End Function

Function GetModes() As String
    GetModes = Modes
End Function

Function GetLag() As String
    GetLag = Lag
End Function

Sub Refresh()
    UserControl.Cls
    UserControl_Paint
End Sub

Sub Reset()
    Nick = ""
    Server = ""
    Away = ""
    Idle = ""
    Modes = ""
    Lag = ""
    Refresh
End Sub

Private Sub UserControl_Paint()
    Dim X1, X2, X3, Y1, Y2
    X1 = 100
    X2 = X1 + 4000
    X3 = X2 + 4000
    Y1 = 0
    Y2 = 225
    DoPrint "Nick", Nick, X1, Y1
    DoPrint "Server", Server, X2, Y1
    DoPrint "Away", Away, X3, Y1
    DoPrint "Idle", Idle, X1, Y2
    DoPrint "Modes", Modes, X2, Y2
    DoPrint "Lag", Lag, X3, Y2
End Sub

Private Sub DoPrint(ByVal VarName As String, ByVal Var As String, ByVal XPos As Integer, ByVal YPos As Integer)
    CurrentX = XPos
    CurrentY = YPos
    Print VarName;
    UserControl.FontUnderline = True
    Print "(";
    UserControl.FontUnderline = False
    UserControl.FontBold = True
    Print Var;
    UserControl.FontBold = False
    UserControl.FontUnderline = True
    Print ")";
    UserControl.FontUnderline = False
End Sub
