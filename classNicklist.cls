VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classNicklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Chan_ As Long
Private Count_ As Long
Private User_() As TypeUser

Public Sub Init(ByVal ChanInit As Long)
    Erase User_
    Count_ = 0
    Chan_ = ChanInit
End Sub

Public Sub AddN(ByVal Nick As String, ByVal Hostmask As String)
    Inc Count_
    ReDim Preserve User_(1 To Count_)
    With User_(Count_)
        .Nick = TrimMode(Nick)
        .Host = Hostmask
        .ModeString = GetModeString(Nick)
        ChannelWnd(Chan_).listNick.ListItems.Add , , GetActiveMode(.ModeString) & TrimMode(Nick)
    End With
End Sub

Public Sub RemoveN(ByVal Nick As String)
    Dim C As Long
    C = UserPos(Nick)
    If NotVal(C) Then Exit Sub
    Dec Count_
    For C = C To Count_
        User_(C) = User_(C + 1)
    Next
    If Count_ = 0 Then Erase User_ Else ReDim Preserve User_(1 To Count_)
    If ChannelWnd(Chan_).FindNickPos(Nick) = 0 Then Exit Sub
    ChannelWnd(Chan_).listNick.ListItems.Remove ChannelWnd(Chan_).FindNickPos(Nick)
End Sub

Public Sub ReplaceN(ByVal Nick As String, NewNick As String, Optional ByVal AddOp As Boolean, Optional ByVal SubtractOp As Boolean, Optional ByVal AddVoice As Boolean, Optional ByVal SubtractVoice As Boolean)
    Dim C As Long
    Dim Usr As TypeUser
    C = UserPos(Nick)
    If NotVal(C) Then Exit Sub
    With User_(C)
         .Nick = NewNick
         If SubtractOp Or AddOp Then .ModeString = Replace(.ModeString, "@", "")
         If SubtractVoice Or AddVoice Then .ModeString = Replace(.ModeString, "+", "")
         If AddOp Then .ModeString = .ModeString & "@"
         If AddVoice Then .ModeString = .ModeString & "+"
    End With
    With ChannelWnd(Chan_)
        If .FindNickPos(Nick) = 0 Then Exit Sub
        With .listNick.ListItems(.FindNickPos(Nick))
            .Text = GetActiveMode(User_(C).ModeString) & TrimMode(NewNick)
        End With
        .listNick.Sorted = True
    End With
End Sub

Public Sub SetChan(ByVal N As Long)
    Chan_ = N
End Sub

Public Function User_Nick(ByVal N As Long) As String
    If NotVal(N) Then Exit Function
    User_Nick = User_(N).Nick
End Function

Public Function User_Host(ByVal N As Long) As String
    If NotVal(N) Then Exit Function
    User_Host = User_(N).Host
End Function

Public Function User_Modes(ByVal N As Long) As String
    If NotVal(N) Then Exit Function
    User_Modes = User_(N).ModeString
End Function

Public Function User_Tag(ByVal N As Long) As String
    If NotVal(N) Then Exit Function
    User_Tag = User_(N).Tag
End Function

Private Function NotVal(ByVal N As Long) As Boolean
    NotVal = True
    If N > Count_ Then Exit Function
    If N < 1 Then Exit Function
    NotVal = False
End Function

Public Function UserPos(ByVal U As String) As Long
    Dim C As Long
    For C = 1 To Count_
        If LCase(U) = LCase(User_(C).Nick) Then Exit For
    Next
    If C > Count_ Then C = 0
    UserPos = C
End Function

Public Sub SetHost(ByVal N As Long, ByVal S As String)
    If NotVal(N) Then Exit Sub
    User_(N).Host = S
End Sub

Public Function Count() As Long 'Prevent setting count
    Count = Count_
End Function

Public Function IsOp(ByVal N As Long) As Boolean
    Dim C As Long
    If NotVal(N) Then Exit Function
    If InStr(1, User_(N).ModeString, "@") > 0 Then IsOp = True
End Function

Function IsVoice(ByVal N As Long) As Boolean
    Dim C As Long
    If NotVal(N) Then Exit Function
    If InStr(1, User_(N).ModeString, "+") > 0 Then IsVoice = True
End Function

