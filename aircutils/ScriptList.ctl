VERSION 5.00
Begin VB.UserControl ScList 
   AccessKeys      =   "123456789"
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   178
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   252
   Windowless      =   -1  'True
End
Attribute VB_Name = "ScList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Event RightClick()
Event ScChange(ByVal ScNum As Integer)

Public Type ScDef
Col As Long
Indx As Integer
Ico As StdPicture
Title As String
End Type

Dim Scr() As ScDef, CurSel As Integer, ItemCount As Integer

Private Sub UserControl_Initialize()
CurSel = 1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If ((x < 0) Or (y < 0)) Or ((x > ScaleWidth) Or (y > ScaleHeight)) Then Exit Sub
butnun = Fix((y / 20) + 1)
If (butnun < 1) Or (butnun > ItemCount) Then Exit Sub ' <= Ugyldig
Dim FX As Form
RaiseEvent ScChange(butnun)
CurSel = butnun
Refresh
If Button = 2 Then 'Right click, raise rightclick event
RaiseEvent RightClick
End If
End Sub

Private Sub UserControl_Paint()
For V = 0 To ItemCount - 1
    If CurSel = V + 1 Then
        Line (0, V * 20)-(ScaleWidth - 2, V * 20 + 18), &H0, B
        Line (1, V * 20 + 1)-(ScaleWidth - 1, V * 20 + 19), &HFFFFFF, B
        Line (2, V * 20 + 2)-(ScaleWidth - 3, V * 20 + 17), &H808080, BF
        ForeColor = &HFFFFFF
    Else
        Line (0, V * 20)-(ScaleWidth - 2, V * 20 + 18), &H0, B
        Line (1, V * 20 + 1)-(ScaleWidth - 1, V * 20 + 19), &HFFFFFF, B
        Line (2, V * 20 + 2)-(ScaleWidth - 3, V * 20 + 17), &H8000000F, BF
        ForeColor = &H0
        If Not Scr(V + 1).Col = 0 Then ForeColor = Scr(V + 1).Col
    End If
    PaintPicture Scr(V + 1).Ico, 2, V * 20 + 2
    CurrentX = 18: CurrentY = V * 20 + 3
    Print Scr(V + 1).Title
Next
End Sub

Sub AddScr(ByVal T As String, ByVal nIco As StdPicture)
ItemCount = ItemCount + 1
ReDim Preserve Scr(1 To ItemCount)
With Scr(ItemCount)
.Title = T
.Indx = ItemCount
Set .Ico = nIco
End With
Refresh
End Sub

Sub ActScr(ByVal N As Integer)
If (N < 1) Or (N > ItemCount) Then Exit Sub
CurSel = N
Refresh
End Sub

Sub ColScr(ByVal N As Integer, ByVal Color As Long)
If (N < 1) Or (N > ItemCount) Then Exit Sub
Scr(N).Col = Color
Refresh
End Sub

Sub RenScr(ByVal N As Integer, ByVal Tit As String)
If (N < 1) Or (N > ItemCount) Then Exit Sub
Scr(N).Title = Tit
Refresh
End Sub

Sub RemScr(ByVal N As Integer)
If (N < 1) Or (N > ItemCount) Then Exit Sub
For Z = N To ItemCount - 1
Scr(Z) = Scr(Z + 1)
Next
ItemCount = ItemCount - 1
If ItemCount <= 0 Then
Erase Scr
Else
ReDim Preserve Scr(1 To ItemCount)
End If
Cls
Refresh
End Sub

Sub SetIcon(ByVal N As Integer, ByVal nIco As StdPicture)
If (N < 1) Or (N > ItemCount) Then Exit Sub
Set Scr(N).Ico = nIco
Refresh
End Sub

Sub Refresh()
UserControl_Paint
End Sub

