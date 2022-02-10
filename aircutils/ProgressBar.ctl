VERSION 5.00
Begin VB.UserControl ProgressBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4185
   DrawMode        =   11  'Not Xor Pen
   ScaleHeight     =   1845
   ScaleWidth      =   4185
   ToolboxBitmap   =   "ProgressBar.ctx":0000
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   207
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Event Click()
Dim LastR As Integer, LastS As Integer, LastB As Integer
Dim RcVal As Double, SnVal As Double, BfVal As Double
Dim OldP As String
Dim MinVal As Double, MaxVal As Double
Dim BgColor As OLE_COLOR, SnColor As OLE_COLOR, RcColor As OLE_COLOR, BfColor As OLE_COLOR

Private Sub Picture1_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
BgColor = vbWhite
SnColor = vbRed
RcColor = vbBlue
BfColor = vbGreen
End Sub

Private Sub UserControl_Resize()
Picture1.Width = ScaleWidth
Picture1.Height = ScaleHeight
End Sub

Public Sub SetMin(ByVal nValue As Double)
MinVal = nValue
Refresh
End Sub

Public Sub SetMax(ByVal nValue As Double)
MaxVal = nValue
Refresh
End Sub

Public Sub SetRc(ByVal nValue As Double)
RcVal = nValue
Refresh
End Sub

Public Sub SetSn(ByVal nValue As Double)
SnVal = nValue
Refresh
End Sub

Public Sub SetBf(ByVal nValue As Double)
BfVal = nValue
Refresh
End Sub

Public Property Let SentColor(Color As OLE_COLOR)
SnColor = Color
Refresh
End Property

Public Property Get SentColor() As OLE_COLOR
SentColor = SnColor
End Property

Public Property Let RecvColor(Color As OLE_COLOR)
RcColor = Color
Refresh
End Property

Public Property Get RecvColor() As OLE_COLOR
RecvColor = RcColor
End Property

Public Property Let BufColor(Color As OLE_COLOR)
BfColor = Color
Refresh
End Property

Public Property Get BufColor() As OLE_COLOR
BufColor = BfColor
End Property

Public Property Let BackColor(Color As OLE_COLOR)
BgColor = Color
Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = BgColor
End Property

Private Sub Refresh()
    If MaxVal = 0 Then Exit Sub
    Dim MuX As Integer, ImH As Integer, CR As Integer, Cs As Integer, CB As Integer
    MuX = Picture1.ScaleWidth - 1
    ImH = Picture1.ScaleHeight - 1
    CR = (RcVal / MaxVal) * MuX
    Cs = (SnVal / MaxVal) * MuX
    CB = (BfVal / MaxVal) * MuX
    If CR > LastR Then Picture1.Line (LastR, 0)-(CR, ImH), RcColor, BF
    If CR + 1 > LastS Then LastS = CR + 1
    If Cs > LastS Then Picture1.Line (LastS, 0)-(Cs, ImH), SnColor, BF
    If Cs + 1 > LastB Then LastB = Cs + 1
    If CB > LastB Then Picture1.Line (LastB, 0)-(CB, ImH), BfColor, BF
    LastR = CR
    LastS = Cs
    LastB = CB
    DrawPercentage
End Sub

Private Sub Picture1_Paint()
    Dim ImH As Integer
    Dim ImW As Integer
    ImH = Picture1.ScaleHeight - 1
    ImW = Picture1.ScaleWidth - 1
    Picture1.Line (0, 0)-(ImW, ImH), BgColor, BF
    If LastR > 0 Then Picture1.Line (0, 0)-(LastR, ImH), RcColor, BF
    If LastR < LastS Then Picture1.Line (LastR, 0)-(LastS, ImH), SnColor, BF
    If LastS < LastB Then Picture1.Line (LastS, 0)-(LastB, ImH), BfColor, BF
    DrawPercentage 'Must be placed *last*
End Sub

Private Sub DrawPercentage()
    Dim ImW As Integer
    Dim ImH As Integer
    ImH = Picture1.ScaleHeight - 1
    ImW = Picture1.ScaleWidth - 1
    Dim P As String
    P = Percentage
    If Not P = OldP Then 'Must picture1_Paint, not to fuck up percent counter
        OldP = P
        Picture1_Paint
    End If
    With Picture1
        .CurrentX = (.ScaleWidth \ 2) - (.TextWidth(P) \ 2)
        .CurrentY = (.ScaleHeight \ 2) - (.TextHeight(P) \ 2)
        Picture1.Print P
    End With
End Sub

Public Function Percentage() As String
    If MaxVal = 0 Then 'Prevent division by zero error
        Percentage = "0"
    Else
        Percentage = Format(100 - (1 - (RcVal / MaxVal)) * 100, "##")
    End If
    Percentage = Percentage & "%"
    If Percentage = "%" Then Percentage = "0%"
End Function


