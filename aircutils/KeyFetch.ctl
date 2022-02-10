VERSION 5.00
Begin VB.UserControl KeyFetch 
   AccessKeys      =   "1234567890qwertyuiop"
   BackStyle       =   0  'Transparent
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   1620
   ScaleWidth      =   2955
   Windowless      =   -1  'True
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KeyFetch"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1365
   End
End
Attribute VB_Name = "KeyFetch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Event ChangeWindow(ByVal WindowNum As Long)

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
Dim ActWnd As Long
Select Case KeyAscii
    Case 49, 50, 51, 52, 53, 54, 55, 56, 57
        ActWnd = KeyAscii - 48
    Case 48
        ActWnd = 10
    Case Else 'Check letters
        Select Case Chr(KeyAscii)
            Case "Q", "q"
                ActWnd = 11
            Case "W", "w"
                ActWnd = 12
            Case "E", "e"
                ActWnd = 13
            Case "R", "r"
                ActWnd = 14
            Case "T", "t"
                ActWnd = 15
            Case "Y", "y"
                ActWnd = 16
            Case "U", "u"
                ActWnd = 17
            Case "I", "i"
                ActWnd = 18
            Case "O", "o"
                ActWnd = 19
            Case "P", "p"
                ActWnd = 20
            Case Else
        End Select
End Select
If ActWnd > 0 Then RaiseEvent ChangeWindow(ActWnd)
End Sub

Private Sub UserControl_Resize(): Size Label1.Width + 90, Label1.Height + 90: End Sub

