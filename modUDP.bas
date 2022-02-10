Attribute VB_Name = "modUDP"
'// AUDPDCCFT Protocol //
'// Currently in planning stage //

#Const USE_UDP = False

#If USE_UDP Then

Type UDP_Info
    FileName As String
    Filesize As Currency
    IP As String
    ID As Long
    PSize As Long 'Packet size
    PReceived() As Boolean 'Array of packets received
End Type

Public UDP_C() As UDP_Info
Public UDP_CU As Long

Sub ParseUDP(ByVal S As String, ByVal IP As String)
Dim V As Variant
If Len(S) = 0 Then Exit Sub
Dim M As Byte
M = Asc(Left(S, 1))
Select Case M
V = Split(S, " ")
Case 0 'UDP request (receiver)
SendUDP Chr(1), IP

Case 1 'UDP request acknowledged (sender)
On Error Resume Next
With frmMain
.ToggleBlock True
.cdSend.ShowOpen
.ToggleBlock False
End With
If Not Err.Number = 0 Then
Err.Clear
On Error GoTo 0
Exit Sub
End If
On Error GoTo 0
With NewUDP(frmMain.cdSend.FileName, FileLen(frmMain.cdSend.FileName), DCCIP, Inc(DCCUnique), DCCInfo.SendeBuffer)
SendUDP Chr(2) & " " & .ID & " " & .FileName & " " & .Filesize & " " & .PSize
End With

Case 2 'File request (receiver)
NewDCCWnd IP, V(2), V(3), IP, frmMain.sckUDP.RemotePort, False, False, V(1), True

Case 3
Case 4
Case 5
Case 6
Case 7
Case 8
End Select
End Sub

Sub SendUDP(ByVal S As String, ByVal IP As String)
With frmMain.sckUDP
.RemoteHost = IP
.SendData S
End With
End Sub

Function NewUDP(ByVal FileName As String, ByVal Filesize As Currency, ByVal IP As String, ByVal ID As Long, ByVal PSize As Long) As UDP_Info
Inc UDP_CU
ReDim Preserve UDP_C(1 To UDP_CU)
With UDP_C(UDP_CU)
.FileName = FileName
.Filesize = Filesize
.IP = IP
.ID = ID
.PSize = PSize
ReDim .PReceived(1 To (.Filesize \ .PSize) + 1)
End With
End Function

Sub DelUDP(ByVal IP As String, ByVal ID As Long)
Dim C As Long
For C = 1 To UDP_CU
If (UDP_C(C).ID = ID) And (UDP_C(C).IP = IP) Then Exit For
Next
If C <= UDP_CU Then 'Delete
Dec UDP_CU
If UDP_CU = 0 Then Erase UDP_C Else ReDim Preserve UDP_C(1 To UDP_CU)
End If
End Sub

Function FindUDP(ByVal IP As String, ByVal ID As Long) As UDP_Info
Dim C As Long
For C = 1 To UDP_CU
If (UDP_C(C).ID = ID) And (UDP_C(C).IP = IP) Then Exit For
Next
If C <= UDP_CU Then FindUDP = UDP_C(C)
End Function

#End If
