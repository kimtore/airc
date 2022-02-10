Attribute VB_Name = "modIPConv"
'Thanks to ESE for this piece of code

Option Explicit


Function GetIP(ByVal S As String) As String
On Error GoTo F2
Dim SIP As Double, V As Byte, IPN As Byte
Dim RS As String
SIP = CDbl(S)
For V = 1 To 4
SIP = SIP / 256
IPN = (SIP - Fix(SIP)) * 256
RS = "." + CStr(IPN) + RS
SIP = Fix(SIP)
Next
GetIP = Mid(RS, 2)
Exit Function
F2:
GetIP = S
End Function

Function PutIP(ByVal S As String) As String
On Error GoTo F2
Dim DIP As Double, V As Byte, IPN As Byte, C As Integer
Dim RS As String, L As Integer, LL As Integer
LL = 1
For C = 1 To 3
L = InStr(LL, S, ".", vbTextCompare)
V = CDbl(Mid(S, LL, L - LL))
DIP = (DIP * 256) + V
LL = L + 1
Next
V = CDbl(Mid(S, LL))
DIP = (DIP * 256) + V
PutIP = CStr(DIP)
Exit Function
F2:
PutIP = S
End Function

Function PutLong(L As Long) As String
Dim BA() As Byte
ReDim BA(0 To 3)       ' Løsning på problem format
BA(3) = (L) Mod &H100               ' &h000000xx
BA(2) = (L \ &H100) Mod &H100       ' &h0000xx00
BA(1) = (L \ &H10000) Mod &H100     ' &h00xx0000
BA(0) = (L \ &H1000000) Mod &H100   ' &hxx000000
PutLong = BA
End Function

Function GetLong(S As String) As Long
Dim BA() As Byte
BA = S
GetLong = BA(3) + (CLng(BA(2)) * &H100&) + (CLng(BA(1)) * &H10000) + (CLng(BA(0)) * &H1000000)
End Function

Function ValidName(F As String, Optional Ext As Variant) As String
Dim vn As String, V As Integer
' Fil navn
For V = Len(F) To 1 Step -1
If Mid(F, V, 1) = "\" Then Exit For
Next
If V <> 0 Then
vn = Mid(F, V + 1)
Else
vn = F
End If
For V = Len(vn) To 1 Step -1
If Mid(vn, V, 1) = "." Then Exit For
Next
If V <> 0 Then
Ext = Mid(vn, V + 1)
Else
Ext = ""
End If
ValidName = vn
End Function
