Attribute VB_Name = "modIP"
Option Explicit

Public Const WSABASEERR = 10000
Public Const WSAEFAULT = (WSABASEERR + 14)
Public Const WSAEINVAL = (WSABASEERR + 22)
Public Const WSAEINPROGRESS = (WSABASEERR + 50)
Public Const WSAENETDOWN = (WSABASEERR + 50)
Public Const WSASYSNOTREADY = (WSABASEERR + 91)
Public Const WSAVERNOTSUPPORTED = (WSABASEERR + 92)
Public Const WSANOTINITIALISED = (WSABASEERR + 93)
Public Const WSAHOST_NOT_FOUND = 11001
Public Const WSADESCRIPTION_LEN = 257
Public Const WSASYS_STATUS_LEN = 129
Public Const WSATRY_AGAIN = 11002
Public Const WSANO_RECOVERY = 11003
Public Const WSANO_DATA = 11004

Public Type WSAData
    wVersion       As Integer
    wHighVersion   As Integer
    szDescription  As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets    As Integer
    iMaxUdpDg      As Integer
    lpVendorInfo   As Long
End Type

Public Type servent
    s_name    As Long
    s_aliases As Long
    S_Port    As Integer
    s_proto   As Long
End Type

Public Type protoent
    p_name    As String 'Official name of the protocol
    p_aliases As Long 'Null-terminated array of alternate names
    p_proto   As Long 'Protocol number, in host byte order
End Type

Public Declare Function WSAStartup _
    Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long

Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long

Public Declare Function gethostbyaddr _
    Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, _
                      ByVal addr_type As Long) As Long

Public Declare Function gethostbyname _
    Lib "ws2_32.dll" (ByVal host_name As String) As Long

Public Declare Function gethostname _
    Lib "ws2_32.dll" (ByVal host_name As String, _
                      ByVal namelen As Long) As Long

Public Declare Function getservbyname _
    Lib "ws2_32.dll" (ByVal serv_name As String, _
                      ByVal proto As String) As Long

Public Declare Function getprotobynumber _
    Lib "ws2_32.dll" (ByVal proto As Long) As Long

Public Declare Function getprotobyname _
    Lib "ws2_32.dll" (ByVal proto_name As String) As Long

Public Declare Function getservbyport _
    Lib "ws2_32.dll" (ByVal Port As Integer, ByVal proto As Long) As Long

Public Declare Function inet_addr _
    Lib "ws2_32.dll" (ByVal cp As String) As Long

Public Declare Function inet_ntoa _
    Lib "ws2_32.dll" (ByVal inn As Long) As Long

Public Declare Function htons _
    Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer

Public Declare Function htonl _
    Lib "ws2_32.dll" (ByVal hostlong As Long) As Long

Public Declare Function ntohl _
    Lib "ws2_32.dll" (ByVal netlong As Long) As Long

Public Declare Function ntohs _
    Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer

Public Declare Sub RtlMoveMemory _
    Lib "kernel32" (hpvDest As Any, _
                    ByVal hpvSource As Long, _
                    ByVal cbCopy As Long)

Public Declare Function lstrcpy _
    Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, _
                                    ByVal lpString2 As Long) As Long

Public Declare Function lstrlen _
    Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

Public Function UnsignedToLong(Value As Double) As Long
    '
    'The function takes a Double containing a value in the 
    'range of an unsigned Long and returns a Long that you 
    'can pass to an API that requires an unsigned Long
    '
    If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
    '
    If Value <= MAXINT_4 Then
        UnsignedToLong = Value
    Else
        UnsignedToLong = Value - OFFSET_4
    End If
    '
End Function

Public Function LongToUnsigned(Value As Long) As Double
    '
    'The function takes an unsigned Long from an API and 
    'converts it to a Double for display or arithmetic purposes
    '
    If Value < 0 Then
        LongToUnsigned = Value + OFFSET_4
    Else
        LongToUnsigned = Value
    End If
    '
End Function



Private Function FetchIPList(hostname As String) As Variant
    Dim I               As Long
    Dim lngPtrToHOSTENT As Long
    Dim udtHostent      As HOSTENT
    Dim lngPtrToIP      As Long
    Dim arrIpAddress()  As Byte
    Dim strIpAddress    As String
    Dim IpList          As Variant
    Dim IpListC         As Long
    '
    '----------------------------------------------------
    '
    'Call the gethostbyname Winsock API function
    'to get pointer to the HOSTENT structure
    lngPtrToHOSTENT = gethostbyname(Trim$(hostname))
    '
    'Check the lngPtrToHOSTENT value
    If lngPtrToHOSTENT = 0 Then
        '
        'If the gethostbyname function has returned 0
        'the function execution is failed. To get
        'error description call the ShowErrorMsg
        'subroutine
        '
        IpList = GetErrorMsg(Err.LastDllError)
        '
    Else
        ReDim IpList(1 To 1)
        IpListC = 0
        '
        'The gethostbyname function has found the address
        '
        'Copy retrieved data to udtHostent structure
        RtlMoveMemory udtHostent, lngPtrToHOSTENT, LenB(udtHostent)
        '
        'Now udtHostent.hAddrList member contains
        'an array of IP addresses
        '
        'Get a pointer to the first address
        RtlMoveMemory lngPtrToIP, udtHostent.hAddrList, 4
        '
        Do Until lngPtrToIP = 0
            '
            'Prepare the array to receive IP address values
            ReDim arrIpAddress(1 To udtHostent.hLength)
            '
            'move IP address values to the array
            RtlMoveMemory arrIpAddress(1), lngPtrToIP, udtHostent.hLength
            '
            'build string with IP address
            For I = 1 To udtHostent.hLength
                strIpAddress = strIpAddress & arrIpAddress(I) & "."
            Next
            strIpAddress = Left$(strIpAddress, Len(strIpAddress) - 1)
            IpListC = IpListC + 1
            ReDim Preserve IpList(1 To IpListC)
            IpList(IpListC) = strIpAddress
            '
            'remove the last dot symbol
            '
            'Clear the buffer
            strIpAddress = ""
            '
            'Get pointer to the next address
            udtHostent.hAddrList = udtHostent.hAddrList + LenB(udtHostent.hAddrList)
            RtlMoveMemory lngPtrToIP, udtHostent.hAddrList, 4
            '
         Loop
        '
    End If
    FetchIPList = IpList
    '
End Function

Function DNS_Startup() As String
    '
    Dim lngRetVal      As Long
    Dim strErrorMsg    As String
    Dim udtWinsockData As WSAData
    Dim lngType        As Long
    Dim lngProtocol    As Long
    '
    'start up winsock service
    lngRetVal = WSAStartup(&H101, udtWinsockData)
    '
    If lngRetVal <> 0 Then
        '
        '
        Select Case lngRetVal
            Case WSASYSNOTREADY
                strErrorMsg = "The underlying network subsystem is not " & _
                    "ready for network communication."
            Case WSAVERNOTSUPPORTED
                strErrorMsg = "The version of Windows Sockets API support " & _
                    "requested is not provided by this particular " & _
                    "Windows Sockets implementation."
            Case WSAEINVAL
                strErrorMsg = "The Windows Sockets version specified by the " & _
                    "application is not supported by this DLL."
        End Select
        '
        DNS_Startup = strErrorMsg
        '
    End If
    '
End Function

Sub DNS_Clean()
    Call WSACleanup
End Sub

Private Function GetErrorMsg(lngError As Long) As String
    Dim strMessage As String
    Select Case lngError
        Case WSANOTINITIALISED
            strMessage = "A successful WSAStartup call must occur " & _
                         "before using this function."
        Case WSAENETDOWN
            strMessage = "The network subsystem has failed."
        Case WSAHOST_NOT_FOUND
            strMessage = "Authoritative answer host not found."
        Case WSATRY_AGAIN
            strMessage = "Nonauthoritative host not found, or server failure."
        Case WSANO_RECOVERY
            strMessage = "A nonrecoverable error occurred."
        Case WSANO_DATA
            strMessage = "Valid name, no data record of requested type."
        Case WSAEINPROGRESS
            strMessage = "A blocking Windows Sockets 1.1 call is in " & _
                         "progress, or the service provider is still " & _
                         "processing a callback function."
        Case WSAEFAULT
            strMessage = "The name parameter is not a valid part of " & _
                         "the user address space."
        Case WSAEINVAL
            strMessage = "A blocking Windows Socket 1.1 call was " & _
                         "canceled through WSACancelBlockingCall."
    End Select
    GetErrorMsg = strMessage
End Function


Function GetIPListStr(ByVal hostname As String) As String
Dim V As Variant
Dim C As Long
DNS_Startup
V = FetchIPList(hostname)
C = 1
If Not IsArray(V) Then 'Error
GetIPListStr = Chr(0) & V & Chr(0)
Else 'Ip array
For C = LBound(V) To UBound(V) - 1
GetIPListStr = GetIPListStr & V(C) & " , "
Next
GetIPListStr = GetIPListStr & V(C)
End If
DNS_Clean
End Function
