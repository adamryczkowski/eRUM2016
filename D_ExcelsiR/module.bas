Attribute VB_Name = "Module"
Option Explicit

Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128

Private Type HOSTENT
   hName As Long
   hAliases As Long
   hAddrType As Integer
   hLength As Integer
   hAddrList As Long
End Type

Private Const WS_VERSION_REQD As Long = &H101

Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Private Const IP_SUCCESS As Long = 0
Private Declare Sub apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nBytes As Long)
Private Declare Function apiWSAStartup Lib "wsock32.dll" Alias "WSAStartup" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function apiGetHostByName Lib "wsock32.dll" Alias "gethostbyname" (ByVal hostname As String) As Long


Private Function InitializeSocket() As Boolean
   Dim WSAD As WSADATA
   
   'attempt to initialize the socket
   InitializeSocket = (apiWSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS)
End Function

Private Function IPToText(ByVal IPAddress As String) As String
   'converts characters to numbers
   IPToText = CStr(Asc(IPAddress)) & "." & _
              CStr(Asc(Mid$(IPAddress, 2, 1))) & "." & _
              CStr(Asc(Mid$(IPAddress, 3, 1))) & "." & _
              CStr(Asc(Mid$(IPAddress, 4, 1)))
End Function

Public Function GetIPFromHostName(ByVal sHostname As String) As String
   'converts a host name to an IP address.
   
   Dim nBytes As Long
   Dim ptrHosent As Long
   Dim hstHost As HOSTENT
   Dim ptrName As Long
   Dim ptrAddress As Long
   Dim ptrIPAddress As Long
   Dim sAddress As String 'declare this as Dim sAddress(1) As String if you want 2 ip addresses returned
   
   'try to initalize the socket
   If InitializeSocket() = True Then
      'try to get the IP
      ptrHosent = apiGetHostByName(sHostname & vbNullChar)
      
      If ptrHosent <> 0 Then
         'get the IP address
         apiCopyMemory hstHost, ByVal ptrHosent, LenB(hstHost)
         apiCopyMemory ptrIPAddress, ByVal hstHost.hAddrList, 4
         
         'fill buffer
         sAddress = Space$(4)
         
         'if you want multiple domains returned,
         'fill all items in sAddress array with 4 spaces
         apiCopyMemory ByVal sAddress, ByVal ptrIPAddress, hstHost.hLength
         'change this to
         'CopyMemory ByVal sAddress(0), ByVal ptrIPAddress, hstHost.hLength
         'if you want an array of ip addresses returned
         '(some domains have more than one ip address associated with it)
         
         'get the IP address
         GetIPFromHostName = IPToText(sAddress)
         'if you are using multiple addresses, you need IPToText(sAddress(0)) & "," & IPToText(sAddress(1))
         'etc
      End If
   Else
      MsgBox "Failed to open Socket."
   End If
End Function

Public Sub S_EncodeStringArrayToR(str() As String)
   Dim I As Long
   Dim strO As New gen_StringBuilder
   For I = LBound(str) To UBound(str)
      Dim J As Long, c As String, lng As Long
      strO.Init
      For J = 1 To Len(str(I))
         c = Mid(str(I), J, 1)
         lng = AscW(c)
         If lng > 127 Then
            c = Hex(lng)
            c = Mid("0000", 1, 4 - Len(c)) & c
            strO.AddString "\u"
            strO.AddString c
         ElseIf c = "'" Then
            strO.AddString "\'"
         ElseIf c = "\" Then
            strO.AddString "\\"
         Else
            strO.AddString c
         End If
      Next J

      str(I) = strO.TheString
   Next I
End Sub

Public Function S_EncodeStringIntoR(str As String) As String
   Dim I As Long, c As String, lng As Long
   Dim strO As New gen_StringBuilder
   For I = 1 To Len(str)
      c = Mid(str, I, 1)
      lng = AscW(c)
      If lng > 127 Then
         c = Hex(lng)
         c = Mid("0000", 1, 4 - Len(c)) & c
         strO.AddString "\u"
         strO.AddString c
      ElseIf c = "'" Then
         strO.AddString "\'"
      ElseIf c = "\" Then
         strO.AddString "\\"
      Else
         strO.AddString c
      End If
   Next I
   S_EncodeStringIntoR = strO.TheString
End Function


Public Function S_CountDimensions(Arr As Variant) As Long
      'Sets up the error handler.
      On Error GoTo FinalDimension
      Dim DimNum As Long
 
      'Visual Basic for Applications arrays can have up to 60000
      'dimensions; this allows for that.
      For DimNum = 1 To 60000

         'It is necessary to do something with the LBound to force it
         'to generate an error.
         Dim ErrorCheck  As Long
         ErrorCheck = LBound(Arr, DimNum)

      Next DimNum
      ' The error routine, which is also valid if no error occured (and we had 60000 dimensions!)
FinalDimension:
      S_CountDimensions = DimNum - 1
End Function




