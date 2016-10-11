VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExcelsiR_Singleton"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eRConnectionStatus
   eRConnectionStatus_None
   eRConnectionStatus_RServe
   eRConnectionStatus_svSocket
   eRConnectionStatus_Error_NoRServe
   eRConnectionStatus_Error_NosvSocket
End Enum

Public Type RconnCType
    'This holds a pointer to the Rconnection object in C
    rc As Variant 'holds actual pointer
    sanity_check As Integer
    Status As Integer 'this depends on Rconnection connect result
    msg As Variant
    hostname As Variant
    port As Integer
End Type

Public Type GraphInfoCType
    'Pass graph information back and forth to C++
    Name As Variant
    FileName As Variant
End Type


Public Function EncodeStringIntoR(str As String) As String
   Dim I As Long, c As String, lng As Long
   EncodeStringIntoR = ""
   For I = 1 To Len(str)
      c = Mid(str, I, 1)
      lng = AscW(c)
      If lng > 127 Then
         c = Hex(lng)
         c = Mid("0000", 1, 4 - Len(c)) & c
         EncodeStringIntoR = EncodeStringIntoR & "\u" & c
      ElseIf c = "'" Then
         EncodeStringIntoR = EncodeStringIntoR & "\'"
      ElseIf c = "\" Then
         EncodeStringIntoR = EncodeStringIntoR & "\\"
      Else
         EncodeStringIntoR = EncodeStringIntoR & c
      End If
   Next I
End Function


Public Function CreateRConnection(Optional strServerIP As String = "10.0.2.2", Optional lngServerPort As Variant = 6311, Optional bUseSVSocket As Boolean = True, Optional bLoadDependencyWalker As Boolean = False) As RConnection
   Set CreateRConnection = New RConnection
   CreateRConnection.InitConnectionParameters strServerIP, lngServerPort, bUseSVSocket, bLoadDependencyWalker
End Function

