VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParallelTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'to jest rekord opisuj¹cy jedno zadanie zdalne

Public m_PID As Long
Public m_cmd As String
Public m_Name As String
Public m_ExecTime As Date
Public m_IsDataCollected As Boolean
Public m_DestVarName As String
Public m_ErrorMsg As String
Private m_ER As RConnection

Friend Sub Init(ER As RConnection)
   Set m_ER = ER
End Sub

Public Property Get IsReady() As Boolean
   If Not m_IsDataCollected Then
      Dim V As Variant
      V = m_ER.Eval(".tmp.parallel.pid." & m_PID & "<-list(pid=" & m_PID & ");setattr(.tmp.parallel.pid." & m_PID & ",'class','process');.tmp.parallel.result.pid." & m_PID & "<-parallel:::mccollect(.tmp.parallel.pid." & m_PID & ", wait=FALSE);is.null(.tmp.parallel.result.pid." & m_PID & ")")
      If V = True Then
         IsReady = False
      ElseIf V = False Then
         IsReady = True
      Else
         Stop
         Exit Property
      End If
      If IsReady Then
         V = m_ER.Eval("if(class(.tmp.parallel.result.pid." & m_PID & ")=='try-error') as.character(.tmp.parallel.result.pid." & m_PID & ") else ''")
         If V <> "" Then
            m_ErrorMsg = V
            m_IsDataCollected = True
            Exit Property
         Else
            If m_DestVarName <> "" Then
               V = m_ER.Exec(m_DestVarName & "<-" & ".tmp.parallel.result.pid." & m_PID & "[[1]]")
               If V <> "OK" Then
                  Stop
               End If
            End If
         End If
         m_IsDataCollected = True
      End If
   Else
      IsReady = True
   End If
End Property

Public Property Get ErrorString() As String
   ErrorString = m_ErrorMsg
End Property

Private Sub Class_Initialize()
   m_ErrorMsg = ""
   m_ExecTime = Now
   m_IsDataCollected = False
End Sub

Public Function CollectData() As Variant 'to be called only when data is ready
   If IsReady Then
      CollectData = m_ER.Eval(".tmp.parallel.result.pid." & m_PID & "[[1]]")
   Else
      CollectData = Null
   End If
End Function

