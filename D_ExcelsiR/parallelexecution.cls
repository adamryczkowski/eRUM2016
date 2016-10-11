VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParallelExecution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Tasks As Scripting.Dictionary 'Name as String -> ParallelTask
Private m_ER As RConnection

Sub Init(ER As RConnection)
   Set m_ER = ER
End Sub

Function DelayedEval(ByRef cmd As String, Name As String, Optional destVarName As String = "") As ParallelTask  'True oznacza b³¹d
   If m_Tasks.Exists(Name) Then
      Set DelayedEval = Nothing
   End If
   Dim T As New ParallelTask
   T.Init m_ER
   T.m_Name = Name
   T.m_cmd = cmd
   T.m_DestVarName = destVarName
   Dim V As Variant
   V = m_ER.Eval(".tmp.parallel.exec<-mcparallel({" & cmd & "})$pid")
   On Error GoTo Blad
   T.m_PID = CLng(V)
   m_Tasks.Add Name, T
   Set DelayedEval = T
   Exit Function
Blad:
   On Error GoTo 0
   cmd = V
   Set DelayedEval = Nothing
End Function

Public Property Get TaskIsRunning(strName As String) As Boolean
   If m_Tasks.Exists(strName) Then
      Dim T As ParallelTask
      Set T = m_Tasks.Item(strName)
      TaskIsRunning = Not T.IsReady
   Else
      TaskIsRunning = False
   End If
End Property

Public Property Get TaskIsReady(strName As String) As Boolean
   If m_Tasks.Exists(strName) Then
      Dim T As ParallelTask
      Set T = m_Tasks.Item(strName)
      TaskIsReady = T.IsReady
   Else
      TaskIsReady = False
   End If
End Property

Public Property Get TaskExists(strName As String) As Boolean
   TaskExists = m_Tasks.Exists(strName)
End Property

Public Function CollectTaskValue(strName As String, Optional bRemoveTask As Boolean) As Variant
   If TaskExists(strName) Then
      Dim T As ParallelTask
      Set T = m_Tasks.Item(strName)
      If T.IsReady Then
         CollectTaskValue = T.CollectData
      Else
         CollectTaskValue = "Task " & strName & " is still running"
      End If
   Else
      CollectTaskValue = "Unknown task " & strName
   End If
End Function

Public Sub RemoveDelayedTask(strName As String)
   If TaskExists(strName) Then
      m_Tasks.Remove strName
   End If
End Sub

Private Sub Class_Initialize()
   Set m_Tasks = New Scripting.Dictionary
   m_Tasks.CompareMode = BinaryCompare
End Sub

