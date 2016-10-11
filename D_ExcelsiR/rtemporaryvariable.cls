VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RTemporaryVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Var As String
Private m_ER As RConnection

Sub Init(ER As RConnection, strPrefiks As String, Optional N As Long = 3)
   Set m_ER = ER
   m_Var = ER.Eval("gen.unique.obj.name('" & strPrefiks & "'," & N & ")")
End Sub

Public Sub ReleaseVar()
   Set m_ER = Nothing
End Sub

Private Sub Class_Terminate()
   If Not m_ER Is Nothing Then
      m_ER.Exec "rm(" & m_Var & ")"
   End If
End Sub

Public Property Get Name() As String
   Name = m_Var
End Property

