VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gen_i_ConstString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Clone() As gen_i_ConstString

End Function

Public Function Find(sToFind As String, bReverseSearch As Boolean, Optional ByVal lStartIndex As Long = 1&, Optional ByVal compare As VbCompareMethod = 1&) As Long

End Function

Public Property Get Length() As Long

End Property

Public Property Get StrMid(lngStart As Long, Optional lngLen As Long = -1&) As String

End Property

Public Sub StrMidInsertToB(WhereToInsert As gen_StringBuilder, lngStart As Long, Optional ByVal lngLen As Long = -1&)

End Sub

Public Property Get TheString() As String

End Property


