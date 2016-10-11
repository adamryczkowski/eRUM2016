VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gen_StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
      (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
      

Private m_sString As String
Private m_iPos As Long
Private m_iLen As Long

Implements gen_i_ConstString

Private Sub Class_Initialize()
   Init
End Sub

Public Sub InitStr(ByRef str As String)
   m_sString = str
   m_iPos = Len(str)
   m_iLen = m_iPos
End Sub

Friend Sub InitClone(ByRef str As String, iPos As Long, iLen As Long)
   m_sString = str
   m_iPos = iPos
   m_iLen = iLen
End Sub

Public Sub Init()
   m_sString = ""
   m_iPos = 0
   m_iLen = 0
   'do nothing
End Sub

Friend Sub ExpandString(NotLessThan As Long, Optional bIncludeCurrentLength As Boolean = True) 'nie rusza stringu, jedynie zmienia zakres zarezerwowanej pamiêci
   Dim newStr As String
   Dim lNewLen As Long
   lNewLen = NotLessThan
   If bIncludeCurrentLength Then
      lNewLen = lNewLen + m_iPos
   End If
   If lNewLen * 1.5 < 8192 Then
      lNewLen = 8192 'to jest najmniejsza legalna d³ugoœæ stringu
   Else
      lNewLen = lNewLen * 1.5
   End If
   If m_iLen < lNewLen Then
      m_sString = m_sString & Space(lNewLen - m_iLen)
      m_iLen = lNewLen
   End If
End Sub

Public Function Clone() As gen_StringBuilder
   Set Clone = New gen_StringBuilder
   Clone.InitClone m_sString, m_iPos, m_iLen
End Function

Public Property Get Length() As Long
   Length = m_iPos
End Property

Friend Property Get Capacity() As Long
   Capacity = m_iLen
End Property

Public Property Get TheString() As String
   ' The internal string:
   If m_iPos > 0 Then
      TheString = Left$(m_sString, m_iPos)
   End If
End Property

Public Property Let TheString(ByRef str As String)
   Dim lLen As Long
   
   ' Setting the string:
   lLen = Len(str)
   If lLen = 0 Then
      'Clear
      m_sString = ""
      m_iPos = 0
      m_iLen = 0
   Else
      Clear
      ExpandString lLen
      CopyMemory ByVal StrPtr(m_sString), ByVal StrPtr(str), lLen * 2
      m_iPos = lLen
   End If
End Property

Friend Sub ClearMemory()
   m_sString = ""
   m_iPos = 0
   m_iLen = 0
End Sub

Public Sub Clear()
   m_iPos = 0
End Sub

Friend Sub prvSetNewLength(newLength As Long) 's³u¿y do usuwania czêœci znaków ze stringu, ale mo¿liwe ¿e dodane zostan¹ specje, jeœli nowa d³ugoœæ > aktualnej.
   ExpandString newLength, False
   m_iPos = newLength
End Sub

Public Sub AddString(ByRef str As String)
   Dim lLen As Long
   lLen = Len(str)
   ' Append an item to the string:
   ExpandString Len(str), True
   
   CopyMemory ByVal UnsignedAdd(StrPtr(m_sString), m_iPos * 2), ByVal StrPtr(str), lLen * 2
   m_iPos = m_iPos + lLen
End Sub

Sub AddStrBuilder(Builder As gen_StringBuilder)
   ExpandString Builder.Length, True
   On Error GoTo NieMaPtr
   CopyMemory ByVal UnsignedAdd(StrPtr(m_sString), m_iPos * 2), ByVal Builder.prvPtr, Builder.Length * 2
   On Error GoTo 0
   m_iPos = m_iPos + Builder.Length
   Exit Sub
NieMaPtr:
   AddString Builder.TheString
   On Error GoTo 0
End Sub

Sub AddConstStr(strO As gen_i_ConstString)
   Dim lLen As Long, str As String
   str = strO.TheString
   lLen = Len(str)
   ' Append an item to the string:
   ExpandString Len(str), True
   
   CopyMemory ByVal UnsignedAdd(StrPtr(m_sString), m_iPos * 2), ByVal StrPtr(str), lLen * 2
   m_iPos = m_iPos + lLen
End Sub

Public Sub AddStringByVal(ByVal str As String)
   AddString str
End Sub

Public Sub Insert(ByVal iIndex As Long, ByRef str As String)
   Dim lLen As Long
   Dim lPos As Long
   Dim lSize As Long
'   Stop
   ' is iIndex within bounds?
   If (iIndex > m_iPos Or iIndex < 0) Then
      Err.Raise 9
   Else
      lLen = LenB(str)
      ExpandString lLen / 2, True
      
      ' Move existing characters from current position
      lPos = UnsignedAdd(StrPtr(m_sString), (iIndex) * 2)
      lSize = (m_iPos - iIndex) * 2
      
      ' moving from iIndex to iIndex + lLen
      CopyMemory ByVal UnsignedAdd(lPos, lLen), ByVal lPos, lSize
      
      ' Insert new characters:
      CopyMemory ByVal lPos, ByVal StrPtr(str), lLen
      
      m_iPos = m_iPos + lLen / 2
   End If
End Sub

Public Sub InsertByVal(ByVal iIndex As Long, ByVal str As String)
   Insert iIndex, str
End Sub

Public Sub Remove(ByVal iIndex As Long, ByVal lLen As Long)
'Removes substring
   Dim lSrc As Long
   Dim lDst As Long
   Dim lSize As Long
'   Stop
   ' is iIndex within bounds?
   If (iIndex > m_iPos) And lLen > 0 Then
      Err.Raise 9
   ElseIf lLen < 0 Or iIndex < 0 Then
      Err.Raise 9
   Else
      ' is there sufficient length?
      If ((iIndex + lLen) > m_iPos) Then
         Err.Raise 9
      Else
         ' Need to copy characters from iIndex*2 to m_iPos back by lLen chars:
         lSrc = UnsignedAdd(StrPtr(m_sString), (iIndex + lLen) * 2)
         lDst = UnsignedAdd(StrPtr(m_sString), (iIndex) * 2)
         lSize = (m_iPos - (iIndex + lLen)) * 2
         CopyMemory ByVal lDst, ByVal lSrc, lSize
         m_iPos = m_iPos - lLen
      End If
   End If
End Sub

Public Function Find(ByRef sToFind As String, _
   bReverseSearch As Boolean, _
   Optional ByVal lStartIndex As Long = 1, _
   Optional ByVal compare As VbCompareMethod = vbTextCompare _
   ) As Long
   
   Dim lInstr As Long
   If bReverseSearch Then
      If lStartIndex > m_iPos Then
         lStartIndex = m_iPos
      End If
      lInstr = InStrRev(m_sString, sToFind, lStartIndex, compare)
   Else
      lInstr = InStr(lStartIndex, m_sString, sToFind, compare)
   End If
   If (lInstr <= m_iPos) Then
      Find = lInstr
   End If
End Function

Public Property Get StrMid(lngStart As Long, Optional lngLen As Long = -1) As String
   If lngStart > m_iPos + 1 Then
'      Stop
   End If
   If lngLen = -1 Or lngLen >= m_iPos Then
      StrMid = Mid(m_sString, lngStart, m_iPos - lngStart + 1)
   Else
      If lngLen > 0 Then
         StrMid = Mid(m_sString, lngStart, lngLen)
      End If
   End If
End Property

Public Sub StrMidInsertToB(WhereToInsert As gen_StringBuilder, lngStart As Long, Optional ByVal lngLen As Long = -1)
'highly efficent Mid. Wkleja od lngStart tekst o d³ugoœci lngLen do obiektu WhereToInsert
   If m_iPos - lngStart + 1 < lngLen Or lngLen = -1 Then
      lngLen = m_iPos - lngStart + 1
   End If
   WhereToInsert.ExpandString lngLen, True
   'CopyMemory WhereToInsert.prvPtr, UnsignedAdd(StrPtr(m_sString), (lngStart - 1) * 2), lngLen * 2
   CopyMemory ByVal WhereToInsert.prvPtr + WhereToInsert.Length * 2, ByVal StrPtr(m_sString) + (lngStart - 1) * 2, lngLen * 2
   WhereToInsert.prvSetNewLength WhereToInsert.Length + lngLen
End Sub

Private Function UnsignedAdd(Start As Long, Incr As Long) As Long
' This function is useful when doing pointer arithmetic,
' but note it only works for positive values of Incr

   If Start And &H80000000 Then 'Start < 0
      UnsignedAdd = Start + Incr
   ElseIf (Start Or &H80000000) < -Incr Then
      UnsignedAdd = Start + Incr
   Else
      UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)
   End If
End Function

Friend Property Get prvPtr() As Long
   prvPtr = StrPtr(m_sString)
End Property

Public Sub RemoveAndInsert(lngAbsPos As Long, lngHowManyDelete As Long, strWhatToInsert As String)
   Remove lngAbsPos, lngHowManyDelete
   Insert lngAbsPos - 1, strWhatToInsert
End Sub

Private Function gen_i_ConstString_Clone() As gen_i_ConstString
   Set gen_i_ConstString_Clone = Clone
End Function

Private Function gen_i_ConstString_Find(sToFind As String, bReverseSearch As Boolean, Optional ByVal lStartIndex As Long = 1&, Optional ByVal compare As VbCompareMethod = 1&) As Long
   gen_i_ConstString_Find = Find(sToFind, bReverseSearch, lStartIndex, compare)
End Function

Private Property Get gen_i_ConstString_Length() As Long
   gen_i_ConstString_Length = Length
End Property

Private Property Get gen_i_ConstString_StrMid(lngStart As Long, Optional lngLen As Long = -1&) As String
   gen_i_ConstString_StrMid = StrMid(lngStart, lngLen)
End Property

Private Sub gen_i_ConstString_StrMidInsertToB(WhereToInsert As gen_StringBuilder, lngStart As Long, Optional ByVal lngLen As Long = -1&)
   StrMidInsertToB WhereToInsert, lngStart, lngLen
End Sub

Private Property Get gen_i_ConstString_TheString() As String
   gen_i_ConstString_TheString = TheString
End Property


