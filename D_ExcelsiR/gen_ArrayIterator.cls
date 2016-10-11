VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gen_ArrayIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'to jest klasa, która iteruje po polach arbitralnej StaticArray

Private m_Arr As gen_StaticArray
Private m_Pos() As Long 'indeks pozycji w tablicy. Zaczynamy od najm³odszych i idziemy do najstarszych

Sub Init(Arr As gen_StaticArray)
   Set m_Arr = Arr
   ReDim m_Pos(1 To Arr.RawDimensionCount)
   Reset
End Sub

Public Sub NextIteration()
   m_Pos(1) = m_Pos(1) + 1
   Dim I As Long
   For I = 1 To UBound(m_Pos)
      If m_Pos(I) > m_Arr.UpperBound(I) Then
         If I < UBound(m_Pos) Then
            m_Pos(I) = m_Arr.LowerBound(I)
            m_Pos(I + 1) = m_Pos(I + 1) + 1
         End If
      Else
         Exit Sub
      End If
   Next I
End Sub

Public Sub Reset() 'zaczynamy iteracje od pocz¹tku
   Dim I As Long
   For I = 1 To m_Arr.RawDimensionCount
      m_Pos(I) = m_Arr.LowerBound(I)
   Next I
End Sub

Public Property Get EOI() As Boolean
   EOI = m_Pos(UBound(m_Pos)) > m_Arr.UpperBound(UBound(m_Pos))
End Property

Public Property Get Value() As Variant
   Value = m_Arr.ValuenD(m_Pos)
End Property

Public Property Get Wsp(DimensionNr As Long) As Long
   Wsp = m_Pos(DimensionNr)
End Property

Public Property Get X() As Long 'Pierwsza wspó³rzêdna
   X = Wsp(1)
End Property

Public Property Get Y() As Long 'Pierwsza wspó³rzêdna
   Y = Wsp(2)
End Property

