VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gen_StaticArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'to jest szybka klasa u³atwiaj¹ca dostêp do macierzy wyników (Array). Obiekt typu Array jest kopiowany tylko raz - podczas konstruktora.
'Indeksowanie macierzy zawsze jest liczone od jedynki(!), niezale¿nie od indeksowania obiektu m_Arr

'Jest to bardzo wygodny sposób na przechowywanie wyniku obliczeñ jakiegoœ zakresu z Excela, gdy¿ taki wynik obliczeñ
'zawsze jest zwracany jako dwuwymiarowa macierz, nawet jeœli to jest jedna kolumna lub jeden wiersz.

'Trójwymiarowe tablice s¹ nie do koñca supportowane - tj. nie dzia³a na nich Drop dimensions; macierz formalnie zadana jako 3-wymiarowa jest dostêpna tylko poprzez interfejs 3-wymiarowy

Private m_Arr As Variant
Private m_VisibleDimensions As Long 'ile jest z praktycznego punktu widzenia wymiarów (mo¿e byæ 0, 1 lub 2 lub nawet 3)
Private m_VBDimensions As Long 'Wynik dzia³ania S_CountDimensions(Arr)
Private m_Transpose As Boolean 'True, jeœli chcemy transpozycjê. Chcemy wtedy, gdy wektor nie jest kolumnowy, lecz wierszowy. Transpozycja wymaga dwuwymiarowej tabeli

Function Clone() As gen_StaticArray
   Set Clone = New gen_StaticArray
   Clone.Init m_Arr
End Function

Sub InitManual2D(lngIleKolumn As Long, IleWierszy As Long)
   Dim Arr() As Variant
   ReDim Arr(1 To lngIleKolumn, 1 To IleWierszy) As Variant
   m_VisibleDimensions = 2
   m_VBDimensions = 2
   m_Arr = Arr
End Sub

Sub Init(ByRef Arr As Variant)
   m_VBDimensions = S_CountDimensions(Arr)
   Select Case m_VBDimensions
   Case 3
      m_Arr = Arr 'trójwymiarowe tablice nie s¹ tak do koñca supportowane...
   Case 2
      If UBound(Arr, 1) - LBound(Arr, 1) = 0 Then
         If UBound(Arr, 2) - LBound(Arr, 2) = 0 Then
            m_Arr = Arr
            m_VisibleDimensions = 0
         Else
            m_Arr = Arr
            m_VisibleDimensions = 1
            m_Transpose = True
         End If
      ElseIf UBound(Arr, 2) - LBound(Arr, 2) = 0 Then
         m_Arr = Arr
         m_VisibleDimensions = 1
         m_Transpose = False
      Else
         m_Arr = Arr
         m_VisibleDimensions = 2
         'do nothing
      End If
   Case 0
      ReDim m_Arr(0, 0)
      m_Arr(0, 0) = Arr
   Case 1
      'nie pozwalamy póki co na wektory jednowymiarowe (dla szybkoœci dzia³ania)
      If UBound(Arr) - LBound(Arr) = 0 Then
         m_VisibleDimensions = 0
         m_Arr = Arr
      Else
         m_VisibleDimensions = 1
         m_Arr = Arr
      End If
   Case Else
      Stop 'nie supportujemy trój- i wiêcej wymiarowych tablic.
   End Select
End Sub

Public Sub ForceTranspose2DimArray() 'transponuje macierz, jeœli jest to rzeczywiœcie dwu-wymiarowa macierz
   If m_VisibleDimensions = 2 Then
      m_Transpose = m_Transpose Xor m_Transpose
   Else
      MsgBox "Transponowaæ mo¿na tylko dwuwymiarowe macierze."
   End If
End Sub

Public Property Get ValuenD(Wsp() As Long) As Variant
   Select Case UBound(Wsp)
   Case 1
      ValuenD = Value1D(Wsp(1))
   Case 2
      ValuenD = Value2D(Wsp(1), Wsp(2))
   Case 3
      ValuenD = Value3D(Wsp(1), Wsp(2), Wsp(3))
   Case Else
      Stop 'to be implemented
   End Select
End Property

Public Property Get Value2D(Col As Long, Row As Long) As Variant
   If m_VBDimensions <> 2 Then
      Stop
   End If
   If m_Transpose Then
      Value2D = m_Arr(Row - 1 + LBound(m_Arr, 1), Col - 1 + LBound(m_Arr, 2))
   Else
      Value2D = m_Arr(Col - 1 + LBound(m_Arr, 2), Row - 1 + LBound(m_Arr, 1))
   End If
End Property

Public Property Get Value3D(X As Long, Y As Long, z As Long) As Variant
   Value3D = m_Arr(X - 1 + LBound(m_Arr, 1), Y - 1 + LBound(m_Arr, 2), z - 1 + LBound(m_Arr, 3))
End Property

Public Property Get Value1D(Idx As Long) As Variant
   If m_VBDimensions = 2 Then
      If m_Transpose Then
         Value1D = m_Arr(LBound(m_Arr, 1), Idx - 1 + LBound(m_Arr, 2))
      Else
         Value1D = m_Arr(Idx - 1 + LBound(m_Arr, 1), LBound(m_Arr, 2))
      End If
   ElseIf m_VBDimensions = 1 Then
      Value1D = m_Arr(Idx - 1 + LBound(m_Arr))
   ElseIf m_VBDimensions = 0 Then
      If Idx = 1 Then
         Value1D = m_Arr(0, 0)
      Else
         MsgBox "Index too large for 0-dimensional array"
         Stop
      End If
   Else
      Stop
   End If
End Property

Public Property Get UpperBound(DimensionNr As Long)
   UpperBound = UBound(m_Arr, DimensionNr)
End Property

Public Property Get LowerBound(DimensionNr As Long)
   LowerBound = LBound(m_Arr, DimensionNr)
End Property

Public Property Let Value2D(Col As Long, Row As Long, NewValue As Variant)
   If m_VBDimensions = 2 Then
      If m_Transpose Then
         m_Arr(Row - 1 + LBound(m_Arr, 1), Col - 1 + LBound(m_Arr, 2)) = NewValue
      Else
         m_Arr(Col - 1 + LBound(m_Arr, 2), Row - 1 + LBound(m_Arr, 1)) = NewValue
      End If
   Else
      Stop
   End If
End Property

Public Property Let Value1D(Idx As Long, NewValue As Variant)
   If m_VBDimensions = 2 Then
      If m_Transpose Then
         m_Arr(LBound(m_Arr, 1), Idx - 1 + LBound(m_Arr, 2)) = NewValue
      Else
         m_Arr(Idx - 1 + LBound(m_Arr, 1), LBound(m_Arr, 2)) = NewValue
      End If
   Else
      m_Arr(Idx - 1 - LBound(m_Arr)) = NewValue
   End If
End Property

Public Property Get ValueScalar() As Variant
   If m_VBDimensions = 2 Then
      ValueScalar = m_Arr(LBound(m_Arr, 1), LBound(m_Arr, 2))
   ElseIf m_VBDimensions = 1 Then
      ValueScalar = m_Arr(LBound(m_Arr))
   Else
      ValueScalar = m_Arr(0, 0)
   End If
End Property

Public Property Get DimensionCount() As Long
   DimensionCount = m_VisibleDimensions
End Property

Public Property Get RawDimensionCount() As Long
   RawDimensionCount = S_CountDimensions(m_Arr)
End Property

Public Property Get IsScalar() As Boolean 'zwraca True, jeœli w³o¿ona macierz jest jednoelementowa
   IsScalar = m_VisibleDimensions = 0
End Property

Public Property Get IsVector() As Boolean 'zwraca True, jeœli w³o¿ona macierz jest jednoelementowa
   IsVector = m_VisibleDimensions = 1
End Property

Public Property Get IsMatrix() As Boolean 'zwraca True, jeœli w³o¿ona macierz jest jednoelementowa
   IsMatrix = m_VisibleDimensions = 2
End Property

Public Property Get RowLength() As Long
   If m_VBDimensions = 2 Then
      If m_Transpose Then
         RowLength = DimensionLength(2)
      Else
         RowLength = DimensionLength(1)
      End If
   ElseIf m_VBDimensions = 1 Then
      RowLength = DimensionLength(1)
   Else
      RowLength = 1
   End If
End Property

Public Property Get ColumnLength() As Long
   If m_VBDimensions = 2 Then
      If m_Transpose Then
         ColumnLength = DimensionLength(2)
      Else
         ColumnLength = DimensionLength(1)
      End If
   ElseIf m_VBDimensions = 1 Then
      ColumnLength = DimensionLength(1)
   Else
      ColumnLength = 1
   End If
End Property

Public Property Get VectorLength() As Long
   If m_VisibleDimensions = 1 Then
      VectorLength = RowLength
   ElseIf m_VisibleDimensions = 0 Then
      If IsNull(m_Arr(0, 0)) Then
         VectorLength = 0
      Else
         VectorLength = 1
      End If
   Else
      MsgBox "To nie jest wektor!"
      Stop
   End If
End Property

Public Property Get DimensionLength(DimensionNr As Long) As Long
   DimensionLength = UBound(m_Arr, DimensionNr) - LBound(m_Arr, DimensionNr) + 1
End Property

Public Property Get Iterator() As gen_ArrayIterator
   Set Iterator = New gen_ArrayIterator
   Iterator.Init Me
End Property

Function ArrayAsExcelArray(Optional bTreat1DimArrAsRow As Boolean = False) As Variant
'Jeœli wybrano "bTreat1DimArrAsRow = True" - wtedy dokonuje transpozycji
   Dim I As Long, J As Long, IMax As Long, JMax As Long, v() As Variant
   Select Case m_VisibleDimensions
   Case 0
      ArrayAsExcelArray = ValueScalar
      Exit Function
   Case 1
      IMax = RowLength
      If bTreat1DimArrAsRow Then
         ReDim v(1 To 1, 1 To IMax)
         For I = 1 To IMax
            v(1, I) = Value1D(I)
         Next I
      Else
         ReDim v(1 To IMax, 1 To 1)
         For I = 1 To IMax
            v(I, 1) = Value1D(I)
         Next I
      End If
   Case 2
      If m_Transpose Xor bTreat1DimArrAsRow Then
         IMax = RowLength
         JMax = ColumnLength
         ReDim v(1 To JMax, 1 To IMax)
         For J = 1 To JMax
            For I = 1 To IMax
               v(J, I) = Value2D(I, J)
            Next I
         Next J
      Else
         IMax = RowLength
         JMax = ColumnLength
         ReDim v(1 To IMax, 1 To JMax)
         For J = 1 To JMax
            For I = 1 To IMax
               v(I, J) = Value2D(I, J)
            Next I
         Next J
      End If
   Case Else
      Stop
   End Select
   ArrayAsExcelArray = v
End Function

