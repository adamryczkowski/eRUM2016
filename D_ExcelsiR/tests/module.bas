Attribute VB_Name = "Module"
Option Explicit

Sub Main()
   Dim X As RConnection
   
   Set X = CreateRConnection
   ProbaZero X
   Proba1 X
   Proba2 X
   Proba3 X
   Proba4 X
   Proba5 X
   Proba6 X
   Proba7 X
End Sub

Sub ProbaZero(X As RConnection)
   If X.GetRVersion = "" Then
      Stop
   End If
End Sub

Sub Proba1(X As RConnection)
   Dim V As Variant
   V = X.Eval("2+2")
   If V <> 4 Then
      Stop
   End If
End Sub

Sub Proba2(X As RConnection)
   Dim V As Variant
   V = X.Eval("1:10")
   If IsArray(V) Then
      If V(3, 0) <> 4 Then
         Stop
      End If
   Else
      Stop
   End If
End Sub

Sub Proba3(X As RConnection)
   Dim V As Variant
   V = X.Eval("""Ala Ma kota""")
   If V <> "Ala Ma kota" Then
      Stop
   End If
End Sub

Sub Proba4(X As RConnection)
   X.Exec "source('/home/Adama-docs/Adam/MyDocs/Statystyka/Maszyny/Dab2/path.cat.R')"
   Dim V As Variant
   V = X.Eval("path.cat('/home/vika','../adam','Documents/Statystyka','Aktywne analizy')")
   If V <> "/home/adam/Documents/Statystyka/Aktywne analizy" Then
      Stop
   End If
End Sub

Sub Proba5(X As RConnection)
   X.Exec "makeSureInstalled(""data.table"")"
   Dim V As Variant
   V = X.Eval("exists('fread')")
   If V = False Then
      Stop
   End If
End Sub

Sub Proba7(X As RConnection)
   Dim V As Variant
   X.Exec "short<-readRDS('/home/Adama-docs/Adam/MyDocs/Statystyka/Aktywne analizy/ATolw/01 7XII2015/short.rds');1"
   Stop
End Sub

Sub Proba6(X As RConnection)
   X.Exec "x<-13"
   Dim V As Variant
   V = X.Eval("x")
   If V <> 13 Then
      Stop
   End If
End Sub
