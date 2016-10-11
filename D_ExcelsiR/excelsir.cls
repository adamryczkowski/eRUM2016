VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Declarations that link into excelsir.dll
Private Declare Function excelsir_get_conn Lib "Excelsi-R" _
    (ByRef R As RconnCType) As Integer
Private Declare Function excelsir_close_conn Lib "Excelsi-R" _
    (ByRef R As RconnCType) As Integer
Private Declare Function excelsir_eval_c Lib "Excelsi-R" _
    (ByRef R As RconnCType, ByRef cmd As Variant, ByRef A() As Variant, _
     ByRef error_str As Variant, ByRef gict As GraphInfoCType, _
     ByVal return_result As Integer) As Integer
Private Declare Function excelsir_assign_c Lib "Excelsi-R" _
    (ByRef R As RconnCType, ByRef varname As Variant, ByRef A() As Variant, _
     ByRef error_str As Variant, ByVal xl_range As Variant) As Integer
Private Declare Function excelsir_range_to_string Lib "Excelsi-R" _
    (ByRef A() As Variant, ByVal xl_range As Variant, ByRef out_str As Variant, _
     ByRef error_str As Variant) As Integer

Private RConnection As RconnCType
Private GraphInfo As GraphInfoCType
Private m_svConnObj As RTemporaryVariable

Private m_DelayedServer As ParallelExecution

Private m_ErrorMessage As String
Private m_ServerIP As String
Private m_ServerPort As Long
Private m_UsesvSocket As Boolean
Private m_LoadDepWalker As Boolean

Friend Sub InitConnectionParameters(Optional strServerIP As String = "10.0.2.2", Optional lngServerPort As Variant = 6311, Optional bUseSVSocket As Boolean = True, Optional bLoadDependencyWalker As Boolean = False)
   m_ServerIP = strServerIP
   m_ServerPort = lngServerPort
   m_UsesvSocket = bUseSVSocket
   m_LoadDepWalker = bLoadDependencyWalker
   m_ErrorMessage = ""
   Set m_DelayedServer = New ParallelExecution
   m_DelayedServer.Init Me
End Sub

Private Function UsesvServer(Optional svServerAddress As String = "localhost", Optional svServerPort As Long = 8888) As Boolean 'True oznacza, ¿e uda³o siê po³¹czyæ
   Dim v As Variant
   If Not m_svConnObj Is Nothing Then
      Set m_svConnObj = Nothing
   End If
   
   v = Exec("if (!require('R.utils')) install.packages('R.utils')", 60)
   If v <> "OK" Then
      UsesvServer = False
      m_ErrorMessage = v
      Exit Function
   End If
   v = Exec("R.utils::use('devtools')", 60)
   If v <> "OK" Then
      UsesvServer = False
      m_ErrorMessage = v
      Exit Function
   End If
   v = Exec("devtools::install_github('SciViews/svSocket')", 60)
   If v <> "OK" Then
      UsesvServer = False
      m_ErrorMessage = v
      Exit Function
   End If
   v = Exec("library(svSocket)")
   If v <> "OK" Then
      UsesvServer = False
      m_ErrorMessage = v
      Exit Function
   End If
   Dim str As RTemporaryVariable
   Set str = GetUniqueRValue(".tmp.svConnection.")
   If str Is Nothing Then
      m_ErrorMessage = "Error with unique variable generation: " & v
      UsesvServer = False
      Exit Function
   End If
   v = Exec(str.Name & "<-R.utils::evalWithTimeout(socketConnection(host='" & svServerAddress & "', port=" & svServerPort & "),timeout=2)")
   If v <> "OK" Then
      m_ErrorMessage = "Error with unique variable generation: " & v
      UsesvServer = False
      Exit Function
   End If
   Set m_svConnObj = str
   v = Exec("source('/home/Adama-docs/Adam/MyDocs/Statystyka/Maszyny/Dab2/common.R')")
   If v <> "OK" Then
      m_ErrorMessage = "Error with unique variable generation: " & v
      UsesvServer = False
      Exit Function
   End If
   If m_LoadDepWalker Then
      v = Exec("source('/home/Adama-docs/Adam/MyDocs/Statystyka/Maszyny/Dab2/dependency_walker_cacher2.R')")
      If v <> "OK" Then
         m_ErrorMessage = "Error with unique variable generation"
         UsesvServer = False
         Set m_svConnObj = Nothing
         Exit Function
      End If
   End If
   v = Exec("mccollect(); gc()") 'Usuwamy wszystkie nadpoczête zadania w tle i czyœcimy pamiêæ
   If v <> "OK" Then
      'do nothing
      Stop
   End If
   UsesvServer = True
   m_ErrorMessage = ""
End Function

Public Function Connect() As Boolean 'True oznacza, ¿e siê uda³o
   Dim v As Variant
   If Not IsConnected Then
      Set m_svConnObj = Nothing
      Dim result As Integer
      RConnection = defineConn()
      RConnection.hostname = GetIPFromHostName(m_ServerIP)
      RConnection.port = m_ServerPort
      result = excelsir_get_conn(RConnection)
      If result < 0 Then 'Connection was unsuccessful
         m_ErrorMessage = RConnection.msg
         Connect = False
         Exit Function
      End If
   End If
   If IsConnected Then
      v = Exec("source('/home/Adama-docs/Adam/MyDocs/Statystyka/Maszyny/Dab2/common.R')")
      If v <> "OK" Then
         m_ErrorMessage = v
         Connect = False
         Exit Function
      End If
   End If
   If m_UsesvSocket Then
      If IsConnected Then
         Connect = UsesvServer
      End If
   Else
      If m_LoadDepWalker Then
         v = Exec("source('/home/Adama-docs/Adam/MyDocs/Statystyka/Maszyny/Dab2/dependency_walker_cacher2.R')")
         If v <> "OK" Then
            m_ErrorMessage = v
            Connect = False
            Exit Function
         End If
      End If
      Connect = True
   End If
End Function

Sub DisconnectR()
   If Not m_svConnObj Is Nothing Then
      Dim tmp As RTemporaryVariable
      Set tmp = m_svConnObj
      Set m_svConnObj = Nothing
      Exec "close(" & tmp.Name & ")"
   End If
   excelsir_close_conn RConnection
   Set m_DelayedServer = New ParallelExecution
   m_DelayedServer.Init Me
End Sub

Public Property Get FailedAttemptToConnectErrorMessage() As String
   FailedAttemptToConnectErrorMessage = m_ErrorMessage
End Property

Private Function defineConn() As RconnCType
'Initialize an rconn object to default values
   defineConn.Status = -1
   defineConn.sanity_check = 42
   defineConn.msg = "Not yet connected"
   defineConn.hostname = m_ServerIP
   defineConn.port = m_ServerPort
End Function

Public Function ExecDelayedEx(ByRef cmd As String, strName As String, strDestVarName As String) As String
   If Not IsConnected Then
      ExecDelayedEx = "#RError: Rserver is not connected"
   Else
      If m_DelayedServer.TaskExists(strName) Then
         If m_DelayedServer.TaskIsReady(strName) Then
            Dim v As Variant
            v = m_DelayedServer.CollectTaskValue(strName)
            ExecDelayedEx = "Done"
         Else
            ExecDelayedEx = "Running"
         End If
      Else
         Dim T As ParallelTask
         Set T = m_DelayedServer.DelayedEval(cmd, strName, strDestVarName)
         If Not T Is Nothing Then
            ExecDelayedEx = "#RError: Cannot queue the command: " & cmd
         Else
            ExecDelayedEx = "Queued"
         End If
      End If
   End If
End Function

Public Sub RemoveDelayedTask(strName As String)
   If Not m_DelayedServer Is Nothing Then
      m_DelayedServer.RemoveDelayedTask strName
   End If
End Sub

Public Function SetExecDelayed(cmd As String, strName As String, strDestVarName As String) As ParallelTask
   Set SetExecDelayed = m_DelayedServer.DelayedEval(cmd, strName, strDestVarName)
End Function

Public Function ExecDelayed(ByRef cmd As String) As String
   ExecDelayed = ExecDelayedEx(cmd, cmd, "")
End Function

Public Function DelayedTaskExists(strNameCommand As String) As Boolean
   If Not IsConnected Then
      DelayedTaskExists = False
   Else
      DelayedTaskExists = m_DelayedServer.TaskExists(strNameCommand)
   End If
End Function

Public Function EvalDelayed(ByRef cmd As String, destVarName As String) As Variant
   If Not IsConnected Then
      EvalDelayed = "#RError: Rserver is not connected"
   Else
      If m_DelayedServer.TaskExists(cmd) Then
         If m_DelayedServer.TaskIsReady(cmd) Then
            EvalDelayed = m_DelayedServer.CollectTaskValue(cmd)
         Else
            EvalDelayed = "ok; still running."
         End If
      Else
         Dim T As ParallelTask
         Set T = m_DelayedServer.DelayedEval(cmd, cmd, destVarName)
         If Not T Is Nothing Then
            EvalDelayed = "#RError: Cannot queue the command: " & cmd
         Else
            EvalDelayed = "ok; command queued."
         End If
      End If
   End If
End Function

Public Function Eval(cmd As String, Optional Timeout As Variant) As Variant
   Dim A() As Variant
   Dim error_str As Variant
   Dim cmd2 As Variant
   
   Dim result As Integer
   Dim lngTimeout As Long
   lngTimeout = 2
   On Error Resume Next
   lngTimeout = Timeout
   On Error GoTo 0
   If lngTimeout < 0 Then
      lngTimeout = 2
   End If
   
   If RConnection.sanity_check <> 42 Then
      result = excelsir_get_conn(RConnection)
      If result < 0 Then 'Connection was unsuccessful
         Eval = "Not connected to remote Rserve server"
         Exit Function
      End If
   End If
   If cmd = "" Then
      Eval = CVErr(xlErrValue)
      Exit Function
   End If
   
   Dim result_val As Integer
   cmd2 = S_EncodeStringIntoR(cmd)   'Need to pass C function a Variant
   Debug.Print "Eval: " & CStr(cmd)
   If m_svConnObj Is Nothing Then
      cmd2 = "eval(parse(text='" & cmd2 & "'))"
   Else
      cmd2 = "R.utils::evalWithTimeout(evalServer(" & m_svConnObj.Name & ",'" & cmd2 & "'),timeout=" & lngTimeout & ",envir=.GlobalEnv)"
   End If
   result_val = excelsir_eval_c(RConnection, cmd2, A, error_str, GraphInfo, 1)
   If result_val < 0 Then
      Eval = "Error: " & error_str
      Exit Function
   End If
   
   If result_val = 1 Then
      Eval = "Charts not supported"
      Exit Function
   End If
   
   Dim Arr As New gen_StaticArray
   Arr.Init A
   If Arr.IsScalar Then
      Eval = Arr.ValueScalar
   Else
      Eval = A
   End If
End Function

Public Function Exec(cmd As String, Optional Timeout As Variant) As Variant
   Dim A() As Variant
   Dim error_str As Variant
   Dim cmd2 As String
   
   Dim lngTimeout As Long
   lngTimeout = 5
   On Error Resume Next
   lngTimeout = Timeout
   On Error GoTo 0
   If lngTimeout < 0 Then
      lngTimeout = 5
   End If
   
   If RConnection.sanity_check <> 42 Then
      Exec = "Not connected to remote Rserve server"
      Exit Function
   End If
   Dim result_val As Integer
   Dim cmd3 As String, cmd4 As String
   cmd2 = S_EncodeStringIntoR(cmd)
   If m_svConnObj Is Nothing Then
      cmd3 = "eval(parse(text='" & cmd2 & "'))"
   Else
      cmd3 = "R.utils::evalWithTimeout(evalServer(" & m_svConnObj.Name & ",'" & cmd2 & ";1'),timeout=" & lngTimeout & ",envir=.GlobalEnv);1"
   End If
   cmd4 = "tryCatch(capture.output({" & cmd3 & ";1}),error=function(err) conditionMessage(err))"
   Debug.Print "Exec: " & CStr(cmd)
'   ReDim A(0 To 1, 0 To 1)
   result_val = excelsir_eval_c(RConnection, CVar(cmd4), A, error_str, GraphInfo, 1)
   If result_val < 0 Then
      Exec = "Error: " & error_str
   Else
      If A(0, 0) <> "[1] 1" Then
         Exec = A(0, 0)
      Else
         Exec = "OK"
      End If
   End If
End Function

Public Function RPut(varname As String, Value As Variant) As String
   'Assign the data contained in xl_range to the variable varname in R
   Dim Arr As New gen_StaticArray
   Dim A() As Variant
   Dim error_str As Variant
   Dim varname2 As Variant
   Dim result_val As Integer
   
'      ReDim a(LBound(Value, 1) To UBound(Value, 1))
'      Dim I As Long
'      For I = LBound(a) To UBound(a)
'         a(I) = Value(I, 1)
'      Next I
'   End If
   Const lngTimeout = 5
   If RConnection.sanity_check <> 42 Then
      RPut = "Not connected to remote Rserve server"
   ElseIf varname = "" Then
      RPut = "Not given proper name for R variable"
   Else
      varname2 = varname
      If Not IsArray(Value) Then
         ReDim A(1 To 1, 1 To 1)
         A(1, 1) = Value
      Else
         If TypeOf Value Is Excel.Range Then
 '           If m_svConnObj Is Nothing Then
               Arr.Init S_EncodeStringIntoR(Value.Value2)
 '           Else
 '              arr.Init value.Value2
 '           End If
         Else
            Dim Varr() As String
            Varr = Value
'            If Not m_svConnObj Is Nothing Then
               S_EncodeStringArrayToR Varr
'            End If
            Arr.Init Varr
         End If
         A = Arr.ArrayAsExcelArray
      End If
      result_val = excelsir_assign_c(RConnection, varname2, A, error_str, Null)
      If result_val = 0 Then
         Dim cmd As String, v As String
         If m_svConnObj Is Nothing Then
            cmd = varname2 & "<-stri_unescape_unicode(" & varname2 & ")"
            v = Exec(cmd)
            If v <> "OK" Then
               Stop
            End If
            RPut = "OK"
         Else
            Dim tmpCon As RTemporaryVariable
            Set tmpCon = m_svConnObj
            Set m_svConnObj = Nothing
            
            cmd = "evalServer(" & tmpCon.Name & ", " & varname2 & ", stringi::stri_unescape_unicode(" & varname2 & "));rm(" & varname2 & ")"
'            cmd = "evalServer(" & tmpCon.Name & ", stri_unescape_unicode(" & varname2 & "));rm(" & varname2 & ")"
            v = Exec(cmd, lngTimeout)
            If v <> "OK" Then
               Stop
            End If
            Set m_svConnObj = tmpCon
            RPut = "OK"
         End If
         
      Else
         RPut = "Error: " & error_str
      End If
   End If
End Function

Public Property Get ConnectionStatus() As eRConnectionStatus
   'Return whether or not we are connected
   Dim IsConnected  As Boolean
   IsConnected = RConnection.sanity_check = 42 And RConnection.Status <> -1
   If Not IsConnected Then
      ConnectionStatus = eRConnectionStatus_None
      Exit Property
   End If
   Dim A() As Variant
   Dim error_str As Variant
   Dim result_val As Integer
   result_val = excelsir_eval_c(RConnection, "2+2", A, error_str, GraphInfo, 1)
   If result_val <> 0 Then
      ConnectionStatus = eRConnectionStatus_Error_NoRServe
      Exit Property
   End If
   ConnectionStatus = eRConnectionStatus_RServe
   If m_svConnObj Is Nothing Then
      If A(0, 0) <> 4 Then
         ConnectionStatus = eRConnectionStatus_Error_NoRServe
      End If
      Exit Property
   End If
   ConnectionStatus = eRConnectionStatus_svSocket
   Dim v As Variant
   Dim strErr As String
   If m_svConnObj Is Nothing Then
      strErr = UsesvServer()
      If strErr <> "OK" Then
         ConnectionStatus = eRConnectionStatus_Error_NosvSocket
      Else
         ConnectionStatus = eRConnectionStatus_svSocket
      End If
   Else
      v = Eval("2+2")
      If v <> "4" Then
         ConnectionStatus = eRConnectionStatus_Error_NosvSocket
      End If
   End If
End Property

Public Function IsConnected() As Boolean
   Select Case ConnectionStatus
   Case eRConnectionStatus_RServe, eRConnectionStatus_svSocket
      IsConnected = True
   Case Else
      IsConnected = False
   End Select
End Function

Function GetRVersion() As String
   If IsConnected Then
      GetRVersion = Eval("paste0(R.Version()$major,'.',R.Version()$minor)")
   End If
End Function

Private Sub Class_Terminate()
   excelsir_close_conn RConnection
End Sub

Public Function GetUniqueRValue(strPrefix As String, Optional N As Long = 3) As RTemporaryVariable
   Set GetUniqueRValue = New RTemporaryVariable
   GetUniqueRValue.Init Me, strPrefix, N
   If Mid(GetUniqueRValue.Name, 1, 7) = "Error: " Then
      Set GetUniqueRValue = Nothing
   End If
End Function

Public Property Get ServerIP() As String
   ServerIP = m_ServerIP
End Property

Public Property Get ServerPort() As Long
   ServerPort = m_ServerPort
End Property

Public Property Get UseSVSocket() As Boolean
   UseSVSocket = m_UsesvSocket
End Property

Public Sub ReInit(Optional ByVal strServerIP As String = "", Optional ByVal lngServerPort As Variant = -1, Optional vbUseSVSocket As VbTriState = vbUseDefault, Optional vbLoadDependencyWalker As VbTriState = vbUseDefault)
   If strServerIP = "" Then
      strServerIP = m_ServerIP
   End If
   If lngServerPort = -1 Then
      lngServerPort = m_ServerPort
   End If
   Dim bUseSVSocket As Boolean
   Select Case vbUseSVSocket
   Case vbFalse
      bUseSVSocket = False
   Case vbTrue
      bUseSVSocket = True
   Case vbUseDefault
      bUseSVSocket = m_UsesvSocket
   End Select
   
   Dim bLoadDependencyWalker As Boolean
   Select Case vbLoadDependencyWalker
   Case vbFalse
      bLoadDependencyWalker = False
   Case vbTrue
      bLoadDependencyWalker = True
   Case vbUseDefault
      bLoadDependencyWalker = m_LoadDepWalker
   End Select
   InitConnectionParameters strServerIP, lngServerPort, bUseSVSocket, bLoadDependencyWalker
End Sub
