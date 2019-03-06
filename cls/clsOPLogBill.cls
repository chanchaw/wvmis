VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOPLogBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String

'设置作废的日志写入
Public Sub WriteLog_SetVoid(ByVal vID As Long, ByVal vCodeID As String, ByVal vObjectID As String)
    Dim szDate As String
    szDate = Format(Now, "YYYY-MM-DD HH:MM:SS")
    strSQL = "exec dbo.P_OPLog_SetVoid '" & Gm.SysID.ComputerName & "','" & Gm.SysID.ComputerUserName & "','" & Gm.SysID.iP & "','" & Gm.SysID.SubSystem & "','" & Gm.SysID.SystemUser & "','" & szDate & "','" & vID & "','" & vCodeID & "','" & vObjectID & "'"
    Gm.cnnTool.cnn.Execute strSQL
    
'    strSQL = "exec dbo.P_OPLog_SetVoid '" & E_ComputerName & "','" & E_ComputerUserName & "','" & E_IP & "','" & sSystemName & "','" & UserName & "','" & szDate & "','" & vID & "','" & vCodeID & "','" & vObjectID & "'"
'    cnn.cnn.Execute strSQL
    
End Sub


