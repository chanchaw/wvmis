VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'写单据打印的日志
Public Sub WriteLog_BillPrint(ByVal vID As Long, ByVal vObjectID As String)
    strSQL = "INSERT INTO Log_PrintBill"
    strSQL = strSQL & " (B_BillID,B_ObjectID,B_UserName,B_Date)"
    strSQL = strSQL & " Values"
    strSQL = strSQL & " ('" & vID & "','" & vObjectID & "','" & Gm.SysID.SystemUser & "','" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
    
    Gm.cnnTool.cnn.Execute strSQL
End Sub
