VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusiness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String

'本项目的带有业务逻辑的工具类

'根据订单主表的B_ID获取对应的辅料计划单的主表B_ID
Public Function GetAccePLNID(ByVal vIDOrder As Long) As Long
    On Error Resume Next
    Dim rs As New RecordSet
    strSQL = "exec dbo.usp_Rpt_GetOrders_Acce '" & vIDOrder & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        GetAccePLNID = IIf(IsNull(rs!B_ID), 0, rs!B_ID)
    Else
        GetAccePLNID = 0
    End If
    
    rs.Close
    Set rs = Nothing
End Function
