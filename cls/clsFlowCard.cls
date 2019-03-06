VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFlowCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'本工具类用于管理色布流转卡
Private strSQL As String

'判断一个卡号是否存在
'存在则返回TRUE，否则返回FALSE
Public Function CheckBCExists(ByVal vCardNumber As String) As Boolean
    Dim rs As New RecordSet
    strSQL = "SELECT * FROM G_BillDetailColor WHERE B_BCIncr='" & vCardNumber & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        CheckBCExists = False
    Else
        CheckBCExists = True
    End If
    rs.Close
    Set rs = Nothing
End Function

'判断多个卡号中是否存在多个客户
'存在多个客户则返回FALSE，否则返回TRUE
Public Function CheckClientUnique(ByVal vCNs As String, _
    ByVal vSymbol As String) As Boolean
    
    Dim rs As New RecordSet
    strSQL = "exec dbo.usp_GetCNClient '" & vCNs & "','" & vSymbol & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount > 1 Then
        CheckClientUnique = False
    Else
        CheckClientUnique = True
    End If
    rs.Close
    Set rs = Nothing
End Function

'参数介绍：
'vCNs：卡号字符串，以固定符号作为间隔
'vSymbol：间隔符号
'返回值：返回卡号字符串查询获取到的合同中的客户编号、客户名称（DISTINCT过的）
Public Function GetClient4CNs(ByVal vCNs As String, _
    ByVal vSymbol As String) As RecordSet
    
    Dim rs As New RecordSet
    strSQL = "exec dbo.usp_GetCNClient '" & vCNs & "','" & vSymbol & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Set GetClient4CNs = rs.Clone
    
    rs.Close
    Set rs = Nothing
End Function

