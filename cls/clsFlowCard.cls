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

'�����������ڹ���ɫ����ת��
Private strSQL As String

'�ж�һ�������Ƿ����
'�����򷵻�TRUE�����򷵻�FALSE
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

'�ж϶���������Ƿ���ڶ���ͻ�
'���ڶ���ͻ��򷵻�FALSE�����򷵻�TRUE
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

'�������ܣ�
'vCNs�������ַ������Թ̶�������Ϊ���
'vSymbol���������
'����ֵ�����ؿ����ַ�����ѯ��ȡ���ĺ�ͬ�еĿͻ���š��ͻ����ƣ�DISTINCT���ģ�
Public Function GetClient4CNs(ByVal vCNs As String, _
    ByVal vSymbol As String) As RecordSet
    
    Dim rs As New RecordSet
    strSQL = "exec dbo.usp_GetCNClient '" & vCNs & "','" & vSymbol & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Set GetClient4CNs = rs.Clone
    
    rs.Close
    Set rs = Nothing
End Function
