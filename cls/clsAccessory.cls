VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccessory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String

'���Ϲ�����

'��ӡ�����ɹ���
'�������vRs�����б����ֶ�B_ItemID������������ϸ��G_BillDetailAccessory��B_ItemID
'ִ���߼���
'1. �Ƚ�����B_ItemIDд���G_WorkOrder_Accessory��ͬʱ������Ϊ��ӡ����־����
'2. ����ϸ��G_BillDetailAccessory���ֶ�B_Print�ۼ�1
'3. ͨ���洢���̻�ȡ���ܵļ�¼�����Ҵ�ӡ
Public Sub PrintWorkOrder(ByRef vRs As RecordSet)
    If vRs.State <> adStateOpen Then
        Exit Sub
    End If
    
    If vRs.RecordCount <= 0 Then
        Exit Sub
    End If
    
    Dim lItemID As Long
    '1. ���м���������ɹ�����д�����ݡ�ͬʱд���ӡ����
    vRs.MoveFirst
    Do While Not vRs.EOF
        lItemID = vRs!B_ItemID
        strSQL = "exec dbo.usp_Insert_WorkOrder_Accessory '" & Gm.HardWareID.CN & "','" & Gm.HardWareID.CUN & "','" & lItemID & "'"
        Gm.cnnTool.cnn.Execute strSQL
        
        strSQL = "exec dbo.usp_AddPrintNumber_Accessory '" & lItemID & "'"
        Gm.cnnTool.cnn.Execute strSQL
        vRs.MoveNext
    Loop
    
    
    '2. ���ݴ洢���̻�ȡ�������ݲ��Ҵ�ӡ
    
    
    vRs.MoveFirst
End Sub