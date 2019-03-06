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

'辅料工具类

'打印辅料派工单
'传入参数vRs，其中必有字段B_ItemID，代表辅料明细表G_BillDetailAccessory的B_ItemID
'执行逻辑：
'1. 先将所有B_ItemID写入表G_WorkOrder_Accessory（同时可以作为打印的日志表）
'2. 向明细表G_BillDetailAccessory的字段B_Print累加1
'3. 通过存储过程获取汇总的记录集并且打印
Public Sub PrintWorkOrder(ByRef vRs As RecordSet)
    If vRs.State <> adStateOpen Then
        Exit Sub
    End If
    
    If vRs.RecordCount <= 0 Then
        Exit Sub
    End If
    
    Dim lItemID As Long
    '1. 向中间表（辅料派工单）写入数据。同时写入打印次数
    vRs.MoveFirst
    Do While Not vRs.EOF
        lItemID = vRs!B_ItemID
        strSQL = "exec dbo.usp_Insert_WorkOrder_Accessory '" & Gm.HardWareID.CN & "','" & Gm.HardWareID.CUN & "','" & lItemID & "'"
        Gm.cnnTool.cnn.Execute strSQL
        
        strSQL = "exec dbo.usp_AddPrintNumber_Accessory '" & lItemID & "'"
        Gm.cnnTool.cnn.Execute strSQL
        vRs.MoveNext
    Loop
    
    
    '2. 根据存储过程获取汇总数据并且打印
    
    
    vRs.MoveFirst
End Sub
