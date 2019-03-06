VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FunctTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'从表G_Config_ComboSource中获取指定的数据源SQL
Public Function GetSQL(ByVal vGroupName As String) As String
    Dim rs As New RecordSet
    strSQL = "SELECT * FROM G_Config_ComboSource WHERE B_GroupName='" & vGroupName & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        GetSQL = ""
    Else
        GetSQL = IIf(IsNull(rs!B_SQL), "", rs!B_SQL)
    End If
    
    rs.Close
    Set rs = Nothing
End Function
    

'使用下拉框控件数据源工具类cls_Link_Data_Ctl
'对UI上的ComboBox进行数据源的绑定
'oCombo：FORM窗体中的COMBOBOX控件
'vBinder：数据源工具类
'vGroupName：在表G_Config_ComboSource中的B_GroupName
Public Function BindCombo(ByRef oCombo As Object, ByRef vBinder As cls_Link_Data_Ctl, _
    ByVal vGroupName As String) As Boolean
    
    Dim strSQL As String
    Dim rs As New RecordSet
    Dim SourceSQL As String
    Dim sql As String
    Dim fieldDisplay As String
    Dim fieldValue As String
    Dim fieldDefault As String
    Dim fieldTip As String
    
    
    strSQL = "SELECT * FROM G_Config_ComboSource WHERE B_GroupName='" & vGroupName & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount <= 0 Then
        fieldTip = "下拉框控件数据源设置表中没有对应的设置：" & vGroupName
        MsgBox fieldTip, vbOKOnly + vbInformation, "提示"
        BindCombo = False
        Exit Function
    End If
    SourceSQL = rs!B_SQL
    
    
    sql = IIf(IsNull(rs!B_SQL), "", rs!B_SQL)
    fieldDisplay = IIf(IsNull(rs!B_FieldDisplay), "", rs!B_FieldDisplay)
    fieldValue = IIf(IsNull(rs!B_FieldValue), "", rs!B_FieldValue)
    fieldDefault = IIf(IsNull(rs!B_FieldDefault), "", rs!B_FieldDefault)
    fieldTip = IIf(IsNull(rs!B_Tip), "", rs!B_Tip)
    
    Set vBinder = New cls_Link_Data_Ctl
    vBinder.InitCls oCombo, sql, fieldDisplay, fieldValue, fieldDefault, fieldTip
    vBinder.ShowData
    
    rs.Close
    Set rs = Nothing
    
End Function

'判断一个数组是否为空
'返回TRUE表示数组有元素
'反之则没有
Public Function ArrNotEmpty(ByVal vArr As Variant) As Boolean
    ArrNotEmpty = True
    'ArrNotEmpty = g_ArrayTool.ArrNotEmpty(vArr)
End Function



