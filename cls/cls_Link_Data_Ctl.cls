VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Link_Data_Ctl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'本类作用与使用方法：
'将一个记录集绑定到ComboBox中，将显示数据和Key数据分离。
'界面显示的Caption，本类会获取对应的Key字段的VALUE来获取。

Private oCombo As Object '注意该ComboBox控件不可是控件数组。否则使用oCombo.clear非法
Private A_rs As Recordset
Private strSQL As String

Private A_FieldNameDisplay As String
Private A_FieldNameGet As String
Private A_FieldNameDefault As String   '默认显示的字段名称

Private A_bIsInit As Boolean
Private A_TipInit As String

Private WithEvents Combo1 As ComboBox
Attribute Combo1.VB_VarHelpID = -1




'参数介绍：
'oCtl：在用户UI上要显示数据的控件，传址
'vSQL：获取数据的SQL
'vFieldNameDisplay：显示在控件上的字段
'vFieldNameGet：用户选择之后本类实际返回的字段名称
'vFieldNameDefault:设置默认标志的字段的名称（可为空）
'vTip_Init: 在初始化失败后的提示信息 (不可为空)
Public Sub InitCls(ByRef oCtl As Object, ByVal vSQL As String, _
    ByVal vFieldNameDisplay As String, ByVal vFieldNameGet As String, _
    ByVal vFieldNameDefault As String, ByVal vTip_Init As String)
    
    Set oCombo = oCtl
    Set A_rs = New Recordset
    Debug.Print vSQL
    A_rs.Open vSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Debug.Print vSQL
    A_FieldNameDisplay = vFieldNameDisplay
    A_FieldNameGet = vFieldNameGet
    A_FieldNameDefault = vFieldNameDefault
    A_TipInit = vTip_Init
    
    
    A_bIsInit = True
End Sub

'根据初始化函数中的查询SQL获取记录集后将CAPTION字段显示在UI上，隐藏KEY字段
Public Sub ShowData()
    If A_bIsInit = False Then
        MsgBox A_TipInit, vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
    If A_rs.RecordCount <= 0 Then
        oCombo.Clear
        'MsgBox "记录集为空！", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
    
    oCombo.Clear
    A_rs.MoveFirst
    oCombo.AddItem ""
    Do While Not A_rs.EOF
        oCombo.AddItem A_rs(Trim$(A_FieldNameDisplay))
        A_rs.MoveNext
    Loop
    
    
    '显示默认
    If Len(A_FieldNameDefault) > 0 Then
        A_rs.Filter = " " & A_FieldNameDefault & "=1"
        If A_rs.RecordCount > 0 Then
            oCombo.Text = A_rs(Trim$(A_FieldNameDisplay))
        End If
    End If
End Sub




'根据当前UI上显示的CAPTION获取对应的KEY字段的VALUE
Public Function GetData() As String
    On Error Resume Next
    If A_bIsInit = False Then
        MsgBox A_TipInit, vbOKOnly + vbInformation, "提示"
        Exit Function
    End If
    
    A_rs.Filter = ""
    If A_rs.RecordCount <= 0 Then
        GetData = ""
        Exit Function
    End If
    
    If Len(Trim$(oCombo.Text)) <= 0 Then
        GetData = ""
        Exit Function
    End If
    
    A_rs.Filter = ""
    '在2016年12月为本工具类获取数据的时候添加了第一行为空数据
    '那么之后获取选中VALUE时候就不用+1了。
    'A_rs.Bookmark = Val(oCombo.ListIndex) + 1
    A_rs.Bookmark = Val(oCombo.ListIndex)
    Debug.Print A_rs(Trim$(A_FieldNameGet))
    GetData = A_rs(Trim$(A_FieldNameGet))
End Function

'获取ComboBox的绑定的记录集中，除了显示给用户看的CAPTION字段和主键字段之外的其他某一字段
Public Function GetData_OtherField(ByVal vFieldName As String) As String
    If A_bIsInit = False Then
        MsgBox A_TipInit, vbOKOnly + vbInformation, "提示"
        Exit Function
    End If
    
    A_rs.Filter = ""
    If A_rs.RecordCount <= 0 Then
        Exit Function
    End If
    
    If Len(Trim$(oCombo.Text)) <= 0 Then
        GetData_OtherField = ""
        Exit Function
    End If
    
    A_rs.Filter = ""
    '在2017年4月6日为新达更新客户端时发现index+1后的错误，于是修正
    'A_rs.Bookmark = Val(oCombo.ListIndex) + 1
    A_rs.Bookmark = Val(oCombo.ListIndex)
    GetData_OtherField = A_rs(Trim$(vFieldName))
End Function

'给定KEY字段的VALUE（vDesignation）之后显示出CAPTION字段的VALUE到UI上
'用于从持久层获取数据后显示用户VALUE
'用于OpenBill中，将持久化层的数据显示到用户UI上
Public Sub ShowData_Designation(ByVal vDesignation As String)
    On Error Resume Next
    Dim iListIndex As Long
    
    A_rs.Filter = ""
    A_rs.MoveFirst
    oCombo.ListIndex = -1
    Do While Not A_rs.EOF
        If A_rs(Trim$(A_FieldNameGet)) = vDesignation Then
            iListIndex = Val(A_rs.Bookmark) - 1
            oCombo.Text = A_rs(Trim$(A_FieldNameDisplay))
            Exit Do
        End If
        A_rs.MoveNext
    Loop
    
End Sub


