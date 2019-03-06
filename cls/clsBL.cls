VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'----程序名称:主从表单据基类                       As String   '单据类别
'保持属性值的局部变量
Private mvarboolIsLoadType              As Boolean  '单据打开类别:True 新建 False 打开
'保持属性值的局部变量
Private mvarboolIsDraft                 As Boolean  '是否为草稿:True 是草稿 False 不是草稿
'保持属性值的局部
'----单位:常熟新东方软件
'----编制人:陈超

'----变量部分
'----公用变量
Private mvarm_BID                       As String
Private mvarboolIsSave                  As Boolean  '是否需要保存:True 是 False 否
'----以下为单据状态部分

Public intIsClosed                      As Integer  '当前单据状态
'----控件
Public WithEvents TDBGrid1              As TDBGrid  '明细表网格
Attribute TDBGrid1.VB_VarHelpID = -1
Public WithEvents ActiveBar21           As ActiveBar2
Attribute ActiveBar21.VB_VarHelpID = -1
Public Adodc1                           As Adodc   '主表
Public Adodc2                           As Adodc   '明细表
Public ScriptControl1                   As Object
Public frm1                             As Object
Private clsReport1                      As clsReport '报表类

'----私有变量
Public m_TableName                      As String   '主表名称
Public m_TableDetailName                As String   '明细表名称
Private m_DetailSQL                     As String   '明细表SQL语句

Private m_MainTable                     As String   '登帐表名称
Private m_DraftMainTable                As String   '草稿表名称
Private m_DetailTable                   As String   '登帐明细表名称
Private m_DraftDetailTable              As String   '草稿明细表名称
Private m_DeleteNullDetail              As Long     '是否删除第一个字段为空的明细记录(登帐时候)

Private rsBLField                       As New RecordSet '取得表单字段属性

Private m_Formula                       '明细表公式
Private m_MaxFormulaIndex               '最大公式明细列

Private m_GridIndex                     '指当前需要重新 Val 的 Index列表
Private m_ReportID                      As String           '打印索引

Private m_Mark
'----类实现
'保持属性值的局部变量
Private mvarm_MaxIndex As Integer '最大网格显示编号,只能读取

Public Event OnGetBillInformation()     '当取得单据信息时引发
Public Event OnUpdateFrameType()        '当表单重新需要刷新时引发
Public Event OnSaveFrame()                                      '保存表单
Public Event OnAddNewItem(ByRef Adodc2 As Object, ByRef Cancel As Integer)             '当新增一项时触发
Public Event OnAddNewObject(ByRef Adodc1 As Object)                       '当新增表单时触发
Public Event OnOpenFrame()                                      '当打开表单时
Public Event OnClearInputItemDetail(ByRef Cancel As Integer)    '当清除一项时
Public Event OnGetDetailSql(ByRef strSQL As String)             '当取得明细表SQL时
Public Event OnInitFrame()
Public Event OnCheckInputFrame(ByRef OK As Boolean)
Public Event OnCheckOnInputItem(ByRef OK As Boolean)
Public Event OnBillCheckIn(ByRef oCnn As Connection)
Public Event OnIDChange(ByVal m_ID As Long)
Public Event OnUpdateSum(ByVal m_Sum As Double, ByVal m_FieldName As String)

Public Event AfterMoveBill()   '当移动单据（第一单、前单、后单、最后单）

'下面的事件在2014年12月6日由chanchaw添加
'==================================================
Public Event BeforeDeleteBill(ByVal vID As Long)
Public Event AfterDeleteBill(ByVal vID As Long)
'==================================================


'保持属性值的局部变量
Private mvariFrameLoadType As Integer '局部复制
'保持属性值的局部变量
Private mvarObjectID As String '局部复制
'保持属性值的局部变量
Private mvarm_ID As Long '局部复制



'============================
Private m_RowIndex As Long  '单据120007  每新增一条记录累计1(即前面添加一列序列号)
'============================
Private mvarDataShow As Boolean

Private A_rsBillField As New RecordSet  '主表字段的设置
Private A_rsBLField As New RecordSet   '明细表字段的设置
Private A_rsIPF As RecordSet   '从表G_IPFilterBL获取参数
Private strSQL As String

Private A_IsVoid As Integer   '1表示该单据已经被作废，0表示正常
Private A_BLFilterSQL_Bill As String   '从表G_BL中的字段B_FilterSQL_Bill中获取的过滤条件
Private A_BLFilterSQL_Detail As String '从表G_BL中的字段B_FilterSQL_Detail中获取的过滤条件


'=============================
'作废功能的配置文件
Private A_rsConfig_SetVoid As RecordSet
Public W_SetVoidEnabled As Boolean   'false为不可使用作废功能，true表示可以使用
'=============================

Private A_CreateOperator As String '该字段表示使用登录软件的用户名匹配该变量内的字段的VALUE
                                '即，制单员只能看到自己做的单据

Public Property Let DraftDetailTable(ByVal vData As String)
    m_DraftDetailTable = vData
End Property

Public Property Get DraftDetailTable() As String
    DraftDetailTable = m_DraftDetailTable
End Property


Public Property Let m_DataShow(ByVal vData As Boolean)
    mvarDataShow = vData
End Property

Public Property Get m_DataShow() As Boolean
    m_DataShow = mvarDataShow
End Property

Public Property Let m_ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.m_ID = 5
    mvarm_ID = vData
    RaiseEvent OnIDChange(mvarm_ID)
End Property

Public Property Get m_ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.m_ID
    m_ID = mvarm_ID
End Property

Public Property Let ObjectID(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ObjectID = 5
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ObjectID
    ObjectID = mvarObjectID
End Property

Public Property Let iFrameLoadType(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.iFrameLoadType = 5
    mvariFrameLoadType = vData
End Property

Public Property Get iFrameLoadType() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.iFrameLoadType
    iFrameLoadType = mvariFrameLoadType
End Property

Public Property Let boolIsSave(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.boolIsSave = 5
    If vData <> mvarboolIsSave Then
        mvarboolIsSave = vData
        UpdateFrameType
    End If
End Property

Public Property Get boolIsSave() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.boolIsSave
    boolIsSave = mvarboolIsSave
End Property

Public Property Let boolIsDraft(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.boolIsDraft = 5
    If vData <> mvarboolIsDraft Then
        mvarboolIsDraft = vData
        UpdateFrameType
    End If
End Property

Public Property Get boolIsDraft() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.boolIsDraft
    boolIsDraft = mvarboolIsDraft
End Property

Public Property Let boolIsLoadType(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.boolIsLoadType = 5
    If vData <> mvarboolIsLoadType Then
        mvarboolIsLoadType = vData
    End If
End Property

Public Property Get boolIsLoadType() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.boolIsLoadType
    boolIsLoadType = mvarboolIsLoadType
End Property

Public Property Let m_BID(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.mvarm_BID = 5
    If Trim(vData) <> mvarm_BID Then
        mvarm_BID = vData
    End If
End Property

Public Property Get m_BID() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.mvarm_BID
    m_BID = mvarm_BID
End Property

Public Property Get m_MaxIndex() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.m_MaxIndex
    m_MaxIndex = mvarm_MaxIndex
End Property

'----初始化类
'----除控件外数据进行重新整理
Public Sub InitClass(ByRef oAdodc1 As Object, ByRef oAdodc2 As Object, ByRef oTDBGrid1 As Object, ByRef oAB As ActiveBar2)
    Set Adodc1 = oAdodc1
    Set Adodc2 = oAdodc2
    Set TDBGrid1 = oTDBGrid1
    Set ActiveBar21 = oAB
    Set frm1 = TDBGrid1.Parent
    
    Gm.RefreshCnn
    SetScriptControl
    SetFielsList
    InitReportList
    GetBillInformation



    '检测当前单据是否有数据需要隐藏
    SetDataShow
    
    '获取字段显示格式
    GetFieldFormat
    GetBillField
    
    
    '初始化作废功能
    InitSetVoid
    
    
    '初始化
    '2013年5月23日 21:26:26
    InitInputFilterBL
    
    
    'Dim cls1 As New clsWorkFlow
    
End Sub

'2013年5月23日 21:26:22
'初始化本对象网格中允许输入的字符
Private Sub InitInputFilterBL()

    Set A_rsIPF = New RecordSet
    strSQL = "Select * From G_IPFilterBL Where B_ObjectID='" & mvarObjectID & "'"
    A_rsIPF.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
End Sub


'获取网格列参数设置
Private Sub GetFieldFormat()
    Dim strSQL As String
    Set A_rsBLField = New RecordSet
    strSQL = "Select * From G_BLField Where B_ObjectID='" & mvarObjectID & "'"
    Debug.Print strSQL
    A_rsBLField.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
End Sub


Private Sub GetBillField()
    Dim strSQL As String
    Set A_rsBillField = New RecordSet
    strSQL = "Select * From G_BLFormTools Where B_ObjectID='" & mvarObjectID & "'"
    Debug.Print strSQL
    A_rsBillField.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
End Sub


'根据表G_DataShow设置数据的显示与隐藏
Private Sub SetDataShow()
    Dim strSQL As String
    Dim rs As RecordSet
    
    Set rs = New RecordSet
    strSQL = "Select * From G_DataShow Where B_ObjectID='" & mvarObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs.RecordCount <= 0 Then
        mvarDataShow = False
    Else
        mvarDataShow = True
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub InitReportList()
    On Error Resume Next
    Dim aReport
    Dim i As Long
    Dim rs As New RecordSet
    Dim strSQL As String

    Set rs = New RecordSet
    strSQL = "Select * From G_BL Where B_ObjectID='" & ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly


    '填充报表
    aReport = Split(rs("B_ReportObject"), ",")
    
    For i = 0 To ((UBound(aReport) + 1) / 2) - 1
        ActiveBar21.Bands("Band1").Tools("报表").CBAddItem aReport(i) & "," & GetReportName(aReport(i))
    Next
    
    ActiveBar21.Bands("Band1").Tools("报表").CBListIndex = 0

    aReport = Split(ActiveBar21.Bands("Band1").Tools("报表").Text, ",")
    
    Set clsReport1 = New clsReport
    With clsReport1
        Debug.Print aReport(0)
        .InitClass frm1, ObjectID, aReport(0)
        .Refresh
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

Private Function GetReportName(ByVal m_ReportObjectID As String) As String
    Dim rs As New RecordSet
    Dim strSQL As String
    
    Set rs = New RecordSet
    strSQL = "Select B_ReportName From G_Report Where B_ObjectID='" & Trim(m_ReportObjectID) & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then
        GetReportName = rs(0)
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Sub SetScriptControl()
    Set ScriptControl1 = CreateObject("ScriptControl")
    ScriptControl1.Language = "VBScript"
    ScriptControl1.AddObject "TDBGrid1", TDBGrid1
    
End Sub

'----新建一张单据
Public Sub AddNewObject()
    If ExitFrame = True Then
        AddNewFrame
    End If
End Sub

Public Sub LoadObject()
    '如果只有一张单据,则新增
    Dim lID As Long
    lID = GetLastBill
    If lID = 0 Then
        AddNewFrame
    Else
        m_ID = lID
        OpenFrame
    End If
    
End Sub

Private Function GetLastBill() As Long
    Dim strSQL As String
    Dim rs As New RecordSet
    
    Set rs = New RecordSet
    strSQL = "Select Top 1 B_ID From " & m_DraftMainTable & " Where B_ObjectID='" & Trim(ObjectID) & "'"
    strSQL = MontageCreateOperator(strSQL)
    strSQL = strSQL & " Order By B_ID Desc"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        GetLastBill = rs(0)
        rs.Close
        Set rs = Nothing
        Exit Function
    Else
        GetLastBill = 0
        rs.Close
        Set rs = Nothing
        Exit Function
    End If

End Function

'----保存单据
Public Sub SaveFrame()
    On Error Resume Next

    CountBillValue
    UpdateSum
    TDBGrid1.Update
    RaiseEvent OnSaveFrame
    mvarboolIsLoadType = False
    
    
    '登帐后写入工作流实例表
    WriteWFInstance_Save mvarm_ID
End Sub

'----关闭表单
Public Function ExitFrame() As Boolean
    On Error Resume Next
    '----当修改的单据退出后自动登帐
    If mvarboolIsSave = True And (intIsClosed = 0 Or intIsClosed = 1) Then
        Dim i As Integer
        i = MsgBox("单据数据已经修改是否要进行保存？", vbInformation + vbYesNoCancel + vbDefaultButton3, "关闭")
        Select Case i
            Case vbYes
                '保存
                SaveFrame
            Case vbNo
                '---放弃保存
                '---当为草稿,并且为新增状态时:对单据进行删除
                If mvarboolIsDraft = True And mvarboolIsLoadType = True And intIsClosed = 0 Then
                    DeleteFrame
                End If
            Case vbCancel
                '取消
                ExitFrame = False
                Exit Function
        End Select
    End If
    ExitFrame = True
    If ExitFrame = True And mvarboolIsDraft = False Then
        If intIsClosed = 2 Then
            If Not (Adodc1.RecordSet Is Nothing) And Adodc1.RecordSet.RecordCount > 0 Then
                
                Adodc1.RecordSet("B_Closed") = 1
                Adodc1.RecordSet.Update
            End If
        End If
    End If

End Function

'----清除当前不需使用的项目
Public Sub ClearInputItemDetail()
    Dim Cancel As Integer
    Dim m_DeleteItemList As String
    
    '不需要删除第一个字段为空的记录的话   则直接退出
    If m_DeleteNullDetail = 0 Then
        Exit Sub
    End If
    
    
    With Adodc2.RecordSet
        If .RecordCount < 1 Then
            Exit Sub
        End If
        
        .MoveFirst
        Do While Not .EOF
            If .Fields(0).Type = adBoolean Then
                If .Fields(0).Value = False Then
                    m_DeleteItemList = m_DeleteItemList & Trim(str(Adodc2.RecordSet("B_ItemID"))) & ","
                End If
            Else
                If Len(Trim(.Fields(0).Value)) < 1 Or IsNull(.Fields(0).Value) Then
                    m_DeleteItemList = m_DeleteItemList & Trim(str(Adodc2.RecordSet("B_ItemID"))) & ","
                Else
                    Cancel = 0
                    RaiseEvent OnClearInputItemDetail(Cancel)
                    If Cancel = 1 Then
                        m_DeleteItemList = m_DeleteItemList & Trim(str(Adodc2.RecordSet("B_ItemID"))) & ","
                    End If
                End If
            End If

            .movenext
        Loop
        .MoveFirst
    End With
    If Len(m_DeleteItemList) > 1 Then
        m_DeleteItemList = Mid(m_DeleteItemList, 1, Len(m_DeleteItemList) - 1)
        GoDelete m_DeleteItemList
    End If
End Sub

'复制表单
Private Sub CopyFrame()
    If MsgBox("是否要复制单据?", vbExclamation + vbOKCancel + vbDefaultButton2, "复制") = vbCancel Then
        Exit Sub
    End If
    On Error GoTo IFERR
    
    Dim m_OldID As Long
    Dim m_MainFieldList As String
    Dim m_DetailFieldList As String
    Dim strSQL As String
    '新增一张表单
    m_OldID = m_ID
    AddNewFrame
    '将原单据数据复制到新表单上
    
    '将登帐表复制到草稿表中
    'm_MainTable                     As String   '登帐表名称
    'm_DraftMainTable                As String   '草稿表名称
    'm_DetailTable                   As String   '登帐明细表名称
    'm_DraftDetailTable              As String   '草稿明细表名称
    
    GetBillFieldList m_DetailFieldList
    
    strSQL = "Delete From " & m_DraftDetailTable
    strSQL = strSQL & " Where B_ID=" & m_ID
    Gm.cnnTool.cnn.Execute strSQL
    
    strSQL = "Insert Into " & m_DraftDetailTable
    strSQL = strSQL & " (B_ID," & m_DetailFieldList & ")"
    strSQL = strSQL & " Select " & m_ID & "," & m_DetailFieldList & " From " & m_DetailTable
    strSQL = strSQL & " Where B_ID=" & m_OldID

    Gm.cnnTool.cnn.Execute strSQL
    
    Adodc2.RecordSet.requery
    boolIsSave = False
    Exit Sub
IFERR:
    MsgBox Err.Description
    Exit Sub
End Sub


'复制表单
'功能介绍:
'原版本只复制明细数据
'现新版本主表与明细表的数据都复制
'并且同时删除被复制的原始单据
Private Sub CopyFrameEx()
    If MsgBox("是否要复制单据?", vbExclamation + vbOKCancel + vbDefaultButton2, "复制") = vbCancel Then
        Exit Sub
    End If
    On Error GoTo IFERR
    
    Dim m_OldID As Long
    Dim m_MainFieldList As String
    Dim m_DetailFieldList As String
    'Dim m_BillFieldList As String
    Dim strSQL As String
    
    Dim i As Long
    Dim rs As RecordSet
    Dim rs1 As RecordSet
    
    
    '新增一张表单
    m_OldID = m_ID
    AddNewFrame
    '将原单据数据复制到新表单上
    
    '将登帐表复制到草稿表中
    'm_MainTable                     As String   '登帐表名称
    'm_DraftMainTable                As String   '草稿表名称
    'm_DetailTable                   As String   '登帐明细表名称
    'm_DraftDetailTable              As String   '草稿明细表名称
    
    
    '获取明细表字段字符串
    GetBillFieldList m_DetailFieldList
    
    '获取主表字段字符串
    GetBillFieldListBill m_MainFieldList
    
    
    
    '删除新创建的草稿表中的草稿数据(其实没有数据)
    strSQL = "Delete From " & m_DraftDetailTable
    strSQL = strSQL & " Where B_ID=" & m_ID
    Gm.cnnTool.cnn.Execute strSQL
    
    
    
    '复制主表数据
    '==========================
    '先获取记录集
    Set rs = New RecordSet
    strSQL = "Select " & m_MainFieldList & " From " & m_MainTable & " Where B_ID=" & m_OldID
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    '语句方法效率高,速度快,不过不可控制写入的数据
'    strSQL = "Insert Into " & m_DraftMainTable
'    strSQL = strSQL & " (B_CodeID," & m_MainFieldList & ")"
'    strSQL = strSQL & " Select '" & Adodc1.RecordSet("B_CodeID") & "'" & "," & m_MainFieldList
'    strSQL = strSQL & " From " & m_MainTable
'    strSQL = strSQL & " Where B_ID=" & m_OldID
'    Gm.cnnTool.cnn.Execute strSQL


    Set rs1 = New RecordSet
    strSQL = "Select * From " & m_DraftMainTable & " Where B_ID=" & m_ID
    rs1.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    'rs1("B_CodeID") = Adodc1.RecordSet("B_CodeID")
    For i = 0 To rs1.Fields.Count - 1
        Select Case rs1(i).name
            Case "B_CodeID"
                rs1(i) = Adodc1.RecordSet("B_CodeID")
            Case "B_Closed"
                rs1(i) = 0
            Case "B_ID"
            Case Else
                Debug.Print rs1(i).name & "=" & rs(rs1(i).name)
                rs1(i) = rs(rs1(i).name)
        End Select
        rs1.Update
    Next
    rs1.Close
    Set rs1 = Nothing
    
    rs.Close
    Set rs = Nothing
    
    
    '主表上的控件数据复制
    GoOpenFrame
    'Adodc1.RecordSet.Requery
    '==========================
    
    
    '复制明细数据
    strSQL = "Insert Into " & m_DraftDetailTable
    strSQL = strSQL & " (B_ID," & m_DetailFieldList & ")"
    strSQL = strSQL & " Select " & m_ID & "," & m_DetailFieldList & " From " & m_DetailTable
    strSQL = strSQL & " Where B_ID=" & m_OldID
    Gm.cnnTool.cnn.Execute strSQL
    
    
    
    Adodc2.RecordSet.requery
    
    
    '删除被复制的单据
    '==========================
'    strSQL = "Delete From " & m_DetailTable & " Where B_ID=" & m_OldID
'    Gm.cnnTool.cnn.Execute strSQL
'
'    strSQL = "Delete From " & m_MainTable & " Where B_ID=" & m_OldID
'    Gm.cnnTool.cnn.Execute strSQL
    '==========================
    
    boolIsSave = False
    Exit Sub
IFERR:
    MsgBox Err.Description
    Exit Sub
End Sub


'获取主表字段
Private Sub GetBillFieldListBill(ByRef m_MainFieldList As String)
     '----以下为增加项的SQL
    'Insert Into MainTable (Fields...) Select (Fields...) From MainTable Where B_ID=m_ID
    Dim rs As New RecordSet        '----草稿记录
    Dim strSQL As String
    Dim i As Integer
    
    Set rs = New RecordSet
    
    '----查询主表字段
    'm_MainFieldList As String
    strSQL = "Select * From " & m_DraftMainTable & " Where 1=0"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    m_MainFieldList = ""
    For i = 0 To rs.Fields.Count - 1
        If rs.Fields(i).name <> "B_ID" And rs.Fields(i).name <> "B_CodeID" Then
            m_MainFieldList = m_MainFieldList & Trim(rs.Fields(i).name) & ","
        End If
    Next
    m_MainFieldList = Mid(m_MainFieldList, 1, Len(m_MainFieldList) - 1)

    rs.Close
    Set rs = Nothing

End Sub


Private Sub GetBillFieldList(ByRef m_DetailFieldList As String)
    '----以下为增加项的SQL
    'Insert Into MainTable (Fields...) Select (Fields...) From MainTable Where B_ID=m_ID
    Dim rs As New RecordSet        '----草稿记录
    Dim strSQL As String
    Dim i As Integer
    
    Set rs = New RecordSet
    
'    '----查询主表字段
'    'm_MainFieldList As String
'    strSQL = "Select * From " & m_DraftMainTable & " Where 1=0"
'    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
'    m_MainFieldList = ""
'    For i = 0 To rs.Fields.Count - 1
'        If rs.Fields(i).name <> "B_ID" And rs.Fields(i).name <> "B_CodeID" Then
'            m_MainFieldList = m_MainFieldList & Trim(rs.Fields(i).name) & ","
'        End If
'    Next
'    m_MainFieldList = Mid(m_MainFieldList, 1, Len(m_MainFieldList) - 1)
'
'    rs.Close
'    Set rs = Nothing
'    '----查询明细表字段
'    'm_DetailFieldList As String
'    Set rs = New Recordset

    strSQL = "Select * From " & m_DraftDetailTable & " Where 1=0"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    m_DetailFieldList = ""
    For i = 0 To rs.Fields.Count - 1
        If rs.Fields(i).name <> "B_ID" And rs.Fields(i).name <> "B_ItemID" Then
            m_DetailFieldList = m_DetailFieldList & Trim(rs.Fields(i).name) & ","
        End If
    Next
    m_DetailFieldList = Mid(m_DetailFieldList, 1, Len(m_DetailFieldList) - 1)
    
    rs.Close
    Set rs = Nothing
End Sub

'----增加一张单据
Public Sub AddNewFrame()
    '判断是否有新增的权限
    If Gm.PI.JudgeNew(Me.ObjectID) = False Then
        Exit Sub
    End If

    On Error GoTo IFERR
    Dim strSQL As String
    m_ID = 0
    mvarboolIsDraft = True
    m_RowIndex = 0
    GetBillInformation
    '----得到单据信息

    strSQL = "Select * From " & m_TableName & " Where 1=0"
    With Adodc1
        
        .ConnectionString = Gm.cnnTool.cnnStr

        .RecordSource = strSQL
        .Refresh

        .RecordSet.AddNew

        .RecordSet("B_CodeID") = GetFrameCodeDetail
        .RecordSet("B_Closed") = 0
        .RecordSet("B_BID") = Trim(mvarm_BID)
        .RecordSet("B_Date") = Format(Now, "YYYY-MM-DD HH:MM:SS")
        .RecordSet("B_ObjectID") = mvarObjectID
        
        .RecordSet("B_UserName") = Gm.SysID.SystemUserName
        '.RecordSet("B_Course") = clsSParameter1.GetParameterString("Course")

        '设置主表字段的默认值
        'SetBillFieldDefault Adodc1.RecordSet

        RaiseEvent OnAddNewObject(Adodc1)

        .RecordSet.Update
        '.RecordSet.Requery
 
        m_ID = .RecordSet("B_ID")
    End With

    strSQL = m_DetailSQL & " And " & Trim(m_TableDetailName) & ".B_ID=" & m_ID & " Order By B_ItemID"
    
    With Adodc2
        .ConnectionString = Gm.cnnTool.cnnStr

        RaiseEvent OnGetDetailSql(strSQL)

        .RecordSource = strSQL
        .Refresh
    End With
    TDBGrid1.DataSource = Adodc2.RecordSet
    intIsClosed = 0
    AddNewItem

    RaiseEvent OnInitFrame
    
    GetGridShowMax
    '设置为新增状态
    boolIsLoadType = True
    boolIsSave = False

    InitFormula
    UpdateSum

    '----当需要重新刷新时引发
    UpdateFrameType
    
    
    
    '补救性修改当前新增了本单据的用户名
    'BJUserName m_ID
    
    
    '=======================================
    Dim clsOptionButton1 As New clsOptionButton
    clsOptionButton1.ClareOptionValue frm1
    '=======================================
    
    Exit Sub
IFERR:
    MsgBox Err.Description
    Exit Sub
End Sub

'新增一个单据后修改其B_UserName为当前的UserName
Private Sub BJUserName(ByVal vID As Long)
    Dim strSQL As String
    strSQL = "Update " & m_DraftMainTable & " Set B_UserName='" & Gm.SysID.SystemUser & "'"
    strSQL = strSQL & " Where B_ID=" & vID
    Debug.Print strSQL
    Gm.cnnTool.cnn.Execute strSQL
End Sub

'----设置网格显示的情况
Public Sub GetGridShowMax()
    '----取得当前最大网格列
    '----m_MaxIndex
    Dim i As Long
    For i = 0 To TDBGrid1.Columns.Count - 1
        If TDBGrid1.Columns(i).Visible = True Then
            mvarm_MaxIndex = i
        End If
    Next

    InitGridIndexList
End Sub

'----取得单据信息
Private Sub GetBillInformation()
    Dim rs As New RecordSet
    Dim strSQL  As String


    Set rs = New RecordSet
    strSQL = "Select B_MainTable,B_DraftMainTable,B_DetailTable,B_DraftDetailTable,B_BID,"
    strSQL = strSQL & " B_DraftDetailSql,B_DetailSql,B_DeleteNullDetail,B_FilterSQL_Bill,B_FilterSQL_Detail,B_CreateOperator"
    strSQL = strSQL & " From G_BL Where B_ObjectID='" & Trim(mvarObjectID) & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then

        m_MainTable = rs("B_MainTable")
        m_DraftMainTable = rs("B_DraftMainTable")
        m_DetailTable = rs("B_DetailTable")
        m_DraftDetailTable = rs("B_DraftDetailTable")
        m_DeleteNullDetail = IIf(IsNull(rs("B_DeleteNullDetail")), 0, rs("B_DeleteNullDetail"))
        A_BLFilterSQL_Bill = IIf(IsNull(rs("B_FilterSQL_Bill")), "", rs("B_FilterSQL_Bill"))
        A_BLFilterSQL_Detail = IIf(IsNull(rs("B_FilterSQL_Detail")), "", rs("B_FilterSQL_Detail"))

        mvarm_BID = rs("B_BID")
        
        '制单员
        A_CreateOperator = IIf(IsNull(rs("B_CreateOperator")), "", rs("B_CreateOperator"))
        
        
        ChangeBillType

        '----如果为草稿
        If mvarboolIsDraft = True Then
            m_TableName = rs("B_DraftMainTable")
            m_TableDetailName = rs("B_DraftDetailTable")
            m_DetailSQL = rs("B_DraftDetailSql")

        Else
            m_TableName = rs("B_MainTable")
            m_TableDetailName = rs("B_DetailTable")
            m_DetailSQL = rs("B_DetailSql")
        End If
        
    End If

    rs.Close
    Set rs = Nothing

    '----当重新取得状态时引发
    RaiseEvent OnGetBillInformation
End Sub

'----增加一项
Public Sub AddNewItem()
    On Error GoTo IFERR
    '判断是否有新增的权限
    If Gm.PI.JudgeNew(Me.ObjectID) = False Then
        Exit Sub
    End If


    On Error GoTo IFERR
    If intIsClosed <> 0 Then
        Exit Sub
    End If
    
    Dim i As Integer
    Dim Cancel As Integer
    Adodc2.RecordSet.AddNew
    Adodc2.RecordSet("B_ID") = m_ID
    
    RaiseEvent OnAddNewItem(Adodc2, Cancel)
    If Cancel = 1 Then
        Adodc2.RecordSet.CancelUpdate
        Adodc2.RecordSet.requery
    Else
        Adodc2.RecordSet.Update
        TDBGrid1.PostMsg 81
        TDBGrid1.Update
        TDBGrid1.Col = 0
        mvarboolIsSave = True
    End If
    '----当需要重新刷新时引发
    UpdateFrameType
    Exit Sub
IFERR:
    Dim szTip As String
    szTip = "在增加明细时出错" & vbNewLine & Err.Description
    MsgBox szTip, vbOKOnly + vbInformation, "提示"
    Exit Sub
End Sub

'----删除一项
Public Sub DeleteItem()

    '判断是否有删除的权限
    If Gm.PI.JudgeDelete(Me.ObjectID) = False Then
        Exit Sub
    End If
    
    
    On Error Resume Next
    If Adodc2.RecordSet.RecordCount < 1 Then
        Exit Sub
    End If
    If Len(Trim(Adodc2.RecordSet(0))) > 0 Then
        If MsgBox("是否要删除这一项?", vbInformation + vbOKCancel + vbDefaultButton2, "删除") = vbOK Then
            GoDelete str(Adodc2.RecordSet("B_ItemID"))
        End If
    Else
        GoDelete str(Adodc2.RecordSet("B_ItemID"))
    End If
End Sub


Public Sub DeleteItemAll()
        '判断是否有删除的权限
    If Gm.PI.JudgeDelete(Me.ObjectID) = False Then
        Exit Sub
    End If
    
    
    On Error Resume Next
    If Adodc2.RecordSet.RecordCount < 1 Then
        Exit Sub
    End If
    
    
    
    If MsgBox("是否要删除所有明细记录?", vbInformation + vbOKCancel + vbDefaultButton2, "删除") = vbOK Then
    
        Adodc2.RecordSet.MoveFirst
        Do While Not Adodc2.RecordSet.EOF
            GoDelete str(Adodc2.RecordSet("B_ItemID"))
            Adodc2.RecordSet.movenext
        Loop
        
        
        Adodc2.RecordSet.MoveFirst
        GoDelete str(Adodc2.RecordSet("B_ItemID"))
    End If
    
End Sub


'----执行删除,可以删除多行
Private Sub GoDelete(ByVal m_DeleteItemList As String)
    On Error GoTo IFERR
    Dim strSQL As String

    Gm.cnnTool.cnn.BeginTrans
    strSQL = "Delete From " & m_TableDetailName & " Where B_ItemID IN (" & m_DeleteItemList & ")"
    Gm.cnnTool.cnn.Execute strSQL
    Gm.cnnTool.cnn.CommitTrans

    Adodc2.RecordSet.requery

    mvarboolIsSave = True
    UpdateFrameType
    Exit Sub
IFERR:

End Sub

'----执行删除单据
Private Function DeleteFrame() As Boolean
    On Error GoTo IFERR
    Dim strSQL As String
    
    
    RaiseEvent BeforeDeleteBill(m_ID)
    
    Gm.cnnTool.IniConnection8DM Gm.SysID.DBInfo
    'cnn.InitializeConnection

    Gm.cnnTool.cnn.BeginTrans

    strSQL = "Delete From " & m_TableDetailName & " Where B_ID=" & m_ID
    Gm.cnnTool.cnn.Execute strSQL

    strSQL = "Delete From " & m_TableName & " Where B_ID=" & m_ID
    Gm.cnnTool.cnn.Execute strSQL

    Gm.cnnTool.cnn.CommitTrans

    DeleteFrame = True
    
    
    RaiseEvent AfterDeleteBill(m_ID)
    
    Exit Function
IFERR:
    MsgBox Err.Description
    DeleteFrame = False
    Exit Function
End Function

'----测试这张单据是否能被删除
Private Function CheckDeleteFrame() As Boolean
    If intIsClosed = 1 Then
        MsgBox "这张单据已审核过了,不能进行删除.", vbExclamation, "删除"
        CheckDeleteFrame = False
        Exit Function
    End If
    CheckDeleteFrame = True
End Function

'----删除单据
Public Sub DeleteNowFrame()
    '判断是否有删除的权限
    If Gm.PI.JudgeDelete(Me.ObjectID) = False Then
        Exit Sub
    End If

    '已经出库的入库单数据，不可删除
    '120007为入库单
    If mvarObjectID = "120007" Then
        If JudgeCanDel = False Then
            Exit Sub
        End If
    End If
    
    

    If CheckDeleteFrame = False Then
        Exit Sub
    End If
    If mvarboolIsSave = True Then
        SaveFrame
    End If
    If MsgBox("是否要删除这张单据?", vbInformation + vbOKCancel + vbDefaultButton2, "删除") = vbOK Then
        '先删除连带的单据数据
        DelRelated
        
        If DeleteFrame = True Then
            Dim m_IDB As Long
            m_IDB = GetOtherFrame(4)
            If m_IDB > 0 Then
                m_ID = m_IDB
                ChangeBillType

                OpenFrame
            Else
                AddNewObject
            End If
        End If
    End If
End Sub

'----打开单据
Public Sub OpenFrame()
    If ExitFrame = True Then
        GoOpenFrame
    End If
End Sub

'----执行打开单据
Private Sub GoOpenFrame()
    On Error GoTo IFERR
    Dim strSQL As String
    
    
    '=======================================
    Dim clsOptionButton1 As New clsOptionButton
    clsOptionButton1.ClareOptionValue frm1
    '=======================================
    
    

    GetBillInformation
    
    strSQL = "Select * From " & m_TableName & " Where B_ID=" & m_ID & " " & A_BLFilterSQL_Bill
    strSQL = MontageCreateOperator(strSQL)
    Debug.Print strSQL
    With Adodc1
        .ConnectionString = Gm.cnnTool.cnnStr
        .RecordSource = strSQL
        .Refresh
    End With
    

    strSQL = m_DetailSQL & " And " & Trim(m_TableDetailName) & ".B_ID=" & m_ID & " " & A_BLFilterSQL_Detail
    strSQL = MontageCreateOperator(strSQL)
    strSQL = strSQL & " Order By B_ItemID"
    Debug.Print strSQL
    With Adodc2
        .ConnectionString = Gm.cnnTool.cnnStr

        RaiseEvent OnGetDetailSql(strSQL)

        .RecordSource = strSQL
        .Refresh
    End With
    TDBGrid1.DataSource = Adodc2.RecordSet

    UpdateSum

    RaiseEvent OnInitFrame
    RaiseEvent OnOpenFrame
    GetGridShowMax
    mvarboolIsLoadType = False
    mvarboolIsSave = False

    InitFormula

    UpdateFrameType
    
    
    '打开单据后就要初始化作废功能是否可用
    InitSetVoid

    Exit Sub
IFERR:
    MsgBox Err.Description
    Exit Sub
End Sub

'----初始化网格数据
Private Sub InitGridIndexList()
    Dim i As Integer
    Dim J As Integer
    Dim m_Index As Integer
    J = 0
    ReDim m_GridIndex(0 To 0)
    For i = 0 To Adodc2.RecordSet.Fields.Count - 1
        '----当字段为数据型时
        If Adodc2.RecordSet.Fields(i).Type = adInteger Or Adodc2.RecordSet.Fields(i).Type = adDouble Or Adodc2.RecordSet.Fields(i).Type = adBigInt Then
            '并且在网格中显示时
            m_Index = GetGridIndex(Adodc2.RecordSet.Fields(i).name)
            If m_Index > 0 And TDBGrid1.Columns(i).Visible = True Then
                m_GridIndex(J) = m_Index
                J = J + 1
                ReDim Preserve m_GridIndex(0 To J)
            End If
        End If
    Next
End Sub

'----初始化明细表公式
Private Sub InitFormula()
    'm_Formula
    Dim m_TmpFormula As String
    Dim m_FieldName As String
    Dim i As Integer
    Dim J As Integer
    Dim m_Index As Integer

    rsBLField.Sort = "B_Index"
    rsBLField.Filter = "B_UseFormula=1 "
    If rsBLField.RecordCount > 0 Then
        rsBLField.MoveFirst

        '----0.当前网格行
        '----1.
        ReDim m_Formula(0 To 1, 0 To rsBLField.RecordCount)
        J = 0
        Do While Not rsBLField.EOF
            '----进行公式替换
            m_TmpFormula = rsBLField("B_Formula")
            If Len(Trim(m_TmpFormula)) > 0 Then

                m_TmpFormula = Replace(m_TmpFormula, "[", "TDBGrid1.Columns(")
                m_TmpFormula = Replace(m_TmpFormula, "]", ").Value")

                For i = 0 To TDBGrid1.Columns.Count - 1
                    m_FieldName = TDBGrid1.Columns(i).DataField
                    m_Index = str(GetGridIndex(m_FieldName))
                    m_TmpFormula = Replace(m_TmpFormula, m_FieldName, Trim(str(m_Index)))
                Next
                m_Index = str(GetGridIndex(rsBLField("B_FieldName")))
                m_Formula(0, J) = m_Index
                m_Formula(1, J) = m_TmpFormula
                J = J + 1

            End If
            rsBLField.movenext
        Loop
        
        
    End If
    m_MaxFormulaIndex = J
    rsBLField.Filter = ""
End Sub

'----进行公式计算
Public Sub UseFormulaCount()
    'on error GoTo IFERR
    On Error Resume Next
    '----使用C1Awk1 进行计算
    '----计算符可以为 + - * / Mod Max Min
    '----字段必须使用 [] 进行分开
    '----计算明细表数据

    Dim m_Value As Double
    Dim i As Integer
    Dim J As Integer
    Dim m_FieldName As String
    Dim m_Index As Integer

    For i = 0 To m_MaxFormulaIndex - 1
        '---使用当前网格数据去除非法字符
    
        m_Value = ScriptControl1.Eval(m_Formula(1, i))
    
        m_Index = Val(m_Formula(0, i))
        TDBGrid1.Columns(m_Index).Value = m_Value
    Next
    TDBGrid1.Update

    UpdateSum
    TDBGrid1.SetFocus
    
    
'IFERR:
'    Exit Sub
End Sub

'----取得网格列数
Private Function GetGridIndex(ByVal m_FieldName As String) As Integer
    Dim i As Integer

    For i = 0 To TDBGrid1.Columns.Count - 1
        If Trim(m_FieldName) = Trim(TDBGrid1.Columns(i).DataField) Then
            GetGridIndex = i
            Exit Function
        End If
    Next
End Function

'----合计计算
Public Sub UpdateSum()
    '设置当前需要求合的字段
    If Adodc2.RecordSet Is Nothing Then
        Exit Sub
    End If

    '----设置当前临明数据表
    Dim rs1 As New RecordSet
    Dim m_Sum As Double                     '----合计值

    rsBLField.Filter = "B_IsCount=1 And B_FrameType=0"
    If rsBLField.RecordCount > 0 Then
        '复制字段
        Set rs1 = Adodc2.RecordSet.Clone
        rsBLField.MoveFirst
    Else
        Exit Sub
    End If

    Do While Not rsBLField.EOF
        m_Sum = 0

        If rs1.RecordCount > 0 Then
            rs1.MoveFirst
        End If
        Do While Not rs1.EOF
            If Not IsNull(rs1(Trim(rsBLField("B_FieldName")))) Then
                m_Sum = m_Sum + rs1(Trim(rsBLField("B_FieldName")))
            End If

            '----显示网格
            ShowGridFootText m_Sum, rsBLField("B_FieldName")
            
            RaiseEvent OnUpdateSum(m_Sum, rsBLField("B_FieldName"))
            rs1.movenext
            
        Loop
        
    
        rsBLField.movenext
    Loop

    
End Sub

'----统计表单数据
Private Sub CountBillValue()
    '----统计合计字段
    rsBLField.Filter = "B_IsCount=1 And B_FrameType=1"
    'cnn.InitializeConnection
    Dim strSQL As String
    If rsBLField.RecordCount > 0 Then
        rsBLField.MoveFirst
    Else
        'cnn.TerminateConnection
        Exit Sub
    End If
    '----建查询语句
    strSQL = "Update " & m_TableName & " Set "

    Do While Not rsBLField.EOF
        strSQL = strSQL & Trim(m_TableName) & "." & Trim(rsBLField("B_FieldName")) & "=P1." & Trim(rsBLField("B_FieldName")) & ","
        rsBLField.movenext
    Loop
    strSQL = Mid(strSQL, 1, Len(strSQL) - 1)

    strSQL = strSQL & " From (Select B_ID,"

    rsBLField.MoveFirst
    Do While Not rsBLField.EOF
        strSQL = strSQL & Trim(rsBLField("B_Formula")) & " as " & Trim(rsBLField("B_FieldName")) & ","
        rsBLField.movenext
    Loop
    strSQL = Mid(strSQL, 1, Len(strSQL) - 1)

    strSQL = strSQL & " From " & m_TableDetailName
    strSQL = strSQL & " Where B_ID=" & m_ID & " Group By B_ID) as P1"
    strSQL = strSQL & " Where P1.B_ID = " & Trim(m_TableName) & ".B_ID"
    strSQL = strSQL & " And " & Trim(m_TableName) & ".B_ID = " & m_ID

    Gm.cnnTool.cnn.Execute strSQL

    'cnn.TerminateConnection
End Sub

'----显示网格脚部文字
Private Sub ShowGridFootText(ByVal m_Sum As Double, ByVal m_FieldName As String)
    Dim i As Integer
    For i = 0 To TDBGrid1.Columns.Count - 1
        If Trim(m_FieldName) = Trim(TDBGrid1.Columns(i).DataField) Then
            A_rsBLField.Filter = " B_FieldName='" & m_FieldName & "'"
            If A_rsBLField.RecordCount > 0 Then
                TDBGrid1.Columns(i).FooterText = Format(m_Sum, IIf(IsNull(A_rsBLField!B_FieldFormat), "", A_rsBLField!B_FieldFormat))
                TDBGrid1.Columns(i).Alignment = A_rsBLField!B_Alignment
            Else
                
                TDBGrid1.Columns(i).FooterText = Format(m_Sum, "##,##0.00")
                TDBGrid1.Columns(i).Alignment = dbgRight
            End If
        End If
    Next
End Sub

'----以下为单据移动部分
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'----移到第一张单据
Public Sub MoveFirst()
    GetOtherBill 1
    
    RaiseEvent AfterMoveBill
End Sub

'----移到上一张单据
Public Sub MovePrevious()
    GetOtherBill 2
    
    RaiseEvent AfterMoveBill
End Sub

'----移到下一张单据
Public Sub movenext()
    GetOtherBill 3
    
    RaiseEvent AfterMoveBill
End Sub

'----移到最后一张单据
Public Sub movelast()
    GetOtherBill 4
    
    RaiseEvent AfterMoveBill
End Sub

'----进行单据移动
Private Sub GetOtherBill(ByVal m_Type As Integer)
    On Error Resume Next
    '先进保存
    If mvarboolIsSave = True Or mvarboolIsLoadType = True Then
        SaveFrame
    End If

    Dim m_IDB As Long
    m_IDB = GetOtherFrame(m_Type)
    If ExitFrame = True Then
        If m_IDB > 0 Then
            m_ID = m_IDB
            mvarboolIsLoadType = False
            ChangeBillType
            OpenFrame
        End If
    End If
End Sub

'----取得当前类别
Private Sub ChangeBillType()
    On Error GoTo IFERR
    'cnn.InitializeConnection

    If m_ID = 0 Then
        mvarboolIsDraft = True
        mvarboolIsLoadType = True
        Exit Sub
    Else
        mvarboolIsLoadType = False
    End If

    Dim rs As New RecordSet
    Dim strSQL As String

    Set rs = New RecordSet
    strSQL = "Select B_ID From " & m_DraftMainTable & " Where B_ID=" & m_ID
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        mvarboolIsDraft = True
        GoTo IFERR
    End If

    Set rs = New RecordSet
    strSQL = "Select B_ID From " & m_MainTable & " Where B_ID=" & m_ID
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        mvarboolIsDraft = False
        GoTo IFERR
    End If

    'cnn.TerminateConnection
IFERR:
    rs.Close
    Set rs = Nothing
    'cnn.TerminateConnection
    Exit Sub
End Sub

'----取得单据编号
Private Function GetBillCodeID() As String
    Dim rs As New RecordSet
    Dim strSQL As String
    Set rs = New RecordSet

    'cnn.InitializeConnection
    strSQL = "Select B_CodeID From " & m_TableName & " Where B_ID=" & m_ID
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        GetBillCodeID = rs(0)
    Else
        GetBillCodeID = ""
    End If
    rs.Close
    Set rs = Nothing

End Function

'----查询其它表单
Private Function GetOtherFrame(ByVal m_Type As Integer) As Long
    Dim rs As New RecordSet
    Dim strSQL As String
    Dim m_CodeID As String

    m_CodeID = GetBillCodeID
    If Len(Trim(m_CodeID)) < 11 Then
        Set rs = Nothing
        m_Type = 4
    End If

    Select Case m_Type
        Case 1  '----第一
            strSQL = "Select Top 1 P1.B_CodeID,P1.B_ID"
            strSQL = strSQL & " From ("
            strSQL = strSQL & " Select Min(B_CodeID) as B_CodeID,B_ID From " & m_DraftMainTable & " Where B_BID='" & mvarm_BID & "'" ' And B_UserName='" & Trim(m_typeSInformation.m_strUserName) & "'"
            strSQL = strSQL & " " & A_BLFilterSQL_Bill
            
            strSQL = MontageCreateOperator8Table(strSQL, m_DraftMainTable)
            
            strSQL = strSQL & " Group By B_ID"
            strSQL = strSQL & " Union"
            strSQL = strSQL & " Select Min(B_CodeID) as B_CodeID,B_ID From " & m_MainTable & " Where B_BID='" & mvarm_BID & "'" 'And B_UserName='" & Trim(m_typeSInformation.m_strUserName) & "'"
            strSQL = strSQL & " " & A_BLFilterSQL_Bill
            
            strSQL = MontageCreateOperator8Table(strSQL, m_MainTable)
            
            strSQL = strSQL & " Group By B_ID"
            strSQL = strSQL & " ) as P1"
            strSQL = strSQL & " Order By P1.B_CodeID"
        Case 2  '----上一
            strSQL = "Select Top 1 P1.B_CodeID,P1.B_ID"
            strSQL = strSQL & " From ("
            strSQL = strSQL & " Select Max(B_CodeID) as B_CodeID,B_ID From " & m_DraftMainTable
            strSQL = strSQL & " Where B_CodeID < '" & Trim(m_CodeID) & "' AND B_BID='" & mvarm_BID & "'" ' And B_UserName='" & Trim(m_typeSInformation.m_strUserName) & "'"
            strSQL = strSQL & " " & A_BLFilterSQL_Bill
            
            strSQL = MontageCreateOperator8Table(strSQL, m_DraftMainTable)
            
            strSQL = strSQL & " Group By B_ID"
            strSQL = strSQL & " Union"
            strSQL = strSQL & " Select Max(B_CodeID) as B_CodeID,B_ID From " & m_MainTable
            strSQL = strSQL & " Where B_CodeID < '" & Trim(m_CodeID) & "' AND B_BID='" & mvarm_BID & "'" ' And B_UserName='" & Trim(m_typeSInformation.m_strUserName) & "'"
            strSQL = strSQL & " " & A_BLFilterSQL_Bill
            
            strSQL = MontageCreateOperator8Table(strSQL, m_MainTable)
            
            strSQL = strSQL & " Group By B_ID"
            strSQL = strSQL & " ) as P1"
            strSQL = strSQL & " Order By P1.B_CodeID DESC"
        Case 3  '----下一
            strSQL = "Select Top 1 P1.B_CodeID,P1.B_ID"
            strSQL = strSQL & " From ("
            strSQL = strSQL & " Select Min(B_CodeID) as B_CodeID,B_ID From " & m_DraftMainTable
            strSQL = strSQL & " Where B_CodeID > '" & Trim(m_CodeID) & "' AND B_BID='" & mvarm_BID & "'" ' And B_UserName='" & Trim(m_typeSInformation.m_strUserName) & "'"
            strSQL = strSQL & " " & A_BLFilterSQL_Bill
            
            strSQL = MontageCreateOperator8Table(strSQL, m_DraftMainTable)
            
            strSQL = strSQL & " Group By B_ID"
            strSQL = strSQL & " Union"
            strSQL = strSQL & " Select Min(B_CodeID) as B_CodeID,B_ID From " & m_MainTable
            strSQL = strSQL & " Where B_CodeID > '" & Trim(m_CodeID) & "' AND B_BID='" & mvarm_BID & "'" ' And B_UserName='" & Trim(m_typeSInformation.m_strUserName) & "'"
            strSQL = strSQL & " " & A_BLFilterSQL_Bill
            
            strSQL = MontageCreateOperator8Table(strSQL, m_MainTable)
            
            strSQL = strSQL & " Group By B_ID"
            strSQL = strSQL & " ) as P1"
            strSQL = strSQL & " Order By P1.B_CodeID"
        Case 4  '----最后
            strSQL = "Select Top 1 P1.B_CodeID,P1.B_ID"
            strSQL = strSQL & " From ("
            strSQL = strSQL & " Select Max(B_CodeID) as B_CodeID,B_ID From " & m_DraftMainTable & " Where B_BID='" & mvarm_BID & "'" ' And B_UserName='" & Trim(m_typeSInformation.m_strUserName) & "'"
            strSQL = strSQL & " " & A_BLFilterSQL_Bill
            
            strSQL = MontageCreateOperator8Table(strSQL, m_DraftMainTable)
            
            strSQL = strSQL & " Group By B_ID"
            strSQL = strSQL & " Union"
            strSQL = strSQL & " Select Max(B_CodeID) as B_CodeID,B_ID From " & m_MainTable & " Where B_BID='" & mvarm_BID & "'" ' And B_UserName='" & Trim(m_typeSInformation.m_strUserName) & "'"
            strSQL = strSQL & " " & A_BLFilterSQL_Bill
            
            strSQL = MontageCreateOperator8Table(strSQL, m_MainTable)
            
            strSQL = strSQL & " Group By B_ID"
            strSQL = strSQL & " ) as P1"
            strSQL = strSQL & " Order By P1.B_CodeID DESC"
    End Select
    
    Debug.Print strSQL
    
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockBatchOptimistic
    If rs.EOF = False Then
        GetOtherFrame = rs("B_ID")
    Else
        GetOtherFrame = 0
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
IFERR:
    GetOtherFrame = 0
    Exit Function
End Function
'----以上为单据移动部分
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'----设置表单数据
Private Sub SetFielsList()
    On Error Resume Next
    Dim rs As New RecordSet
    Dim i As Integer
    Dim strSQL As String

    'cnn.InitializeConnection

    '----取得表单字段属性
    strSQL = "Select * From G_BLFieldUse Where B_ObjectID='" & Trim(mvarObjectID) & "'"
    'Debug.Print strSQL
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    Set rsBLField = New RecordSet
    For i = 0 To rs.Fields.Count - 1
        rsBLField.Fields.Append rs.Fields(i).name, rs.Fields(i).Type, rs.Fields(i).DefinedSize, rs.Fields(i).Attributes
        Debug.Print rs.Fields(i).name
    Next
    rsBLField.Open
    Do While Not rs.EOF
        rsBLField.AddNew
        For i = 0 To rs.Fields.Count - 1
            If Not IsNull(rs(i).Value) Then
                rsBLField.Fields(i).Value = rs.Fields(i).Value
            End If
        Next
        rsBLField.Update
        rs.movenext
    Loop

    rs.Close
    Set rs = Nothing
End Sub

'----取得记录编号
Public Function GetFrameCodeDetail() As String
    On Error Resume Next
    Dim strTmpBH As String
    Dim strTmpBHLast As String
    Dim strTmpMonth As String
    Dim strTmpDay As String
    Dim mstrSQL As String
    Dim rs As New RecordSet
    Dim gdateSystemDat As Date

    gdateSystemDat = Now

    Set rs = New RecordSet
    strTmpMonth = Trim(Month(gdateSystemDat))
    If Len(Trim(strTmpMonth)) = 1 Then
        strTmpMonth = "0" & strTmpMonth
    End If
    strTmpDay = Trim(Day(gdateSystemDat))
    If Len(Trim(strTmpDay)) = 1 Then
        strTmpDay = "0" & strTmpDay
    End If

    strTmpBH = Trim(mvarm_BID) & Trim(Year(gdateSystemDat)) & strTmpMonth & strTmpDay

    mstrSQL = "Select CASE WHEN ISNULL(P1.B_CodeID,0)>ISNULL(P2.B_CodeID,0) THEN P1.B_CodeID"
    mstrSQL = mstrSQL & " Else P2.B_CodeID End as B_PCodeID"
    mstrSQL = mstrSQL & " From (Select Max(B_CodeID) as B_CodeID"
    mstrSQL = mstrSQL & " From " & m_DraftMainTable
    mstrSQL = mstrSQL & " Where B_CodeID Like '" & Trim(strTmpBH) & "%') as P1,"

    mstrSQL = mstrSQL & " (Select Max(B_CodeID) as B_CodeID"
    mstrSQL = mstrSQL & " From " & m_MainTable
    mstrSQL = mstrSQL & " Where B_CodeID Like '" & Trim(strTmpBH) & "%') as P2"

    rs.Open mstrSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    If IsNull(rs(0)) Then
        '如果没有以前的记录
        strTmpBH = strTmpBH & "0001"
        GetFrameCodeDetail = strTmpBH
    Else
        strTmpBHLast = Trim(str(Val(Mid(Trim(rs(0)), 12, 4)) + 1))
        If Len(Trim(strTmpBHLast)) < 4 Then
            strTmpBH = strTmpBH & String(4 - Len(Trim(strTmpBHLast)), "0") & strTmpBHLast
        Else
            strTmpBH = strTmpBH & strTmpBHLast
        End If
        GetFrameCodeDetail = strTmpBH
    End If
    rs.Close
    Set rs = Nothing
End Function

'----设置显示方式
Private Sub UpdateFrameType()
    On Error Resume Next
    '显示输入项类型
    Dim i As Integer

    intIsClosed = Adodc1.RecordSet("B_Closed")
    A_IsVoid = IIf(IsNull(Adodc1.RecordSet("B_IsVoid")), 0, Adodc1.RecordSet("B_IsVoid"))
    

    Select Case intIsClosed
        Case 0  '新增
            ActiveBar21.Bands("P操作").Tools("新增").Enabled = True
            ActiveBar21.Bands("P操作").Tools("删除").Enabled = True
            ActiveBar21.Bands("P操作").Tools("修改").Enabled = False
            ActiveBar21.Bands("Band1").Tools("打印菜单").Enabled = False
            ActiveBar21.Bands("Band1").Tools("登帐标记").Visible = False
            
            ActiveBar21.Bands("Band1").Tools("结转标记").Visible = False
            

            ActiveBar21.Parent.UCButton1(0).Enabled = True
            ActiveBar21.Parent.UCButton1(1).Enabled = True
            
            TDBGrid1.AllowUpdate = True
        Case 1  '已登帐
            ActiveBar21.Bands("P操作").Tools("新增").Enabled = True
            ActiveBar21.Bands("P操作").Tools("删除").Enabled = False
            ActiveBar21.Bands("P操作").Tools("修改").Enabled = True
            ActiveBar21.Bands("Band1").Tools("打印菜单").Enabled = True
            ActiveBar21.Bands("Band1").Tools("登帐标记").Visible = True
            
            
            ActiveBar21.Parent.UCButton1(0).Enabled = False
            ActiveBar21.Parent.UCButton1(1).Enabled = False
            
            TDBGrid1.AllowUpdate = False
            
        Case 2  '修改
            ActiveBar21.Bands("P操作").Tools("新增").Enabled = True
            ActiveBar21.Bands("P操作").Tools("删除").Enabled = True
            ActiveBar21.Bands("P操作").Tools("修改").Enabled = False
            ActiveBar21.Bands("Band1").Tools("打印菜单").Enabled = False
            ActiveBar21.Bands("Band1").Tools("登帐标记").Visible = False
            
            
            ActiveBar21.Parent.UCButton1(0).Enabled = True
            ActiveBar21.Parent.UCButton1(1).Enabled = False
            
            TDBGrid1.AllowUpdate = True
            
        Case 3  '已结转
            ActiveBar21.Bands("P操作").Tools("新增").Enabled = True
            ActiveBar21.Bands("P操作").Tools("删除").Enabled = False
            ActiveBar21.Bands("P操作").Tools("修改").Enabled = False
            ActiveBar21.Bands("Band1").Tools("打印菜单").Enabled = True
            ActiveBar21.Bands("Band1").Tools("登帐标记").Visible = False

            ActiveBar21.Bands("Band1").Tools("结转标记").Visible = True

            
            ActiveBar21.Parent.UCButton1(0).Enabled = False
            ActiveBar21.Parent.UCButton1(1).Enabled = False
            
            TDBGrid1.AllowUpdate = True
    End Select
    
    '设置作废按钮和作废标记的显示与否
    SetDisplay_Void
    
    
    ActiveBar21.RecalcLayout
    RaiseEvent OnUpdateFrameType
End Sub

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    On Error Resume Next

    Select Case Tool.name
        Case "第一单"
            MoveFirst
        Case "前单"
            MovePrevious
        Case "后单"
            movenext
        Case "最后单"
            movelast
        Case "新增"
            AddNewFrame
        Case "复制"
            CopyFrameEx
        Case "删除"
            DeleteNowFrame
            
        Case "增加明细"
            AddNewItem
        Case "删除明细"
            DeleteItem
        Case "删除所有"
            DeleteItemAll
        Case "修改", "反登帐"
            UnCheckIn
        Case "刷新"
            Refresh
        Case "业务"
            BLSearch
            
        Case "打印"
            If JudgeBeforePrint = False Then
                Exit Sub
            End If
            PrintBillReport
            
            WriteLogPrint mvarm_ID
        Case "打印预览", "打印菜单"
            If JudgeBeforePrint = False Then
                Exit Sub
            End If
            PrintBillReportPreview
            
        Case "导入发货汇总"
            frm1.ImportData
            
        Case "作废"
            SetToVoid
            
        Case "保存样式"
            
        Case Else
            
    End Select
    'ActiveBar21.Parent.SetFocus
End Sub

Private Sub PrintBillReport()
    clsReport1.KeyID = m_ID
    clsReport1.PrintDirect
End Sub

Private Sub PrintBillReportPreview()
    clsReport1.Refresh  '获取打印报表的存储过程语句
    
    Dim rs As New RecordSet
    Dim szReportObjectID As String
    Dim strSQL As String
    
    szReportObjectID = clsReport1.ReportObjectID
    strSQL = clsReport1.sql & " " & m_ID
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    PrintPreview szReportObjectID, rs
    
    rs.Close
    Set rs = Nothing
End Sub


Private Function GetExeSQL(ByVal vObjectID As String) As String
    Dim rs As RecordSet
    Dim strSQL As String
    Set rs = New RecordSet
    strSQL = "Select * From G_Report Where b_ObjectID='" & vObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        GetExeSQL = rs!B_ExeSQL
    Else
        GetExeSQL = ""
    End If
    rs.Close
    Set rs = Nothing
End Function


'根据选择：
'打印发货单-无缸号   净重
'打印发货单-无缸号   毛重
'打印发货单-分缸号   净重
'打印发货单-分缸号   毛重
'来获取当前需要采用哪个报表模板
Private Function GetFHDObjectID(ByVal vType01 As String, ByVal vType02 As String) As String
    Dim rs As New RecordSet
    Dim strSQL As String
    strSQL = "Select * From G_Config_FHD where B_Type01='" & vType01 & "' And B_Type02='" & vType02 & "'"

    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        GetFHDObjectID = ""
    Else
        GetFHDObjectID = rs!B_ObjectID
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Sub Refresh()
    On Error Resume Next
    TDBGrid1.Refresh
    Dim o As Object
    For Each o In ActiveBar21.Parent.Controls
        If Mid(o.name, 1, 2) = "B_" Then
            o.Refresh
            Set o = Nothing
        End If
    Next
End Sub

Private Sub BLSearch()
    Dim frm1 As New frmBLSearch
    With frm1
        Set .frm1 = ActiveBar21.Parent
        .m_ObjectID = mvarObjectID
        .Show
        .SetFocus
    End With
End Sub

Private Sub UnCheckIn()

    '判断是否有修改的权限
    If Gm.PI.JudgeUpdateBill(Me.ObjectID, mvarm_ID) = False Then
        Exit Sub
    End If
    
    
    On Error Resume Next
    Dim strSQL As String
    
    Select Case mvarObjectID
        Case "120007"
            If JudgeCanDel = False Then
                Exit Sub
            End If
    End Select
    
    
    '检测是否可修改
    If JudgeCanModify = False Then
        Exit Sub
    End If
    
    
    strSQL = "Update " & m_MainTable & " Set B_Closed=2 Where B_ID=" & m_ID
    Gm.cnnTool.cnn.Execute strSQL
    Adodc1.RecordSet.requery
    TDBGrid1.DataSource = Adodc2
    TDBGrid1.ReBind
    intIsClosed = 2
    UpdateFrameType
End Sub

Private Function CheckInputFrame() As Boolean
    Dim OK As Boolean
    OK = True
    RaiseEvent OnCheckInputFrame(OK)
    CheckInputFrame = OK
End Function

Private Function CheckInputItem() As Boolean
    If Adodc2.RecordSet.RecordCount > 0 Then
        Adodc2.RecordSet.MoveFirst
    Else
        MsgBox "无明细数据!", vbInformation, "明细"
        CheckInputItem = False
        Exit Function
    End If

    Do While Not Adodc2.RecordSet.EOF
        If CheckOneInputItem = False Then
            MsgBox "明细数据输入不正确!", vbInformation, "输入"
            CheckInputItem = False
            Exit Function
        End If

        Adodc2.RecordSet.movenext
    Loop
    Adodc2.RecordSet.MoveFirst
    CheckInputItem = True
End Function

'----判定单条明细表输入的项目
Private Function CheckOneInputItem() As Boolean
    Dim OK As Boolean
    OK = True
    
    
'    If Len(Trim(TDBGrid1.Columns(0).Text)) < 1 Then
'        CheckOneInputItem = False
'        Exit Function
'    End If

    RaiseEvent OnCheckOnInputItem(OK)
    CheckOneInputItem = OK
End Function


Private Sub BLCheckIn()
    On Error GoTo IFERR
    
    '某些列不可为空
'    If JudgeNotNull = False Then
'        Exit Sub
'    End If
    
    If MsgBox("是否要登帐?", vbInformation + vbOKCancel + vbDefaultButton2, "登帐") = vbCancel Then
        Exit Sub
    End If

    '删除第一个字段为空的明细记录
    ClearInputItemDetail
    SaveFrame

    If CheckInputFrame = False Then
        Exit Sub
    End If

    If CheckInputItem = False Then
        Exit Sub
    End If
    
    
    
    Dim clsCheckIn1 As New clsCheckIn
    Set clsCheckIn1 = New clsCheckIn
    With clsCheckIn1
        .m_ID = m_ID
        .ObjectID = mvarObjectID
        .CheckIn

        If .OK = True Then
            MsgBox "已成功登帐!", vbInformation, "保存"
            m_ID = .m_ID
        Else
            Exit Sub
        End If
    End With

    boolIsSave = False
    boolIsLoadType = False
    boolIsDraft = False
    OpenFrame
    
    Exit Sub
IFERR:
    MsgBox Err.Description, vbExclamation, "保存"
    Exit Sub
End Sub


Private Sub BLCheckInEx()
    On Error GoTo IFERR
    If MsgBox("是否要登帐?", vbInformation + vbOKCancel + vbDefaultButton2, "登帐") = vbCancel Then
        Exit Sub
    End If

    'ClearInputItemDetail
    SaveFrame

    If CheckInputFrame = False Then
        Exit Sub
    End If

    If CheckInputItem = False Then
        Exit Sub
    End If
    
    Dim clsCheckIn1 As New clsCheckIn
    Set clsCheckIn1 = New clsCheckIn
    With clsCheckIn1
        .m_ID = m_ID
        .ObjectID = mvarObjectID
        .CheckInEx

        If .OK = True Then
            MsgBox "已成功登帐!", vbInformation, "保存"
            m_ID = .m_ID
        Else
            Exit Sub
        End If
    End With

    boolIsSave = False
    boolIsLoadType = False
    boolIsDraft = False
    OpenFrame
    
    Exit Sub
IFERR:
    MsgBox Err.Description, vbExclamation, "保存"
    Exit Sub
End Sub

'----单据登帐
Public Sub BillCheckIn()
    If intIsClosed = 0 Then
        BLCheckIn
    Else
        SaveFrame
        Dim strSQL As String
        strSQL = "Update " & m_MainTable & " Set B_Closed=1 Where B_ID=" & m_ID
        Debug.Print strSQL
        Gm.cnnTool.cnn.Execute strSQL


        'Adodc1.RecordSet.Fields!B_Closed = 1
        On Error Resume Next
        Adodc1.RecordSet.Update
        Adodc1.RecordSet.requery
        Adodc2.RecordSet.requery
        intIsClosed = 1
        UpdateFrameType
        TDBGrid1.DataSource = Adodc2
        
    End If
    
    
    '财务的单据沿用仓库的单据保存时候会造成B_ObjectID不一样,在此修正
    JudgeBLConvert
End Sub

'当前单据是沿用仓库的单据的话  要修改B_ObjectID
Private Sub JudgeBLConvert()
    '财务的单据沿用仓库的单据保存时候会造成B_ObjectID不一样,在此修正
    '=========================
    Dim rs1 As RecordSet
    Dim strSQL As String
    
    Set rs1 = New RecordSet
    strSQL = "Select * From G_BLConvert where B_oObjectID='" & mvarObjectID & "'"
    Debug.Print strSQL
    rs1.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    If rs1.RecordCount > 0 Then
        strSQL = "Update " & m_MainTable & " Set B_ObjectID='" & rs1("B_tObjectID") & "' Where B_ID=" & m_ID
        Debug.Print strSQL
        Gm.cnnTool.cnn.Execute strSQL
    End If
    rs1.Close
    Set rs1 = Nothing
    '=========================
    
End Sub


Public Sub BillCheckInEx()
    If intIsClosed = 0 Then
        BLCheckInEx
    Else
        SaveFrame
        Dim strSQL As String
        strSQL = "Update " & m_MainTable & " Set B_Closed=1 Where B_ID=" & m_ID
        Gm.cnnTool.cnn.Execute strSQL
        Adodc1.RecordSet.requery
        Adodc2.RecordSet.requery
        intIsClosed = 1
        UpdateFrameType
        TDBGrid1.DataSource = Adodc2
    End If
End Sub

Private Sub clsCheckIn1_OnCheckIn(oCnn As ADODB.Connection)
    RaiseEvent OnBillCheckIn(oCnn)
End Sub

Private Sub ActiveBar21_ToolComboClose(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    On Error Resume Next
    Dim aReport

    aReport = Split(ActiveBar21.Bands("Band1").Tools("报表").Text, ",")
    
    clsReport1.ReportObjectID = aReport(0)
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set ScriptControl1 = Nothing
    A_rsBillField.Close
    Set A_rsBillField = Nothing
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo IFERR


    m_Mark = TDBGrid1.Bookmark
    ADORequery ColIndex
    TDBGrid1.Bookmark = m_Mark
    

    '在2017年3月8日 16:59:39测试时弃用
    If intIsClosed = 2 Then
        If ColIndex = 0 Or ColIndex = 1 Then
            Adodc2.RecordSet.requery
            TDBGrid1.PostMsg 81
        End If

        '在2017年5月4日使用下面的方式
        '单据中其他列上有下拉网格数据时，也会自动刷新
        '如经测试后有问题，将下面两行代码注释
        '仍然采用上面的代码块即可
'        ADORequery ColIndex
'        TDBGrid1.PostMsg 81
    End If
    
    
    boolIsSave = True
    Exit Sub
IFERR:
    Exit Sub
End Sub



Private Function InputFilterBLEx(ByRef vTDBGrid1 As TDBGrid, ByVal vColIndex As Integer, ByVal vStr As String) As Boolean
    If vColIndex < 0 Then
        InputFilterBLEx = True
        Exit Function
    End If
    
    A_rsIPF.Filter = ""
    If A_rsIPF.RecordCount <= 0 Then
        InputFilterBLEx = True
        Exit Function
    End If
    
    
    A_rsIPF.Filter = ""
    A_rsIPF.Filter = " B_FieldName='" & vTDBGrid1.Columns(vColIndex).DataField & "'"
    If A_rsIPF.RecordCount <= 0 Then
        InputFilterBLEx = True
        Exit Function
    End If
    
    Dim i As Long
    InputFilterBLEx = True
    If Len(Trim(vStr)) > 0 Then
        For i = 0 To Len(Trim(vStr)) - 1
            If CheckFilterOne(Mid$(Trim(vStr), i, 1)) = False Then
                InputFilterBLEx = False
                Exit Function
            End If
        Next
    End If
End Function

Private Function CheckFilterOne(ByVal vStr As String) As Boolean
    CheckFilterOne = False
    If IIf(IsNull(A_rsIPF!B_YSZ), 0, A_rsIPF!B_YSZ) = 1 Then
        If IsNumeric(vStr) = True Then
            CheckFilterOne = True
            Exit Function
        End If
    End If
    
    
    If IIf(IsNull(A_rsIPF!B_YYW), 0, A_rsIPF!B_YYW) = 1 Then
        If Asc(vStr) >= 65 And Asc(vStr) <= 90 Then
            CheckFilterOne = True
            Exit Function
        End If
        
        If Asc(vStr) >= 97 And Asc(vStr) <= 122 Then
            CheckFilterOne = True
            Exit Function
        End If
    End If
    
    
    Dim m_Array
    Dim i As Long
    Dim m_szToASC As String
    Dim m_lToASC As Long
    
    If Len(IIf(IsNull(A_rsIPF!B_AFH), "", A_rsIPF!B_AFH)) > 0 Then
        CheckFilterOne = False
        m_Array = Split(A_rsIPF!B_AFH, ",")
        For i = 0 To UBound(m_Array)
            If Left$(m_Array(i), 1) = "(" Then
                m_szToASC = Replace$(m_Array(i), "(", "")
                m_szToASC = Replace$(m_szToASC, ")", "")
                m_lToASC = Val(Trim$(m_szToASC))
                
                
                If Asc(vStr) = m_lToASC Then
                    CheckFilterOne = True
                    Exit Function
                End If
            Else
                If vStr = m_Array(i) Then
                    CheckFilterOne = True
                    Exit Function
                End If
            End If
        Next
    End If
End Function

'在网格中编辑时候过滤掉非法的字符
'返回值：
'返回TRUE表示过滤后结果为合法，表示当前字符可以被输入
'相反则不可以被输入。返回FALSE
Private Function InputFilterBL(ByRef vTDBGrid1 As TDBGrid, ByVal vColIndex As Integer, ByVal vKeyAscii As Long) As Boolean
    If A_rsIPF.RecordCount <= 0 Then
        InputFilterBL = True
        Exit Function
    End If
    
    A_rsIPF.Filter = ""
    A_rsIPF.Filter = " B_FieldName='" & vTDBGrid1.Columns(vColIndex).DataField & "'"
    If A_rsIPF.RecordCount <= 0 Then
        InputFilterBL = True
        Exit Function
    End If
    
    
    '先初始化为FALSE，经过下面任何一个检验通过的话则设置为TRUE
    InputFilterBL = False
    If IIf(IsNull(A_rsIPF!B_YSZ), 0, A_rsIPF!B_YSZ) = 1 Then
        If IsNumeric(Chr(vKeyAscii)) = True Then
            InputFilterBL = True
            Exit Function
        End If
    End If
    
    
    
    If IIf(IsNull(A_rsIPF!B_YYW), 0, A_rsIPF!B_YYW) = 1 Then
        If vKeyAscii >= 65 And vKeyAscii <= 90 Then
            InputFilterBL = True
            Exit Function
        End If
        
        If vKeyAscii >= 97 And vKeyAscii <= 122 Then
            InputFilterBL = True
            Exit Function
        End If
    End If
    
    
    Dim m_Array
    Dim i As Long
    Dim m_szToASC As String
    Dim m_lToASC As Long
    
    '在字段B_AFH中使用圆括号中间填写数字的格式为过滤ASC码，本ASC码转换成的字符为合法
    If Len(IIf(IsNull(A_rsIPF!B_AFH), "", A_rsIPF!B_AFH)) > 0 Then
        InputFilterBL = False
        m_Array = Split(A_rsIPF!B_AFH, ",")
        For i = 0 To UBound(m_Array)
            If Left$(m_Array(i), 1) = "(" Then
                m_szToASC = Replace$(m_Array(i), "(", "")
                m_szToASC = Replace$(m_szToASC, ")", "")
                m_lToASC = Val(Trim$(m_szToASC))
                
                
                If vKeyAscii = m_lToASC Then
                    InputFilterBL = True
                    Exit Function
                End If
            Else
                If Chr(vKeyAscii) = m_Array(i) Then
                    InputFilterBL = True
                    Exit Function
                End If
            End If
        Next
    End If
    
End Function

Private Sub TDBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If TDBGrid1.Columns(ColIndex) = OldValue Then
        Exit Sub
    Else
        '----当金额数量进行修改时
        If TDBGrid1.Col > 2 Then
            UseFormulaCount
        End If
    End If
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If InputFilterBL(TDBGrid1, TDBGrid1.Col, KeyAscii) = False Then
        MsgBox "当前字" & Chr(KeyAscii) & "符非法,您必须删除该字符！", vbOKOnly + vbInformation, "提示"
        TDBGrid1.SetFocus
    End If
    
    
    '----当回车时
    '----进行回车跳列
    '----锁定状态或无显示状态
    Dim i As Integer
    If KeyAscii = 13 Then
        For i = TDBGrid1.Col To TDBGrid1.Columns.Count - 1
            If TDBGrid1.Columns(i).Visible = True And TDBGrid1.Columns(i).Locked = False Then
                TDBGrid1.Col = i
                Exit For
            End If
        Next
        If TDBGrid1.Col = m_MaxIndex And CheckOneInputItem = True Then

            AddNewItem
            
            If mvarObjectID = "120007" Then
'                Select Case frm1.Controls("B_ColIndex").Text
'                    Case "0"
'                        TDBGrid1.Col = 0
'                    Case "1"
'                        TDBGrid1.Col = 1
'                    Case Else
'                        TDBGrid1.Col = 0
'                End Select

                TDBGrid1.Col = Val(Trim(frm1.Controls("B_ColIndex").Text))
            Else
                TDBGrid1.Col = 0
            End If
        End If
    End If
    
    
    
End Sub



Private Sub TDBGrid1_PostEvent(ByVal MsgId As Integer)
    On Error Resume Next
    Dim m_Col As Integer
    Select Case MsgId
        Case 81 '表单中修改
            m_Mark = TDBGrid1.Bookmark
            m_Col = TDBGrid1.Col
            Adodc2.RecordSet.requery
            TDBGrid1.Bookmark = m_Mark
            TDBGrid1.Col = m_Col
    End Select
End Sub



'判断当前入库单是否可以删除
'已经出过库的东西 不可删除
Private Function JudgeCanDel() As Boolean
    Dim rs As RecordSet
    Dim strSQL As String
    
    
    JudgeCanDel = True
    Set rs = New RecordSet
    strSQL = "Select G_BillDetailCP.B_ItemID,G_BillDetailCP.B_ID,"
    strSQL = strSQL & " G_BillDetailCP.B_BelongItemID"
    strSQL = strSQL & " From G_BillCP,G_BillDetailCP"
    strSQL = strSQL & " Where G_BillCP.B_ID=G_BillDetailCP.B_ID"
    strSQL = strSQL & " And G_BillCP.B_ObjectID In ('120008','120009')"
    strSQL = strSQL & " And B_BelongItemID In"
    strSQL = strSQL & " ("
    strSQL = strSQL & "     Select B_ItemID"
    strSQL = strSQL & "     From G_BillDetailCP"
    strSQL = strSQL & "     Where B_ID=" & m_ID
    strSQL = strSQL & " )"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        JudgeCanDel = False
        MsgBox "本入库单部分或者全部货物已经出库，不可操作！", vbOKOnly + vbInformation, "提示"
    End If
    
    rs.Close
    Set rs = Nothing

End Function

'判断当前单据的对象编号是否存在于规则之中
'存在的话返回TRUE，不存在的话返回FALSE
'通过传址，不仅判断当前单据编号是否存在于某规则内
'同时还定位到其所在的规则。在引用本函数的上级函数中直接使用记录中的规则
Private Function JudgeCurObjectIDInRules(ByVal vRs As RecordSet) As Boolean
    On Error GoTo IFERR
    Dim rs As New RecordSet
    Dim szObjectIDs As String
    
    
    JudgeCurObjectIDInRules = False
    If vRs.RecordCount <= 0 Then
        JudgeCurObjectIDInRules = False
        Exit Function
    End If
    
    szObjectIDs = ""
    vRs.MoveFirst
    Do While Not vRs.EOF
        szObjectIDs = vRs!B_ObjectID
        StringToRecordset szObjectIDs, ",", rs
        rs.Filter = " B_Field0='" & mvarObjectID & "'"
        If rs.RecordCount > 0 Then
            JudgeCurObjectIDInRules = True
            Exit Function
        End If
        vRs.movenext
    Loop


    Exit Function
IFERR:
    Dim szErr As String
    szErr = Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
    Exit Function
End Function


'返回TRUE表示当前可修改，FALSE表示不可修改
Private Function JudgeCanModify() As Boolean
    On Error GoTo IFERR
    Dim strSQL As String
    Dim rs As RecordSet
    Dim rsAdmin As RecordSet
    Dim szCNAdmin As String
    
    Dim szDateChecked As String
    Dim szNow As String
    Dim szDT As String
    Dim dDateDiff As Double
    
    Dim bObjectIDExists As Boolean
    
    Dim szTip As String
    
    JudgeCanModify = True
    
    
    '获取规则配置表
    '============================
    Dim rsConfig As RecordSet
    Set rsConfig = New RecordSet
    strSQL = "Select * From G_Config_CPFH Where B_GroupName='成品发货单修改规则'"
    Debug.Print strSQL
    rsConfig.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rsConfig.RecordCount <= 0 Then
        JudgeCanModify = True
        rsConfig.Close
        Set rsConfig = Nothing
        Exit Function
    End If
    '============================
    
    
    
    '如果正式表中没有数据则返回TRUE并且退出
    Set rs = New RecordSet
    strSQL = "Select * From " & m_MainTable & " Where B_ID=" & mvarm_ID
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        JudgeCanModify = True
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    szDateChecked = Format(rs!B_Date, "YYYY-MM-DD HH:MM:SS")
    rs.Close
    Set rs = Nothing
    
    
    
    '判断当前单据对象在应用于的对象范围之外的话则返回TRUE并且直接退出
    '================================================================
    If JudgeCurObjectIDInRules(rsConfig) = False Then
        JudgeCanModify = True
        Exit Function
    End If
    '================================================================
    
    
    
    
    'rsConfig.MoveFirst
    'Do While Not rsConfig.EOF
        szCNAdmin = rsConfig!B_CN
        StringToRecordset szCNAdmin, ",", rs
        
        '设置了管理员权限的计算机名的话，则先验证本机是否管理员计算机
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                If UCase(rs!B_Field0) = UCase(Gm.SysID.ComputerName) Then
                    JudgeCanModify = True
                    rs.Close
                    Set rs = Nothing
                    
                    rsConfig.Close
                    Set rsConfig = Nothing
                    Exit Function
                End If
                rs.movenext
            Loop
        End If
        
    
        szNow = GetSQLTime
        
        '自然时间的当天可修改
        If rsConfig!B_DT = 0 Then
            If Format(szDateChecked, "YYYY-MM-DD") = Format(szNow, "YYYY-MM-DD") Then
                JudgeCanModify = True
                Exit Function
            Else
                szTip = "修改验证规则为：当天【本单据日期为" & Format(szDateChecked, "YYYY-MM-DD") & "您只能在当天修改数据】" & vbNewLine
                szTip = szTip & "当前验证结果：失败！" & vbNewLine
                szTip = szTip & "没有修改权限"
                MsgBox szTip, vbOKOnly + vbInformation, "提示"
            End If
        End If
        
        '登帐后的一段时间内可修改
        If rsConfig!B_DT = 1 Then
            dDateDiff = Val(DateDiff("h", szDateChecked, szNow))
            If dDateDiff <= rsConfig!B_DateDiff Then
                JudgeCanModify = True
                Exit Function
            Else
                szTip = "修改验证规则为：规定时间段内【登帐后" & rsConfig!B_DateDiff & "小时内可修改】" & vbNewLine
                szTip = szTip & "当前验证结果：失败！" & vbNewLine
                szTip = szTip & "没有修改权限"
                MsgBox szTip, vbOKOnly + vbInformation, "提示"
            End If
        End If
        
        
        '登帐后立即不可修改
        If rsConfig!B_DT = 2 Then
            szTip = "修改验证规则为：登帐后不可修改！" & vbNewLine
            MsgBox szTip, vbOKOnly + vbInformation, "提示"
        End If

        
        'rsConfig.MoveNext
    'Loop
    
    
    rsConfig.Close
    Set rs = Nothing
    
    Exit Function
IFERR:
    JudgeCanModify = True
End Function



'将带有分割符号的字符串转移到记录集中
Private Sub StringToRecordset(ByVal m_str As String, ByVal m_Filter As String, ByRef rs As RecordSet)
    Dim i As Long
    Dim m_Array
       
       
    m_Array = Split(m_str, m_Filter)
       
    Set rs = New RecordSet
    rs.Fields.Append "B_Field0", adVarChar, 100, adFldIsNullable
    rs.Open
       
       
    For i = 0 To UBound(m_Array)
        If Len(Trim(m_Array(i))) > 0 Then
           
            rs.AddNew
            rs("B_Field0") = m_Array(i)
            rs.Update
        End If
    Next
       
       
    rs.MoveFirst
       
End Sub

Private Function GetSQLTime() As String
    Dim strSQL As String
    Dim rs As RecordSet
    
    Set rs = New RecordSet
    strSQL = "Select GetDate() as B_Date"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    GetSQLTime = Format(rs!B_Date, "YYYY-MM-DD HH:MM:SS")
    rs.Close
    Set rs = Nothing
    
End Function

'设置审核状态
Private Sub SetDisplay_Audit()
    
End Sub

'设置当前作废按钮   和  作废标记的显示样式
'B_IsVoid=1:表示当前单据已经作废，对应的 作废按钮不可用，作废标记要显示
'B_IsVoid=0：表示当前单据没有作废，这种情况下还要看当前单据的状态如果是登帐状态则作废按钮可用，否则不可用
Private Sub SetDisplay_Void()
    Dim lbtVoid As Long   '作废按钮是否可用
    Dim lbjVoid As Long   '作废标记是否可见
    

    
    lbjVoid = IIf(IsNull(Adodc1.RecordSet!B_IsVoid), 0, Adodc1.RecordSet!B_IsVoid)
    Select Case Adodc1.RecordSet!B_Closed
        Case 1  '登帐状态
            lbtVoid = 1
        Case Else
            lbtVoid = 0
    End Select
    
    
    '1.登帐   作废
    If lbtVoid = 1 And lbjVoid = 1 Then
        ActiveBar21.Bands("Band1").Tools("作废").Enabled = False
        ActiveBar21.Bands("Band1").Tools("作废标记").Visible = True
    End If
    
    '2.登帐   未作废
    If lbtVoid = 1 And lbjVoid = 0 Then
        If W_SetVoidEnabled = True Then   '如果当前单据不允许作废
            ActiveBar21.Bands("Band1").Tools("作废").Enabled = True
        End If
        ActiveBar21.Bands("Band1").Tools("作废标记").Visible = False
    End If
    
    '3.草稿或修改   作废
    If lbtVoid <> 1 And lbjVoid = 1 Then
        ActiveBar21.Bands("Band1").Tools("作废").Enabled = False
        ActiveBar21.Bands("Band1").Tools("作废标记").Visible = False
    End If
    
    '4.草稿或修改   未作废
    If lbtVoid <> 1 And lbjVoid = 0 Then
        ActiveBar21.Bands("Band1").Tools("作废").Enabled = False
        ActiveBar21.Bands("Band1").Tools("作废标记").Visible = False
    End If
End Sub

'设置当前单据为作废状态
Private Sub SetToVoid()
    Dim strSQL As String

    If MsgBox("您确定要作废本单据么，该操作不可逆！", vbYesNo + vbExclamation + vbDefaultButton2, "警告") = vbNo Then
        Exit Sub
    End If
    
    Adodc1.RecordSet!B_IsVoid = 1
    Adodc1.RecordSet.Update
    
    Dim cls1 As New clsOPLogBill
    Dim rs As RecordSet
    Set rs = New RecordSet
    strSQL = "Select * From " & m_MainTable & " Where B_ID=" & mvarm_ID
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    cls1.WriteLog_SetVoid IIf(IsNull(rs!B_id), 0, rs!B_id), IIf(IsNull(rs!B_CodeID), "", rs!B_CodeID), mvarObjectID
    rs.Close
    Set rs = Nothing
    
    '作废后自动执行登帐操作
    BillCheckIn
End Sub

'获取作废功能的配置文件（判断当前计算机，当前子系统，当前单据是否可用该功能）
Private Sub InitSetVoid()
    Dim strSQL As String
    Dim i As Long
    Dim arr01
    Dim bExistsObjectID As Boolean
    
    bExistsObjectID = False
    W_SetVoidEnabled = True
    
    Set A_rsConfig_SetVoid = New RecordSet
    strSQL = "Select * from G_Config_SetVoid Where B_GroupName='不可用作废功能的单据'"
    Debug.Print strSQL
    A_rsConfig_SetVoid.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    
    If A_rsConfig_SetVoid.RecordCount <= 0 Then
        W_SetVoidEnabled = True
        Exit Sub
    End If
    
    '是管理员计算机的话则赋值TRUE并且退出
    If Len(IIf(IsNull(A_rsConfig_SetVoid!B_AdminCN), "", A_rsConfig_SetVoid!B_AdminCN)) > 0 Then
        arr01 = Split(A_rsConfig_SetVoid!B_AdminCN, ",")
        For i = 0 To UBound(arr01)
            If UCase(Gm.SysID.ComputerName) = UCase(arr01(i)) Then
                W_SetVoidEnabled = True
                Exit Sub
            End If
        Next
    End If
    
    '如果是管理员子系统的话则赋值TRUE并且退出
    If Len(IIf(IsNull(A_rsConfig_SetVoid!B_AdminSubSys), "", A_rsConfig_SetVoid!B_AdminSubSys)) > 0 Then
        arr01 = Split(A_rsConfig_SetVoid!B_AdminSubSys, ",")
        For i = 0 To UBound(arr01)
            
            If Gm.SysID.SubSystem = arr01(i) Then
                W_SetVoidEnabled = True
                Exit Sub
            End If
        Next
    End If
    
    
    If Len(IIf(IsNull(A_rsConfig_SetVoid!B_ObjectID), "", A_rsConfig_SetVoid!B_ObjectID)) > 0 Then
        arr01 = Split(A_rsConfig_SetVoid!B_ObjectID, ",")
        For i = 0 To UBound(arr01)
            If mvarObjectID = arr01(i) Then
                bExistsObjectID = True
            End If
        Next
    End If
    
    If bExistsObjectID = True Then
        W_SetVoidEnabled = False
    Else
        W_SetVoidEnabled = True
    End If
End Sub




'以下均为生成成品退仓回修单而制作的函数
'================================================================
'复制一个表中的数据（完全复制，一般为复制主表，之后再修改相关系统级别的参数）
'并且返回新创建的主表的主键字段的VALUE
Private Function BL_Auto_Copy_Bill(ByVal vDraftBill As String, ByVal vBill As String, ByVal vKeyField As String, ByVal vKeyValue As String) As Long
    Dim oRs As RecordSet
    Dim tRs As RecordSet
    Dim strSQL As String
    Dim i As Long
    
    Set oRs = New RecordSet
    Set tRs = New RecordSet
    strSQL = "Select * From " & vBill & " Where " & vKeyField & "='" & vKeyValue & "'"
    Debug.Print strSQL
    oRs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    strSQL = "Select * From " & vDraftBill & " where 1=0"
    Debug.Print strSQL
    tRs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    tRs.AddNew
    For i = 0 To oRs.Fields.Count - 1
        '非自增的字段才进行复制
        If oRs(i).name <> vKeyField Then
            tRs(Trim(oRs(i).name)) = oRs(i)
        End If
    Next
    tRs.Update
    BL_Auto_Copy_Bill = tRs(vKeyField)
    tRs.Close
    Set tRs = Nothing
    
End Function

'将明细表vTableName中vIDFieldName(B_ID)为vOID的明细数据全部复制为vTID的明细
Private Sub BL_Auto_Copy_Detail(ByVal vDraftDetail As String, ByVal vDetail As String, ByVal vIDFieldName As String, ByVal vItemIDFieldName As String, ByVal vOID As Long, ByVal vTID As Long)
    Dim strSQL As String
    Dim oRs As New RecordSet
    Dim tRs As New RecordSet
    Dim i As Long
    
    strSQL = "Select * From " & vDetail & " Where " & vIDFieldName & "='" & vOID & "'"
    Debug.Print strSQL
    oRs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If oRs.RecordCount <= 0 Then
        oRs.Close
        Set oRs = Nothing
        Exit Sub
    End If
    
    
    Do While Not oRs.EOF
        Set tRs = New RecordSet
        strSQL = "Select * From " & vDraftDetail & " Where 1=0"
        Debug.Print strSQL
        tRs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        tRs.AddNew
        For i = 0 To oRs.Fields.Count - 1
            If oRs(i).name <> vItemIDFieldName Then
                '当目标草稿明细表中有该字段的时候才进行赋值
                If Judge_Field_InRs(tRs, oRs(i).name) = True Then
                    tRs(Trim(oRs(i).name)) = oRs(i)
                End If
            End If
        Next
        tRs(vIDFieldName) = vTID
        tRs.Update
        tRs.Close
        Set tRs = Nothing
        
        oRs.movenext
    Loop
    
    oRs.Close
    Set oRs = Nothing
End Sub

'判断字段vFieldName在记录集中是否存在。存在返回TRUE，不存在返回FALSE
Private Function Judge_Field_InRs(ByRef vRs As RecordSet, ByVal vFieldName As String) As Boolean
    Dim i As Long
    Judge_Field_InRs = False
    For i = 0 To vRs.Fields.Count - 1
        If vRs(i).name = vFieldName Then
            Judge_Field_InRs = True
            Exit Function
        End If
    Next
End Function

'将草稿表vDraftBillName中的主键字段vKeyField=vKeyValue
'的记录的下列系统级别的字段：B_ObjectID,B_CodeID,B_BID修改为目标vTObjectID的参数
'并且修改单据日期为当前的日期
Private Sub BL_Auto_Copy_SaveTo(ByVal vDraftBillName As String, ByVal vKeyField As String, ByVal vKeyValue As Long, ByVal vTObjectID As String)
    Dim szDateTime As String
    Dim rs As New RecordSet
    Dim oRs As New RecordSet
    Dim strSQL As String
    
    strSQL = "Select * From " & vDraftBillName & " Where " & vKeyField & "='" & vKeyValue & "'"
    Debug.Print strSQL
    Set oRs = New RecordSet
    oRs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If oRs.RecordCount <= 0 Then
        oRs.Close
        Set oRs = Nothing
        Exit Sub
    End If
    
    
    
    szDateTime = Format(Now, "YYYY-MM-DD HH:MM:SS")
    strSQL = "Select * From G_BL Where B_ObjectID='" & vTObjectID & "'"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    oRs!B_Date = szDateTime
    oRs!B_ObjectID = rs!B_ObjectID
    oRs!B_BID = rs!B_BID
    oRs!B_CodeID = GetFrameCodeDetail_01(vTObjectID)
    oRs.Update
    oRs.Close
    Set oRs = Nothing
    
    rs.Close
    Set rs = Nothing
End Sub

'删除草稿表数据
Private Sub BL_Auto_Delete_Draft(ByVal vBill As String, ByVal vDetail As String, ByVal vKeyField As String, ByVal vKeyValue As Long)
    Dim strSQL As String
    strSQL = "Delete From " & vDetail & " Where " & vKeyField & "='" & vKeyValue & "'"
    Gm.cnnTool.cnn.Execute strSQL
    
    strSQL = "Delete From " & vBill & " Where " & vKeyField & "='" & vKeyValue & "'"
    Gm.cnnTool.cnn.Execute strSQL
End Sub


'将草稿表的单据登帐到正式表
'vDraftBill：草稿主表名称
'vDraftDetail：草稿明细表名称
'vBill：正式表主表名称
'vDetail：正式表明细表名称
'vKeyField：主键字段名称
'vID：主键字段上的VALUE
Private Sub BL_Auto_Copy_Save(ByVal vDraftBill As String, ByVal vDraftDetail As String, ByVal vBill As String, ByVal vDetail As String, ByVal vKeyField As String, ByVal vID As Long)
    Dim oRs As New RecordSet
    Dim tRs As New RecordSet
    Dim strSQL As String
    Dim i As Long
    
    
    '拷贝主表
    strSQL = "Select * From " & vDraftBill & " Where " & vKeyField & "='" & vID & "'"
    Set oRs = New RecordSet
    oRs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If oRs.RecordCount <= 0 Then
        oRs.Close
        Set oRs = Nothing
        Exit Sub
    End If
    
    strSQL = "Select * From " & vBill & " Where 1=0"
    Set tRs = New RecordSet
    tRs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    tRs.AddNew
    For i = 0 To oRs.Fields.Count - 1
        If Judge_Field_InRs(tRs, oRs(i).name) = True Then
            tRs(Trim$(oRs(i).name)) = oRs(i)
        End If
    Next
    tRs.Update
    tRs.Close
    Set tRs = Nothing
    
    oRs.Close
    Set oRs = Nothing
    
    
    '拷贝明细表
    strSQL = "Select * From " & vDraftDetail & " Where " & vKeyField & "='" & vID & "'"
    Set oRs = New RecordSet
    oRs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If oRs.RecordCount > 0 Then
        Do While Not oRs.EOF
            Set tRs = New RecordSet
            strSQL = "Select * From " & vDetail & " Where 1=0"
            tRs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            tRs.AddNew
            For i = 0 To oRs.Fields.Count - 1
                If Judge_Field_InRs(tRs, oRs(i).name) = True Then
                    tRs(Trim(oRs(i).name)) = oRs(i)
                End If
            Next
            tRs.Update
            tRs.Close
            Set tRs = Nothing
            oRs.movenext
        Loop
    End If
    oRs.Close
    Set oRs = Nothing
End Sub

'生成单据编号B_CodeID
'单据对象的设置表为G_BL，根据传入参数（对象编号），自动获取其引文前缀和正式表名
Public Function GetFrameCodeDetail_01(ByVal m_ObjectID As String) As String
    On Error Resume Next
    Dim strTmpBH As String
    Dim strTmpBHLast As String
    Dim strTmpMonth As String
    Dim strTmpDay As String
    Dim mstrSQL As String
    Dim rs As New RecordSet
    Dim gdateSystemDat As Date
    Dim rstemp As RecordSet
    Dim mvarm_BID As String
    Dim strSQL As String
    Dim m_DraftMainTable As String
    Dim m_MainTable As String
       
    Set rstemp = New RecordSet
    strSQL = "Select * From G_BL Where B_ObjectID='" & m_ObjectID & "'"
    rstemp.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    mvarm_BID = rstemp("B_BID")
    m_DraftMainTable = rstemp("B_DraftMainTable")
    m_MainTable = rstemp("B_MainTable")
    rstemp.Close
    Set rstemp = Nothing
      
    gdateSystemDat = Now
    Set rs = New RecordSet
    strTmpMonth = Trim(Month(gdateSystemDat))
    If Len(Trim(strTmpMonth)) = 1 Then
        strTmpMonth = "0" & strTmpMonth
    End If
    strTmpDay = Trim(Day(gdateSystemDat))
    If Len(Trim(strTmpDay)) = 1 Then
        strTmpDay = "0" & strTmpDay
    End If
    strTmpBH = Trim(mvarm_BID) & Trim(Year(gdateSystemDat)) & strTmpMonth & strTmpDay
    'Debug.Print strTmpBH
    mstrSQL = "Select CASE WHEN ISNULL(P1.B_CodeID,0)>ISNULL(P2.B_CodeID,0) THEN P1.B_CodeID"
    mstrSQL = mstrSQL & " Else P2.B_CodeID End as B_PCodeID"
    mstrSQL = mstrSQL & " From (Select Max(B_CodeID) as B_CodeID"
    mstrSQL = mstrSQL & " From " & m_DraftMainTable
    mstrSQL = mstrSQL & " Where B_CodeID Like '" & Trim(strTmpBH) & "%') as P1,"
    mstrSQL = mstrSQL & " (Select Max(B_CodeID) as B_CodeID"
    mstrSQL = mstrSQL & " From " & m_MainTable
    mstrSQL = mstrSQL & " Where B_CodeID Like '" & Trim(strTmpBH) & "%') as P2"
    'Debug.Print mstrSQL
    rs.Open mstrSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    If IsNull(rs(0)) Then
        '如果没有以前的记录
        strTmpBH = strTmpBH & "0001"
        GetFrameCodeDetail_01 = strTmpBH
    Else
        strTmpBHLast = Trim(str(Val(Mid(Trim(rs(0)), 12, 4)) + 1))
        If Len(Trim(strTmpBHLast)) < 4 Then
            strTmpBH = strTmpBH & String(4 - Len(Trim(strTmpBHLast)), "0") & strTmpBHLast
        Else
            strTmpBH = strTmpBH & strTmpBHLast
        End If
        GetFrameCodeDetail_01 = strTmpBH
    End If
    rs.Close
    Set rs = Nothing
End Function


'生成成品退仓回修单
'vID：发货单主表G_BillCP中的B_ID
Public Sub CreateCPTC(ByVal vID As Long)
    Dim lID As Long
    '根据原单据复制一个新的单据在草稿表，并且返回草稿主表中新建的B_ID
    lID = BL_Auto_Copy_Bill("G_DraftBillCP", "G_BillCP", "B_ID", vID)
    
    '根据新的B_ID从正式表中复制数据到草稿明细表
    BL_Auto_Copy_Detail "G_DraftBillDetailCP", "G_BillDetailCP", "B_ID", "B_ItemID", vID, lID
    
    '将草稿主表中的B_ObjectID,B_CodeID,B_BID修改为目标单据对象编号对应的参数
    BL_Auto_Copy_SaveTo "G_DraftBillCP", "B_ID", lID, "120009"
    
    '将草稿表数据登帐为正式表
    BL_Auto_Copy_Save "G_DraftBillCP", "G_DraftBillDetailCP", "G_BillCP", "G_BillDetailCP", "B_ID", lID
    
    '删除草稿表的明细表和主表的数据
    BL_Auto_Delete_Draft "G_DraftBillCP", "G_DraftBillDetailCP", "B_ID", lID
    
    MsgBox "成品退仓回修单生成完毕！", vbOKOnly + vbInformation, "提示"
End Sub

'================================================================



'以下为连带删除：
'===============================================================
'删除一个成品出库单的时候，同时删除其来源的成品入库单的数据
Private Sub DelRelated()
'    If g_rsDelRelated.State <> adStateOpen Then
'        Exit Sub
'    End If
'
'    g_rsDelRelated.Filter = ""
'    If g_rsDelRelated.RecordCount <= 0 Then
'        Exit Sub
'    End If
'
'    g_rsDelRelated.Filter = " B_oObjectID='" & mvarObjectID & "'"
'    If g_rsDelRelated.RecordCount <= 0 Then
'        g_rsDelRelated.Filter = ""
'        Exit Sub
'    End If
'
'    Dim rs As New Recordset
'    Dim rs1 As New Recordset
'
'    Dim strSQL As String
'    strSQL = "Select DISTINCT B_ID From " & g_rsDelRelated!B_tDetail & " Where " & g_rsDelRelated!B_TFieldName & " In"
'    strSQL = strSQL & " ("
'    strSQL = strSQL & " Select " & g_rsDelRelated!B_OFieldName
'    strSQL = strSQL & " From " & g_rsDelRelated!B_oDetail
'    strSQL = strSQL & " Where B_ID=" & mvarm_ID
'    strSQL = strSQL & " )"
'    Debug.Print strSQL
'    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
'
'
'    '执行删除明细数据
'    strSQL = "Delete From " & g_rsDelRelated!B_tDetail & " Where " & g_rsDelRelated!B_TFieldName & " IN"
'    strSQL = strSQL & " ("
'    strSQL = strSQL & " Select " & g_rsDelRelated!B_OFieldName
'    strSQL = strSQL & " From " & g_rsDelRelated!B_oDetail
'    strSQL = strSQL & " Where B_ID=" & mvarm_ID
'    strSQL = strSQL & " )"
'    Debug.Print strSQL
'    Gm.cnnTool.cnn.Execute strSQL
'
'    If rs.RecordCount > 0 Then
'        Do While Not rs.EOF
'            Set rs1 = New Recordset
'            strSQL = "Select * From " & g_rsDelRelated!B_tDetail & " Where B_ID=" & rs!B_ID
'            Debug.Print strSQL
'            rs1.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
'
'            If rs1.RecordCount <= 0 Then
'                strSQL = "Delete from " & g_rsDelRelated!B_tBill & " Where B_ID=" & rs!B_ID
'                Gm.cnnTool.cnn.Execute strSQL
'            End If
'            rs1.Close
'            Set rs1 = Nothing
'
'            rs.MoveNext
'        Loop
'    End If
'
'    rs.Close
'    Set rs = Nothing
End Sub
'===============================================================

Private Sub TDBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Adodc2.RecordSet.RecordCount <= 0 Then
        Exit Sub
    End If
    
    If mvarboolIsDraft = False Then
        Exit Sub
    End If
    
    Dim rs As RecordSet
    Dim i As Long
    
    'Ctl+Enter复制上面一行数据，创建新的行
    If Shift = vbCtrlMask And KeyCode = 13 Then
        CopyDetailOne Adodc2.RecordSet!B_ItemID
    End If

End Sub

'草稿状态下复制当前所在行的明细数据
Private Sub CopyDetailOne(ByVal vItemID As Long)
    Dim szFields As String
    Dim cls1 As New clsDataBase
    
    szFields = cls1.GetTableFields(m_DraftDetailTable)
    Debug.Print szFields
    
    strSQL = "INSERT INTO " & m_DraftDetailTable
    strSQL = strSQL & " (" & szFields & ")"
    strSQL = strSQL & " Select " & szFields & " From " & m_DraftDetailTable & " WHERE B_ItemID='" & vItemID & "'"
    Debug.Print strSQL
    Gm.cnnTool.cnn.Execute strSQL
    
    Adodc2.RecordSet.requery
End Sub

Private Function JudgeNotNull() As Boolean
    On Error GoTo IFERR
    JudgeNotNull = True
    
    If Adodc2.RecordSet.RecordCount <= 0 Then
        Exit Function
    End If
    
    'A_rsBLField
    
    A_rsBLField.Filter = " isnull(B_NotNull,0)=1 or isnull(B_NotNull,0)=-1"
    If A_rsBLField.RecordCount <= 0 Then
        Exit Function
    End If
    
    Dim rs As New RecordSet
    Dim szTemp As String
    Set rs = Adodc2.RecordSet.Clone
    Do While Not rs.EOF
        szTemp = IIf(IsNull(rs(Trim$(A_rsBLField!B_FieldName))), "", rs(Trim$(A_rsBLField!B_FieldName)))
        If Len(szTemp) <= 0 Then
            szTemp = A_rsBLField!B_CnName & "不可为空！"
            JudgeNotNull = False
            rs.Close
            Set rs = Nothing
            MsgBox szTemp, vbOKOnly + vbInformation, "提示"
            Exit Function
        End If
        rs.movenext
    Loop
    
    Exit Function
IFERR:
    Dim szTip As String
    szTip = Err.Description
    MsgBox szTip, vbOKOnly + vbInformation, "提示"
End Function


Public Function GetSuitTables(ByVal vObjectID As String) As dmSuitTables
    Dim strSQL As String
    Dim rs As New RecordSet
    
    strSQL = "SELECT * FROM G_BL WHERE B_ObjectID='" & vObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    Dim oDM As New dmSuitTables
    With oDM
        .DraftBill = rs!B_DraftMainTable
        .DraftDetail = rs!B_DraftDetailTable
        .Bill = rs!B_MainTable
        .Detail = rs!B_DetailTable
    End With
    
    Set GetSuitTables = oDM
    
    rs.Close
    Set rs = Nothing
End Function

'设置主表字段默认
Private Sub SetBillFieldDefault(ByRef vRs As RecordSet)
    If A_rsBillField.State <> adStateOpen Then
        Exit Sub
    End If
    
    If A_rsBillField.RecordCount <= 0 Then
        Exit Sub
    End If
    Dim szDefault As String
    
    Do While Not A_rsBillField.EOF
        szDefault = IIf(IsNull(A_rsBillField!B_Default), "", A_rsBillField!B_Default)
        If Len(szDefault) > 0 Then
            vRs(Trim$(A_rsBillField!B_FieldName)) = szDefault
        End If
        A_rsBillField.movenext
    Loop
End Sub


Private Sub WriteWFInstance_Save(ByVal vID As Long)
    Dim cls1 As New clsWorkFlow
    cls1.InitCls mvarObjectID
    cls1.WriteInstance vID, Action.Save
End Sub

Private Function JudgeBeforePrint() As Boolean
'    Dim cls1 As New clsWorkFlow
'    cls1.initCls mvarObjectID
'    JudgeBeforePrint = cls1.IsTheProgress(mvarm_ID, Action.toPrint)

    Dim rs As New RecordSet
    strSQL = "exec dbo.usp_GetAudit '" & mvarm_ID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        JudgeBeforePrint = True
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    Dim lAudit As Long
    lAudit = IIf(IsNull(rs!B_Audit), 0, rs!B_Audit)
    If lAudit = 0 Then
        'JudgeBeforePrint = False
        'MsgBox "未审核不可打印！", vbOKOnly + vbInformation, "提示"
        JudgeBeforePrint = True
    Else
        JudgeBeforePrint = True
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function


Private Sub WriteLogPrint(ByVal vID As Long)
    Dim cls1 As New clsLog
    cls1.WriteLog_BillPrint vID, mvarObjectID
End Sub


Private Sub ADORequery(ByVal vColIndex As Long)
    Dim szField As String
    Dim lCtlType As Long   '网格控件中CELL中的控件类型
    
    Dim bmark
    
    szField = TDBGrid1.Columns(vColIndex).DataField
    
    A_rsBLField.Filter = " B_FieldName='" & szField & "'"
    If A_rsBLField.RecordCount <= 0 Then
        A_rsBLField.Filter = ""
        Exit Sub
    End If
    
    lCtlType = IIf(IsNull(A_rsBLField!B_ControlValue), 0, A_rsBLField!B_ControlValue)
    Select Case lCtlType
        Case 21
            bmark = TDBGrid1.Bookmark
            Adodc2.RecordSet.requery
            TDBGrid1.PostMsg 81
            TDBGrid1.Bookmark = bmark
            TDBGrid1.Col = vColIndex + 1
    End Select
End Sub

'拼接制单员只能查看自己做的单据的SQL语句
Private Function MontageCreateOperator(ByVal vSQL As String) As String
    If Gm.OnlyDataBreak = 1 Then
        MontageCreateOperator = vSQL
        Exit Function
    End If
    
    If Len(A_CreateOperator) > 0 Then
        vSQL = vSQL & " And " & m_TableName & "." & A_CreateOperator & "='" & Gm.SysID.SystemUser & "'"
    End If
    MontageCreateOperator = vSQL
End Function


Private Function MontageCreateOperator8Table(ByVal vSQL As String, _
    ByVal vTableName As String) As String
    
    If Gm.OnlyDataBreak = 0 Then
    If Len(A_CreateOperator) > 0 Then
        vSQL = vSQL & " And " & vTableName & "." & A_CreateOperator & "='" & Gm.SysID.SystemUser & "'"
    End If
    End If
    MontageCreateOperator8Table = vSQL
End Function
