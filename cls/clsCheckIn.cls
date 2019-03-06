VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCheckIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'----新东方软件
'----编制人:陈超
'----单据编号
Public m_ID As Long
'----单据类别
Public OK As Boolean

Public Event OnCheckIn(ByRef oCnn As Connection)

Dim m_MainTable As String                   '登帐主表
Dim m_DetailTable As String                 '登帐明细表
Dim m_DraftMainTable As String              '草稿主表
Dim m_DraftDetailTable As String            '草稿

Dim m_MainFieldList As String
Dim m_DetailFieldList As String

Dim rsStatus As New Recordset
'保持属性值的局部变量
Private mvarObjectID As String '局部复制
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



Public Sub CheckInEx()
    On Error GoTo IFERR
    
    GetBillStatus Gm.cnnTool.cnn
    If CheckBillItems(Gm.cnnTool.cnn) = False Then
        OK = False
        Exit Sub
    End If
    
    Dim oCnn As New Connection
    Set oCnn = New Connection

    Err.Clear
    
    oCnn.Open Gm.cnnTool.cnnStr
    oCnn.BeginTrans

    '获取字段名
    GetBillFieldList oCnn
    
    '设置B_Closed=1(为登帐状态)
    CloseBill oCnn
    
    
    '将草稿数据登帐到正式表中
    AddBillEx oCnn
    
    '删除草稿表数据
    DeleteOldBillList oCnn

    oCnn.CommitTrans
    
    OK = True
    Exit Sub
IFERR:
    OK = False
    'oCnn.RollbackTrans
    Set oCnn = Nothing
    MsgBox Err.Description
    Exit Sub
End Sub

'----登帐
Public Sub CheckIn()
    On Error GoTo IFERR
    
    
    GetBillStatus Gm.cnnTool.cnn
    If CheckBillItems(Gm.cnnTool.cnn) = False Then
        OK = False
        Exit Sub
    End If
    
    Dim oCnn As New Connection
    Set oCnn = New Connection

    Err.Clear
    oCnn.Open Gm.cnnTool.cnnStr
    oCnn.BeginTrans

    '获取字段名
    GetBillFieldList oCnn
    
    '设置B_Closed=1(为登帐状态)
    CloseBill oCnn
    
    '将草稿表中的数据登帐到正式表中
    AddBill oCnn
    
    '删除草稿表数据
    DeleteOldBillList oCnn

    oCnn.CommitTrans
    
    OK = True
    Exit Sub
IFERR:
    OK = False
    'oCnn.RollbackTrans
    Set oCnn = Nothing
    MsgBox Err.Description
    Exit Sub
End Sub

'----锁定单据
Private Sub CloseBill(ByRef oCnn As Connection)
    Dim strSQL As String
    strSQL = "Update " & m_DraftMainTable & " Set B_Closed=1 Where B_ID=" & m_ID
    oCnn.Execute strSQL
End Sub

'----取得单据状态
Private Sub GetBillStatus(ByRef oCnn As Connection)
    Dim strSQL As String
    Set rsStatus = New Recordset
    strSQL = "Select * From G_BL Where B_ObjectID='" & Trim(mvarObjectID) & "'"
    rsStatus.Open strSQL, oCnn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsStatus.EOF Then
        m_MainTable = rsStatus("B_MainTable")
        m_DetailTable = rsStatus("B_DetailTable")
        m_DraftMainTable = rsStatus("B_DraftMainTable")
        m_DraftDetailTable = rsStatus("B_DraftDetailTable")
    End If
End Sub

'----增加表单数据
'----增加明细表数据
Private Sub AddBillEx(ByRef oCnn As Connection)
    Dim strSQL As String
    Dim rs As Recordset
    
        '明细表规则
    '=================
    Set rs = New Recordset
    strSQL = "Select * From G_RuleBillDetail Where B_ObjectID='" & mvarObjectID & "' Order by B_Order"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    '=================
    
    
    
    '----增加主表数据
    strSQL = "Insert Into " & m_MainTable
    strSQL = strSQL & " (" & m_MainFieldList & ")"
    strSQL = strSQL & " Select " & m_MainFieldList & " From " & m_DraftMainTable
    strSQL = strSQL & " Where B_ID=" & m_ID
    Debug.Print strSQL
    oCnn.Execute strSQL

    '----增加明细表数据
    strSQL = "Insert Into " & m_DetailTable
    strSQL = strSQL & " (" & m_DetailFieldList & ")"
    strSQL = strSQL & " Select " & m_DetailFieldList & " From " & m_DraftDetailTable
    strSQL = strSQL & " Where B_ID=" & m_ID
    strSQL = strSQL & " And (B_CheckIDCCC=1 or B_CheckIDCCC=-1)"
    Debug.Print strSQL
    
    
    '=================
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
        
            strSQL = strSQL & rs("B_RuleString")
    
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    '=================



    'strSQL = strSQL & " Order By B_Qty,B_KQty"
    'Debug.Print strSQL
    
    oCnn.Execute strSQL
    RaiseEvent OnCheckIn(oCnn)
End Sub

'----增加表单数据
'----增加明细表数据

'从草稿表拷贝数据到正式表
'并且有选择性的拷贝(即实现规则)
Private Sub AddBill(ByRef oCnn As Connection)
    Dim strSQL As String
    Dim rs As Recordset
        
    
    '明细表规则
    '=================
    Set rs = New Recordset
    strSQL = "Select * From G_RuleBillDetail Where B_ObjectID='" & mvarObjectID & "' Order by B_Order"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    '=================
    
    
    
    '----增加主表数据
    strSQL = "Insert Into " & m_MainTable
    strSQL = strSQL & " (" & m_MainFieldList & ")"
    strSQL = strSQL & " Select " & m_MainFieldList & " From " & m_DraftMainTable
    strSQL = strSQL & " Where B_ID=" & m_ID

    oCnn.Execute strSQL
    
    
    
    '财务的单据沿用仓库的单据保存时候会造成B_ObjectID不一样,在此修正
    '=========================
'    Dim rs1 As RecordSet
'    Set rs1 = New RecordSet
'    strSQL = "Select * From G_BLConvert where B_oObjectID='" & mvarObjectID & "'"
'    rs1.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
'
'    If rs1.RecordCount > 0 Then
'        strSQL = "Update " & m_MainTable & " Set B_ObjectID='" & rs1("B_tObjectID") & "' Where B_ID=" & m_ID
'        oCnn.Execute strSQL
'    End If
'    rs1.Close
'    Set rs1 = Nothing
    '=========================
    
    
    
    

    '----增加明细表数据
    strSQL = "Insert Into " & m_DetailTable
    strSQL = strSQL & " (" & m_DetailFieldList & ")"
    strSQL = strSQL & " Select " & m_DetailFieldList & " From " & m_DraftDetailTable
    strSQL = strSQL & " Where B_ID=" & m_ID
    
    
    
    '=================
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            strSQL = strSQL & rs("B_RuleString")
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    '=================
    
    
    oCnn.Execute strSQL
    RaiseEvent OnCheckIn(oCnn)
End Sub

Private Sub GetBillFieldList(ByRef oCnn As Connection)
    '----以下为增加项的SQL
    'Insert Into MainTable (Fields...) Select (Fields...) From MainTable Where B_ID=m_ID
    Dim rs1 As New Recordset        '----草稿记录
    Dim rs2 As New Recordset        '----帐记录
    Dim strSQL As String
    Dim i As Integer
    
    Set rs1 = New Recordset
    Set rs2 = New Recordset
    
    '----查询主表字段
    'm_MainFieldList As String
    strSQL = "Select * From " & m_MainTable & " Where 1=0"
    rs1.Open strSQL, oCnn, adOpenStatic, adLockReadOnly
    strSQL = "Select * From " & m_DraftMainTable & " Where 1=0"
    rs2.Open strSQL, oCnn, adOpenStatic, adLockReadOnly
    m_MainFieldList = ""
    For i = 0 To rs1.Fields.Count - 1
        If IsInList(rs2, rs1.Fields(i).name) = True Then
            m_MainFieldList = m_MainFieldList & Trim(rs1.Fields(i).name) & ","

        End If
    Next
    m_MainFieldList = Mid(m_MainFieldList, 1, Len(m_MainFieldList) - 1)
    
    '----查询明细表字段
    'm_DetailFieldList As String
    Set rs1 = New Recordset
    Set rs2 = New Recordset
    strSQL = "Select * From " & m_DetailTable & " Where 1=0"
    rs1.Open strSQL, oCnn, adOpenStatic, adLockReadOnly
    
    strSQL = "Select * From " & m_DraftDetailTable & " Where 1=0"
    rs2.Open strSQL, oCnn, adOpenStatic, adLockReadOnly
    m_DetailFieldList = ""
    For i = 0 To rs1.Fields.Count - 1
        If IsInList(rs2, rs1.Fields(i).name) = True Then
            m_DetailFieldList = m_DetailFieldList & Trim(rs1.Fields(i).name) & ","
        End If
    Next
    m_DetailFieldList = Mid(m_DetailFieldList, 1, Len(m_DetailFieldList) - 1)
    
End Sub

Private Function IsInList(ByRef rs As Recordset, ByVal m_FieldName As String) As Boolean
    Dim i As Integer
    For i = 0 To rs.Fields.Count - 1
        If Trim(rs.Fields(i).name) = Trim(m_FieldName) Then
            IsInList = True
            Exit Function
        End If
    Next
    IsInList = False
End Function

'----删除原明细表
'----删除原表
Private Sub DeleteOldBillList(ByRef oCnn As Connection)
    Dim strSQL As String
    strSQL = "Delete From " & m_DraftDetailTable & " Where B_ID=" & m_ID
    Debug.Print strSQL
    oCnn.Execute strSQL
    
    strSQL = "Delete From " & m_DraftMainTable & " Where B_ID=" & m_ID
    Debug.Print strSQL
    oCnn.Execute strSQL
End Sub

'----判定单据项是否符合情况
Private Function CheckBillItems(ByRef oCnn As Connection) As Boolean
    Dim adoCom As New ADODB.Command
    Dim m_Err As Integer

    With adoCom
        .ActiveConnection = oCnn
        .CommandText = "S_CheckBill"
        .CommandType = adCmdStoredProc
        
        .Parameters.Append adoCom.CreateParameter("@m_ObjectID", adVarChar, adParamInput, 20, mvarObjectID)
        .Parameters.Append adoCom.CreateParameter("@m_ID", adInteger, adParamInput, 4, m_ID)
        .Parameters.Append adoCom.CreateParameter("@m_MainTable", adVarChar, adParamInput, 50, m_DraftMainTable)
        .Parameters.Append adoCom.CreateParameter("@m_DetailTable", adVarChar, adParamInput, 50, m_DraftDetailTable)
        .Parameters.Append adoCom.CreateParameter("@m_ErrID", adVarChar, adParamInputOutput, 4, m_Err)
        
        .Execute
    End With
    If adoCom.Parameters("@m_ErrID").Value > 0 Then
        ShowError oCnn, adoCom.Parameters("@m_ErrID").Value
        CheckBillItems = False
        Exit Function
    End If
    CheckBillItems = True
End Function

Private Sub ShowError(ByRef oCnn As Connection, ByVal m_Index As Integer)
    Dim rs As New Recordset
    Dim strSQL As String
    
    Set rs = New Recordset
    strSQL = "Select * From G_OnErrShow Where B_OnErrShow=" & m_Index
    rs.Open strSQL, oCnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        MsgBox rs("B_ShowChinese"), vbExclamation, "错误"
    End If
    
    rs.Close
    Set rs = Nothing
End Sub


