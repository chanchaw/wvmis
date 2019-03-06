VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'系统命令类
'用一全局
'保持属性值的局部变量
Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_HELP_CONTEXT = &HF

Private mvarObjectID As String '局部复制
Private oCtl As Object
'保持属性值的局部变量
Private mvarCommand As String '局部复制
'保持属性值的局部变量
Private mvarObjectName As String '局部复制

Private clsCommValues1 As New clsCommValues

Private rsCommand As New RecordSet

Private aValue
Private sValue As String


'单据窗体打开基础资料窗体中参与的工作的参数
Public frmName As String
Public frm1 As Object


'弹出明细报表的参数设置
'==========================
Public A_SetPara As String
'==========================

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Property Let ObjectName(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ObjectName = 5
    mvarObjectName = vData
End Property

Public Property Get ObjectName() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ObjectName
    ObjectName = mvarObjectName
End Property

Public Property Let Command(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Command = 5
    mvarCommand = vData
End Property

Public Property Get Command() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Command
    Command = mvarCommand
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

'命令
Public Sub Execute(ByVal m_ObjectID As String, ByVal m_ObjectName As String, ByVal m_Command As String, ByRef oSelf As Object, Optional ByVal Value As String = "")
    On Error GoTo IFERR
    
    Gm.log4Runtime "clsCommand.Execute"
    
    ObjectID = m_ObjectID
    Command = m_Command
    ObjectName = m_ObjectName
    
    If Not oSelf Is Nothing Then
        Set oCtl = oSelf
    End If
    
    aValue = Split(Value, ",")
    sValue = Value
    
    If CheckCommand = False Then
        InsertLogin "未授权的访问!"
        MsgBox "未授权的访问!", vbOKOnly + vbInformation, "提示"
        Gm.log4Runtime "没有访问权限，即将退出"
        Exit Sub
    End If
    
    Gm.log4Runtime "开始进入分支打开页面"
    Select Case Mid(m_ObjectID, 1, 2)
        '单表
        Case "11"
            Gm.log4Runtime "进入单表的打开命令"
            ExecBLSObject
        '表单
        Case "12"
            ExecBLObject
        '报表
        Case "13"
            
            ExecBLRObject
            
        Case "14"
            
            ExecTIFObject
            
        '自定义
        Case "19"
            ExecCustomObject
        '系统
        Case "00"
            ExecSystemObject
    End Select
    InsertLogin "正常"
    Exit Sub
IFERR:
    InsertLogin Err.Description
    Exit Sub
End Sub



'以下为2012-2-12之后添加
'为弹出基础资料的窗体做准备
'=================================
'参数解释:
'frm:当前的单据窗体
'm_fObjectID:弹出窗体的对象(如:110007 客户资料)
'm_fFieldName:单据窗体上主表控件的名字
'm_SendIndex:被弹出对象记录集中即将被传输出去的数据的Index
'后面的参数均为系统原配置参数
Public Sub ExecutePopUp(ByVal BillOrDetail As Integer, ByRef frm As Object, _
    ByVal m_fObjectID As String, ByVal m_fFieldName As String, _
    ByVal m_SendIndex As Integer, ByVal m_ObjectID As String, _
    ByVal m_ObjectName As String, ByVal m_Command As String, _
    ByRef oSelf As Object, Optional ByVal Value As String = "")
    
    On Error GoTo IFERR
    
    ObjectID = m_ObjectID
    Command = m_Command
    ObjectName = m_ObjectName
    
    If Not oSelf Is Nothing Then
        Set oCtl = oSelf
    End If
    
    aValue = Split(Value, ",")
    sValue = Value
    
    If CheckCommand = False Then
        InsertLogin "未授权的访问!"
        MsgBox "未授权的访问!", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
    Select Case Mid(m_ObjectID, 1, 2)
        '单表
        Case "11"
            ExecBLSObjectPopUp BillOrDetail, frm, m_fObjectID, m_fFieldName, m_SendIndex
        '表单
        Case "12"
            ExecBLObject
        '报表
        Case "13"
            ExecBLRObjectPopUp BillOrDetail, frm, m_fObjectID, m_fFieldName, m_SendIndex
            
        Case "14"
            ExecTIFObject

        '自定义
        Case "19"
            ExecCustomObject
        '系统
        Case "00"
            ExecSystemObject
    End Select
    InsertLogin "正常"
    Exit Sub
IFERR:
    InsertLogin Err.Description
    Exit Sub
End Sub


Private Function CheckCommand() As Boolean

    '如果是超级管理员，那么无条件返回TRUE
    Dim sql As String
    Dim rs As New RecordSet
    sql = "select * from G_systemuser where B_username ='" & Gm.SysID.SystemUser & "'"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs!B_SuperAdmin = 1 Then
        CheckCommand = True
        Exit Function
    End If
    If Gm.SysID.SystemUser = "管理员" Then
        CheckCommand = True
        Exit Function
    End If
    
    rsCommand.Filter = "B_ObjectID='" & ObjectID & "'"
    If rsCommand.RecordCount > 0 Then
        CheckCommand = True
    End If
End Function

'执行图片事件
Private Sub ExecTIFObject()
    Dim o As Object
    Dim m_BillName As String
    Dim strSQL As String
    Dim rs As New RecordSet
    Dim m_Doc As String
    strSQL = "Select * From G_DocumentObject Where B_ObjectID='" & ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        m_Doc = rs("B_DocumentType")
        Select Case Command
            Case "LoadObject"
                'GetFormObject o, rs("B_FormName")
                Set o = GetFormNew(rs("B_FormName"))
                o.ObjectID = ObjectID
                o.aValue = sValue
                o.LoadObject
                
                If m_Doc = "tif" Then
                    o.Show
                Else
                    o.Show vbModal
                End If
                
                BringWindow2Top o.hwnd
                
        End Select
    End If
    rs.Close
    Set rs = Nothing
End Sub


'执行自定义事件
Private Sub ExecCustomObject()
    On Error Resume Next
    Dim o As Object
    Dim m_BillName As String
    Dim strSQL As String
    Dim rs As New RecordSet
    
    strSQL = "Select * From G_CustomObject Where B_ObjectID='" & ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        Select Case Command
            Case "LoadObject"
                Set o = Gm.CacheFrms.GetFrm(mvarObjectID)
                'GetFormObject o, rs("B_FormName")
                o.ObjectID = ObjectID
                o.Show
                o.LoadObject
                
                BringWindow2Top o.hwnd
        End Select
    End If
    rs.Close
    Set rs = Nothing
End Sub

'增加事件
Private Sub InsertLogin(ByVal m_ErrMsg As String)
    On Error Resume Next
    Dim strSQL As String
    strSQL = "Insert Into G_SystemLogin (B_ObjectID,B_ObjectName,B_UserName,B_ComputerName,B_Operation,B_IP,B_ErrMsg,B_Parameter)"
    strSQL = strSQL & " Values ("
    
    strSQL = strSQL & "'" & ObjectID & "',"
    strSQL = strSQL & "'" & ObjectName & "',"
    strSQL = strSQL & "'" & Gm.SysID.SystemUser & "',"
    strSQL = strSQL & "'" & clsCommValues1.ComputerName & "',"
    strSQL = strSQL & "'" & Command & "',"
    strSQL = strSQL & "'" & clsCommValues1.GetIPAddress & "',"
    strSQL = strSQL & "'" & m_ErrMsg & "',"
    strSQL = strSQL & "'" & sValue & "'"
    strSQL = strSQL & " )"
    
    Gm.cnnTool.IniConnection8DM Gm.SysID.DBInfo
    Gm.cnnTool.cnn.Execute strSQL
End Sub


Private Sub ExecBLRObjectEx(ByVal m_ObjectName As String)
    Select Case ObjectID
        Case "190044"
            'ExecReportPassWord m_ObjectName
        Case Else
            ExecBLRObject
    End Select
End Sub

'进入报表前的密码验证
Private Sub ExecReportPassWord(ByVal m_ObjectName As String)
'    Dim frm1 As New frmCJAdminReport
'
'    frm1.m_ObjectID = ObjectID
'    frm1.m_ObjectName = m_ObjectName
'
'    frm1.Show vbModal
End Sub

'执行报表事件
Private Sub ExecBLRObject()
    Dim o As Object
    Dim m_BillName As String
    Dim strSQL As String
    Dim rs As New RecordSet
    
    strSQL = "Select * From G_BLR Where B_ObjectID='" & ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        Select Case Command
            Case "LoadObject"
                Set o = Gm.CacheFrms.GetFrm(mvarObjectID)
                'Set o = CacheFrms.GetFrm(mvarObjectID)
                o.ObjectID = ObjectID
                o.aValue = aValue
                
                
                '弹出明细报表的参数的设置
                '============================
                If Len(A_SetPara) > 0 Then
                    o.A_SetPara = A_SetPara
                    o.SetPara
                End If
                '============================
                
                
                
                o.Show
                
                'If ObjectID = "130018" Then
                 '   o.LoadObjectCaptionOnly
                'Else
                    o.LoadObject
                'End If
                
                BringWindow2Top o.hwnd
                
        End Select
    End If
    rs.Close
    Set rs = Nothing
End Sub

'执行单表事件
Private Sub ExecBLSObject()
    Gm.log4Runtime "clsCommand.ExecBLSObject()"

    Dim o As Object
    Dim m_BillName As String
    Dim strSQL As String
    Dim rs As New RecordSet
    
    strSQL = "Select * From G_BLS Where B_ObjectID='" & ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    Gm.log4Runtime "从DB中获取数据完毕"
    If Not rs.EOF Then
        Select Case Command
            Case "LoadObject"
                Gm.log4Runtime "LoadObject"
                
                'GetFormObject o, rs("B_FormName")
                'Set o = Gm.CacheFrms.GetFrm(mvarObjectID)
                Set o = GetFormNew(rs("B_FormName"))
                o.ObjectID = ObjectID
                
                '当frmName长度大于0时候才时候弹出窗体
                If Len(frmName) > 0 Then
                    Set o.frm1 = frm1
                    o.frmName = frmName
                End If
                
                Gm.log4Runtime "即将o.show"
                o.Show
                o.LoadObject
                BringWindow2Top o.hwnd
                Gm.log4Runtime "LoadObject执行完毕"
                
            Case "ViewObject"
                'GetFormObject o, rs("B_EditFormName")
                'Set o = Gm.CacheFrms.GetFrm(mvarObjectID)
                Set o = GetFormNew(rs("B_FormName"))
                o.ObjectID = ObjectID
                o.m_KeyID = aValue(0)
                o.ViewObject ObjectID
                o.Show vbModal
                
                BringWindow2Top o.hwnd
                
            Case "AddNewObject"
                'GetFormObject o, rs("B_EditFormName")
                'Set o = Gm.CacheFrms.GetFrm(mvarObjectID)
                Set o = GetFormNew(rs("B_FormName"))
                o.ObjectID = ObjectID
                o.AddNewObject ObjectID
                o.Show vbModal
                
                BringWindow2Top o.hwnd
                
            Case "EditObject"
                'GetFormObject o, rs("B_EditFormName")
                'Set o = Gm.CacheFrms.GetFrm(mvarObjectID)
                Set o = GetFormNew(rs("B_FormName"))
                o.ObjectID = ObjectID
                o.m_KeyID = aValue(0)
                o.EditObject ObjectID
                o.Show vbModal
                
                BringWindow2Top o.hwnd
                
        End Select
    End If
    rs.Close
    Set rs = Nothing
    
    Gm.log4Runtime "ExecBLSObject执行完毕"
End Sub

'执行表单事件
Private Sub ExecBLObject()
    Dim o As Object
    Dim m_BillName As String
    Dim strSQL As String
    Dim rs As New RecordSet
    
    strSQL = "Select * From G_BL Where B_ObjectID='" & ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        Select Case Command
            Case "LoadObject"
                Set o = Gm.CacheFrms.GetFrm(mvarObjectID)
                'GetFormObject o, rs("B_FormName")
                o.ObjectID = ObjectID
                o.Show
                o.LoadObject
                BringWindow2Top o.hwnd
                
            Case "ViewObject"
                Set o = Gm.CacheFrms.GetFrm(mvarObjectID)
                'GetFormObject o, rs("B_FormName")
                o.ObjectID = ObjectID
                o.Show
                o.ViewObject aValue(0)
                BringWindow2Top o.hwnd
                
            Case "EditObject"
                Set o = Gm.CacheFrms.GetFrm(mvarObjectID)
                'GetFormObject o, rs("B_FormName")
                o.ObjectID = ObjectID
                o.Show
                o.EditObject aValue(0)
                BringWindow2Top o.hwnd
                
            Case "AddNewObject"
                Set o = Gm.CacheFrms.GetFrm(mvarObjectID)
                'GetFormObject o, rs("B_FormName")
                o.ObjectID = ObjectID
                o.Show
                o.AddNewObject
                BringWindow2Top o.hwnd
                
        End Select
    End If
    rs.Close
    Set rs = Nothing
End Sub


'执行系统事件
Private Sub ExecSystemObject()
    Dim hwndHelp As Long
    Dim lContextID As Long
    Select Case ObjectID
        Case "000001" '退出系统
            Unload Gm.frmMain
        Case "000002" '数据库设置
            'cnn.SetConnect

        Case "000004" '用户设置
            frmSetUser.Show vbModal
        Case "000005" '设置口令
            frmSetPassword.Show vbModal

        Case "000007" '窗体动画
            SetAnimate

        Case "000010" '背景
            'frmSelectPicture.Show vbModal

            
        
        Case "000301"
            'frmBackData.Show vbModal
    End Select
End Sub

Private Sub SetAnimate()
    If iAnimate = 0 Then
        iAnimate = 1 '不显示动画
        oCtl.Checked = False
        SaveSetting App.Title, "Settings", "AnimateWindows", "1"
    Else
        iAnimate = 0 '显示动画
        oCtl.Checked = True
        SaveSetting App.Title, "Settings", "AnimateWindows", "0"
    End If
End Sub

Private Sub PrintSetup()
    On Error Resume Next
    Dim o As Object
    Set o = CreateObject("MSComDlg.CommonDialog")
    With o
        .DialogTitle = "页面设置"
        .CancelError = True
        .flags = cdlPDPrintSetup
        .ShowPrinter
    End With
End Sub

Public Sub InitClass()
    If Gm.SysID.SystemUser = "管理员" Then
        Exit Sub
    End If
    '初始化权限
    Dim rs As New RecordSet
    Dim strSQL As String
    
    Set rs = New RecordSet
    Set rsCommand = New RecordSet
    rsCommand.Fields.Append "B_ObjectID", adVarChar, 20, adFldIsNullable
    
    
    strSQL = "Select B_ObjectID From G_UserPro Where B_UserName='" & Gm.SysID.SystemUser & "'"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    rsCommand.Open
    Do While Not rs.EOF
        rsCommand.AddNew
        rsCommand(0) = rs(0)
        rsCommand.Update
        rs.movenext
    Loop
    
    rs.Close
    Set rs = Nothing
    
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set rsCommand = Nothing
End Sub


'执行报表事件
Private Sub ExecBLRObjectPopUp(ByVal BillOrDetail As Integer, ByRef frm As Object, ByVal m_fObjectID As String, ByVal m_fFieldName As String, ByVal m_SendIndex As Integer)
    Dim o As Object
    Dim m_BillName As String
    Dim strSQL As String
    Dim rs As New RecordSet
    
    strSQL = "Select * From G_BLR Where B_ObjectID='" & ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        Select Case Command
            Case "LoadObject"
                'GetFormObject o, rs("B_FormName")
                Set o = GetFormNew(rs("B_FormName"))
                o.ObjectID = ObjectID
                
                '==================
                Set o.fatherFrm = frm
                o.fObjectID = m_fObjectID
                o.fFieldName = m_fFieldName
                o.SendIndex = m_SendIndex
                o.BillOrDetail = BillOrDetail
                '==================
                
                
                o.aValue = aValue
                o.Show
                o.LoadObject
                
                BringWindow2Top o.hwnd
        End Select
    End If
    rs.Close
    Set rs = Nothing
End Sub



Private Sub ExecBLSObjectPopUp(ByVal BillOrDetail As Integer, ByRef frm As Object, ByVal m_fObjectID As String, ByVal m_fFieldName As String, ByVal m_SendIndex As Integer)
    Dim o As Object
    Dim m_BillName As String
    Dim strSQL As String
    Dim rs As New RecordSet
    
    strSQL = "Select * From G_BLS Where B_ObjectID='" & ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        Select Case Command
            Case "LoadObject"
                Set o = GetFormNew(rs("B_FormName"))
                'GetFormObject o, rs("B_FormName")
                o.ObjectID = ObjectID
                '==================
                Set o.fatherFrm = frm
                o.fObjectID = m_fObjectID
                o.fFieldName = m_fFieldName
                o.SendIndex = m_SendIndex
                o.BillOrDetail = BillOrDetail
                '==================
                o.Show
                o.LoadObject
                
                BringWindow2Top o.hwnd
                
            Case "ViewObject"
                Set o = GetFormNew(rs("B_FormName"))
                'GetFormObject o, rs("B_EditFormName")
                o.ObjectID = ObjectID
                o.m_KeyID = aValue(0)
                o.ViewObject ObjectID
                o.Show vbModal
                
                BringWindow2Top o.hwnd
                
            Case "AddNewObject"
                Set o = GetFormNew(rs("B_FormName"))
                'GetFormObject o, rs("B_EditFormName")
                o.ObjectID = ObjectID
                o.AddNewObject ObjectID
                o.Show vbModal
                
                BringWindow2Top o.hwnd
                
            Case "EditObject"
                Set o = GetFormNew(rs("B_FormName"))
                'GetFormObject o, rs("B_EditFormName")
                o.ObjectID = ObjectID
                o.m_KeyID = aValue(0)
                o.EditObject ObjectID
                o.Show vbModal
                
                BringWindow2Top o.hwnd
                
        End Select
    End If
    rs.Close
    Set rs = Nothing
End Sub


'从新盛毛纺拷贝来2014年11月6日 09:55:30
'为弹出基础资料的窗体做准备
'=================================
'参数解释:
'frm:当前的单据窗体
'm_fObjectID:弹出窗体的对象(如:110007 客户资料)
'm_fFieldName:单据窗体上主表控件的名字
'm_SendIndex:被弹出对象记录集中即将被传输出去的数据的Index
'后面的参数均为系统原配置参数
Public Sub ExecutePopUp01(ByVal BillOrDetail As Integer, ByRef frm As Object, ByVal m_fObjectID As String, ByVal m_fFieldName As String, ByVal m_SendFieldName As String, ByVal m_ObjectID As String, ByVal m_ObjectName As String, ByVal m_Command As String, ByRef oSelf As Object, Optional ByVal Value As String = "")
    On Error GoTo IFERR
    
    ObjectID = m_ObjectID
    Command = m_Command
    ObjectName = m_ObjectName
    
    If Not oSelf Is Nothing Then
        Set oCtl = oSelf
    End If
    
    aValue = Split(Value, ",")
    sValue = Value
    
    If CheckCommand = False Then
        InsertLogin "未授权的访问!"
        Exit Sub
    End If
    
    Select Case Mid(m_ObjectID, 1, 2)
        '单表
        Case "11"
            ExecBLSObjectPopUp BillOrDetail, frm, m_fObjectID, m_fFieldName, m_SendFieldName
        '表单
        Case "12"
            ExecBLObject
        '报表
        Case "13"
            ExecBLRObjectPopUp01 BillOrDetail, frm, m_fObjectID, m_fFieldName, m_SendFieldName
            
        Case "14"
            ExecTIFObject

        '自定义
        Case "19"
            ExecCustomObject
        '系统
        Case "00"
            ExecSystemObject
    End Select
    InsertLogin "正常"
    Exit Sub
IFERR:
    InsertLogin Err.Description
    Exit Sub
End Sub



'执行报表事件
Private Sub ExecBLRObjectPopUp01(ByVal BillOrDetail As Integer, ByRef frm As Object, ByVal m_fObjectID As String, ByVal m_fFieldName As String, ByVal m_SendFieldName As String)
    Dim o As Object
    Dim m_BillName As String
    Dim strSQL As String
    Dim rs As New RecordSet
    
    strSQL = "Select * From G_BLR Where B_ObjectID='" & ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        Select Case Command
            Case "LoadObject"
                'GetFormObject o, rs("B_FormName")
                Set o = GetFormNew(rs("B_FormName"))
                o.ObjectID = ObjectID
                
                '==================
                Set o.fatherFrm = frm
                o.fObjectID = m_fObjectID
                o.fFieldName = m_fFieldName
                'o.SendIndex = m_SendIndex
                o.SendFieldName = m_SendFieldName
                o.BillOrDetail = BillOrDetail
                '==================
                
                
                o.aValue = aValue
                o.Show
                o.LoadObject
                
                BringWindow2Top o.hwnd
        End Select
    End If
    rs.Close
    Set rs = Nothing
End Sub

