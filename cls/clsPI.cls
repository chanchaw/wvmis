VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'在2016年6月21日 15:55:46在银海制作
'作者：chanchaw
'在表G_UserPro中将各个权限独立为一个个单独的字段，1表示拥有该权限，0表示没有

Private A_rs As New RecordSet
Private strSQL As String

Public Sub InitCls()
    '当前登录软件的用户名 - 软件级全局变量
    Set A_rs = New RecordSet
    strSQL = "Select * From G_UserPro where B_UserName='" & Gm.SysID.SystemUser & "'"
    Debug.Print strSQL
    A_rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
End Sub

'刷新权限缓存
Public Sub RefreshPermission()
    InitCls
End Sub

'判断是否拥有一个权限
Private Function JudgePermission(ByVal vObjectID As String, ByVal vPermissionField As String) As Boolean
    A_rs.Filter = " B_ObjectID='" & vObjectID & "'"
    If A_rs.RecordCount <= 0 Then
        JudgePermission = False
        Exit Function
    End If

    Dim lPermission As Long
    vPermissionField = Trim$(vPermissionField)
    lPermission = IIf(IsNull(A_rs(vPermissionField)), 0, A_rs(vPermissionField))
    If lPermission = 1 Then
        JudgePermission = True
    Else
        JudgePermission = False
    End If
    
End Function

Public Function JudgeNew(ByVal vObjectID As String) As Boolean
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim a As String
    a = Gm.SysID.SystemUser
    sql = "select * from G_systemuser where B_UserName='" & a & "'"
    rs1.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs1!B_SuperAdmin = 1 Then
        JudgeNew = True
        Exit Function
    End If
    If Gm.IsDeveloper = True Then
        JudgeNew = True
        Exit Function
    End If
    
    
    '不使用权限功能
    If Gm.DBCfgOneInt.UsingFineGrainedAuthority = False Then
        JudgeNew = True
        Exit Function
    End If
    
    If JudgePermission(vObjectID, "B_New") = False Then
        JudgeNew = False
        MsgBox "当前用户没有新增权限！", vbOKOnly + vbInformation, "提示"
    Else
        JudgeNew = True
    End If

End Function


Public Function JudgeUpdate(ByVal vObjectID As String) As Boolean
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim a As String
    a = Gm.SysID.SystemUser
    sql = "select * from G_systemuser where B_UserName='" & a & "'"
    rs1.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs1!B_SuperAdmin = 1 Then
        JudgeUpdate = True
        Exit Function
    End If
    If Gm.IsDeveloper = True Then
        JudgeUpdate = True
        Exit Function
    End If
    
    
    '不使用权限功能
    If Gm.DBCfgOneInt.UsingFineGrainedAuthority = False Then
        JudgeUpdate = True
        Exit Function
    End If
    
    
    If JudgePermission(vObjectID, "B_Update") = False Then
        JudgeUpdate = False
        MsgBox "当前用户没有修改权限！", vbOKOnly + vbInformation, "提示"
    Else
        JudgeUpdate = True
    End If
    
End Function

'用于对于单据进行修改时
'vObjectID:单据对象编号
'vID：单据主表主键字段B_ID
Public Function JudgeUpdateBill(ByVal vObjectID As String, ByVal vID As Long) As Boolean
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim a As String
    a = Gm.SysID.SystemUser
    sql = "select * from G_systemuser where B_UserName='" & a & "'"
    rs1.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs1!B_SuperAdmin = 1 Then
        JudgeUpdateBill = True
        Exit Function
    End If
    If Gm.IsDeveloper = True Then
        JudgeUpdateBill = True
        Exit Function
    End If
    
    
    '不使用权限功能
    If Gm.DBCfgOneInt.UsingFineGrainedAuthority = False Then
        JudgeUpdateBill = True
        Exit Function
    End If
    
    
    If JudgePermission(vObjectID, "B_Update") = False Then
        JudgeUpdateBill = False
        MsgBox "当前用户没有修改权限！", vbOKOnly + vbInformation, "提示"
        JudgeUpdateBill = False
        Exit Function
    Else
        JudgeUpdateBill = True
    End If
    
    
    '不做检测
    If Gm.DBCfgOneInt.UsingFineGrainedAuthority = False Then
        Exit Function
    End If
    
    Dim oDM As New dmSuitTables
    Dim cls1 As New clsBL
    Dim strSQL As String
    Dim rs As New RecordSet
    Dim szUser As String
    Dim szTip As String
    
    Set oDM = cls1.GetSuitTables(vObjectID)
    strSQL = "select * from " & oDM.Bill & " where B_ID=" & vID
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        JudgeUpdateBill = False
        MsgBox "在检测是否拥有修改权限时未找到对应的单据！", vbOKOnly + vbInformation, "提示"
        Exit Function
    End If
    
    szUser = IIf(IsNull(rs!B_UserName), "", rs!B_UserName)
    rs.Close
    Set rs = Nothing
    
    
    If Len(szUser) <= 0 Then
        JudgeUpdateBill = True
        Exit Function
    End If
    
    
    
    If szUser <> Gm.SysID.SystemUser Then
        JudgeUpdateBill = False
        szTip = "该单据由[" & szUser & "]制作" & vbNewLine
        szTip = szTip & "与当前登录用户[" & Gm.SysID.SystemUser & "]不匹配" & vbNewLine
        szTip = szTip & "不可修改！"
        MsgBox szTip, vbOKOnly + vbInformation, "提示"
    Else
        JudgeUpdateBill = True
    End If
    
End Function

Public Function JudgeDelete(ByVal vObjectID As String) As Boolean
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim a As String
    a = Gm.SysID.SystemUser
    sql = "select * from G_systemuser where B_UserName='" & a & "'"
    rs1.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs1!B_SuperAdmin = 1 Then
        JudgeDelete = True
        Exit Function
    End If
    If Gm.IsDeveloper = True Then
        JudgeDelete = True
        Exit Function
    End If
    
    
    '不使用权限功能
    If Gm.DBCfgOneInt.UsingFineGrainedAuthority = False Then
        JudgeDelete = True
        Exit Function
    End If
    
    
    If JudgePermission(vObjectID, "B_Delete") = False Then
        JudgeDelete = False
        MsgBox "当前用户没有删除权限！", vbOKOnly + vbInformation, "提示"
    Else
        JudgeDelete = True
    End If
End Function

'判断是否有查看的权限
Public Function JudgeView(ByVal vObjectID As String) As Boolean
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim a As String
    a = Gm.SysID.SystemUser
    sql = "select * from G_systemuser where B_UserName='" & a & "'"
    rs1.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs1!B_SuperAdmin = 1 Then
        JudgeView = True
        Exit Function
    End If
    If Gm.IsDeveloper = True Then
        JudgeView = True
        Exit Function
    End If
    
    
    '不使用权限功能
    If Gm.DBCfgOneInt.UsingFineGrainedAuthority = False Then
        JudgeView = True
        Exit Function
    End If
    
    
    A_rs.Filter = " B_ObjectID='" & vObjectID & "'"
    If A_rs.RecordCount > 0 Then
        JudgeView = True
    Else
        JudgeView = False
    End If
    
    A_rs.Filter = ""
End Function




Public Function DestructionCls()
    A_rs.Close
    Set A_rs = Nothing
End Function

'遍历当前打开的所有窗体
Public Sub RefreshFrmsInCache()
    Dim oFrm As Form
    For Each oFrm In Forms
        If Left$(oFrm.name, 12) = "frmNavigator" Then
        If oFrm.name <> "frmNavigatorLeft" Then
            Gm.PI.RefreshPermission
            oFrm.ConfirmPermission
        End If
        End If
    Next
End Sub

Public Function authenticate(ByVal theid As String) As Boolean
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    sql = "Select * From G_SystemUser where B_UserName='" & Gm.SysID.SystemUser & "'"
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs!B_SuperAdmin = 1 Then
        authenticate = True
        Exit Function
    End If
    sql1 = "select * from G_BillOrder where B_ID='" & theid & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs1.RecordCount <= 0 Then
        authenticate = True
        Exit Function
    End If
    If rs1.RecordCount > 0 Then
        If rs1!B_Audit = 0 Then
            If rs1!B_UserName = Gm.SysID.SystemUser Then
                    authenticate = True
                Else
                    authenticate = False
                    MsgBox "不是本制单人不能修改", vbInformation, "提示"
                    Exit Function
                End If
                If DateDiff("s", rs1!B_datecreate, Now) > 84600 Then
                    authenticate = False
                    MsgBox "已经超过制作本单据的时间不能进行修改", vbInformation, "提示"
                Else
                    authenticate = True
                End If
        Else
            MsgBox "数据已经审核，不能进行操作", vbInformation, "提示"
             authenticate = False
        End If
    End If
End Function


