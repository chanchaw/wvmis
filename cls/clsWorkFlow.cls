VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const A_TableFlow As String = "G_WF_Flow"
Private Const A_TableAction As String = "G_WF_Action"
Private Const A_TableStep As String = "G_WF_Step"
Private Const A_ProgressInstance As String = "G_WF_ProgressInstance"
Private strSQL As String
Private A_dmSuitTable As dmSuitTables
Private A_ObjectID As String

Public Enum Action
    Save = 1
    Audit = 2
    toPrint = 3
End Enum

'工作流类
'本工作流类是紧密结合当前开发环境中的架构（单据对象）
'进行流程控制的

Public Sub initCls(ByVal vObjectID As String)
    A_ObjectID = vObjectID
    
    Dim rs As New RecordSet
    strSQL = "SELECT * FROM G_BL WHERE B_ObjectID='" & vObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Set A_dmSuitTable = New dmSuitTables
    With A_dmSuitTable
        .Bill = rs!B_MainTable
        .Detail = rs!B_DetailTable
        .DraftBill = rs!B_DraftMainTable
        .DraftDetail = rs!B_DraftDetailTable
    End With
    
    rs.Close
    Set rs = Nothing
End Sub

'判断当前操作是否工作流中NEXT步骤
'是或者不存在工作流返回TRUE，否则返回FALSE
'传入参数：单据主表的B_ID
Public Function IsTheProgress(ByVal vID As Long, ByVal vlAction As Long) As Boolean
    '如果单据不存在工作流，那么直接返回TRUE
    If IsInProgress = False Then
        IsTheProgress = True
        Exit Function
    End If
    
    Dim szAction As String
    szAction = GetAction(vlAction)
    
    
    '即将要做的操作的StepID
    Dim rs As RecordSet
    Dim lStepID As Long
    Dim szTemp As String
    Set rs = New RecordSet
    strSQL = "SELECT * FROM G_WF_Step WHERE B_FlowID='" & A_ObjectID & "' AND B_ActionID='" & szAction & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        IsTheProgress = False
        rs.Close
        Set rs = Nothing
        
        szTemp = "为" & A_ObjectID & "开启了工作流，但是没有为[" & szAction & "]设置步骤！"
        MsgBox szTemp, vbOKOnly + vbInformation, "提示"
        Exit Function
    End If
    
    lStepID = IIf(IsNull(rs!B_FlowOrder), 1, rs!B_FlowOrder)
    If lStepID = 1 Then
        IsTheProgress = True '是第一个步骤，不需要判断
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    
    '从工作流实例表中查询获取当前实例的最后一个动作
    '如果大于等于lStepID - 1，那么当前动作被允许
    lStepID = lStepID - 1
    Dim rsInstance As RecordSet
    Set rsInstance = New RecordSet
    strSQL = "SELECT top 1 * FROM G_WF_ProgressInstance "
    strSQL = strSQL & " WHERE B_ProgressID='" & vID & "' "
    strSQL = strSQL & " order by B_ID DESC"
    Debug.Print strSQL
    rsInstance.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rsInstance.RecordCount <= 0 Then
        IsTheProgress = False
        szTemp = IIf(IsNull(rs!B_Tip), "", rs!B_Tip)
        MsgBox szTemp, vbOKOnly + vbInformation, "提示"

    End If
    
    If rsInstance!B_StepID < lStepID Then
        IsTheProgress = False
        szTemp = GetActionTip(lStepID)
        MsgBox szTemp, vbOKOnly + vbInformation, "提示"
    End If
    
    IsTheProgress = True
    rs.Close
    Set rs = Nothing
    rsInstance.Close
    Set rsInstance = Nothing
End Function

'判断当前单据是否有工作流
'判断单据是否有工作流
Private Function IsInProgress() As Boolean
    Dim rs As RecordSet
    strSQL = "SELECT * FROM G_WF_Flow WHERE B_SID='" & A_ObjectID & "'"
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        IsInProgress = False
    Else
        IsInProgress = True
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function GetAction(ByVal vEM As Long) As String
    Dim szReturn As String
    Select Case vEM
        Case 1
            szReturn = "Save"
        Case 2
            szReturn = "Audit"
        Case 3
            szReturn = "Print"
    End Select
    
    GetAction = szReturn
End Function


Private Function GetActionTip(ByVal vStepID As Long) As String
    Dim rs As New RecordSet
    strSQL = "SELECT * FROM G_WF_Step WHERE B_FlowID='" & A_ObjectID & "'"
    strSQL = strSQL & " And B_FlowOrder=" & vStepID
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        GetActionTip = "上级步骤没有完成！"
    Else
        GetActionTip = IIf(IsNull(rs!B_Tip), "", rs!B_Tip)
    End If
    rs.Close
    Set rs = Nothing
End Function

'将当前操作写入实例表
Public Function WriteInstance(ByVal vID As Long, ByVal vlAction As Long) As Boolean
    strSQL = "Select * From G_WF_ProgressInstance where 1=0"
    Dim rs As RecordSet
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    rs.AddNew
    rs!B_UserName = Gm.SysID.SystemUser
    rs!B_ProgressID = vID
    rs!B_StepID = GetActionStepID(vlAction)
    rs!B_Date = Format(Now, "YYYY-MM-DD HH:MM:SS")
    rs.Update
    rs.Close
    Set rs = Nothing
End Function

Private Function GetActionStepID(ByVal vlAction As Long) As Long
    Dim szAction As String
    szAction = GetAction(vlAction)
    
    strSQL = "SELECT * FROM G_WF_Step WHERE B_FlowID='" & A_ObjectID & "' AND B_ActionID='" & szAction & "'"
    Dim rs As RecordSet
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        GetActionStepID = 0
    Else
        GetActionStepID = IIf(IsNull(rs!B_FlowOrder), 0, rs!B_FlowOrder)
    End If
    
    rs.Close
    Set rs = Nothing
End Function

