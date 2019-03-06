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

'��������
'�����������ǽ��ܽ�ϵ�ǰ���������еļܹ������ݶ���
'�������̿��Ƶ�

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

'�жϵ�ǰ�����Ƿ�������NEXT����
'�ǻ��߲����ڹ���������TRUE�����򷵻�FALSE
'�������������������B_ID
Public Function IsTheProgress(ByVal vID As Long, ByVal vlAction As Long) As Boolean
    '������ݲ����ڹ���������ôֱ�ӷ���TRUE
    If IsInProgress = False Then
        IsTheProgress = True
        Exit Function
    End If
    
    Dim szAction As String
    szAction = GetAction(vlAction)
    
    
    '����Ҫ���Ĳ�����StepID
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
        
        szTemp = "Ϊ" & A_ObjectID & "�����˹�����������û��Ϊ[" & szAction & "]���ò��裡"
        MsgBox szTemp, vbOKOnly + vbInformation, "��ʾ"
        Exit Function
    End If
    
    lStepID = IIf(IsNull(rs!B_FlowOrder), 1, rs!B_FlowOrder)
    If lStepID = 1 Then
        IsTheProgress = True '�ǵ�һ�����裬����Ҫ�ж�
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    
    '�ӹ�����ʵ�����в�ѯ��ȡ��ǰʵ�������һ������
    '������ڵ���lStepID - 1����ô��ǰ����������
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
        MsgBox szTemp, vbOKOnly + vbInformation, "��ʾ"

    End If
    
    If rsInstance!B_StepID < lStepID Then
        IsTheProgress = False
        szTemp = GetActionTip(lStepID)
        MsgBox szTemp, vbOKOnly + vbInformation, "��ʾ"
    End If
    
    IsTheProgress = True
    rs.Close
    Set rs = Nothing
    rsInstance.Close
    Set rsInstance = Nothing
End Function

'�жϵ�ǰ�����Ƿ��й�����
'�жϵ����Ƿ��й�����
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
        GetActionTip = "�ϼ�����û����ɣ�"
    Else
        GetActionTip = IIf(IsNull(rs!B_Tip), "", rs!B_Tip)
    End If
    rs.Close
    Set rs = Nothing
End Function

'����ǰ����д��ʵ����
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
