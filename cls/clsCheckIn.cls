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

'----�¶�������
'----������:�³�
'----���ݱ��
Public m_ID As Long
'----�������
Public OK As Boolean

Public Event OnCheckIn(ByRef oCnn As Connection)

Dim m_MainTable As String                   '��������
Dim m_DetailTable As String                 '������ϸ��
Dim m_DraftMainTable As String              '�ݸ�����
Dim m_DraftDetailTable As String            '�ݸ�

Dim m_MainFieldList As String
Dim m_DetailFieldList As String

Dim rsStatus As New Recordset
'��������ֵ�ľֲ�����
Private mvarObjectID As String '�ֲ�����
Public Property Let ObjectID(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ObjectID = 5
    mvarObjectID = vData
End Property


Public Property Get ObjectID() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
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

    '��ȡ�ֶ���
    GetBillFieldList oCnn
    
    '����B_Closed=1(Ϊ����״̬)
    CloseBill oCnn
    
    
    '���ݸ����ݵ��ʵ���ʽ����
    AddBillEx oCnn
    
    'ɾ���ݸ������
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

'----����
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

    '��ȡ�ֶ���
    GetBillFieldList oCnn
    
    '����B_Closed=1(Ϊ����״̬)
    CloseBill oCnn
    
    '���ݸ���е����ݵ��ʵ���ʽ����
    AddBill oCnn
    
    'ɾ���ݸ������
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

'----��������
Private Sub CloseBill(ByRef oCnn As Connection)
    Dim strSQL As String
    strSQL = "Update " & m_DraftMainTable & " Set B_Closed=1 Where B_ID=" & m_ID
    oCnn.Execute strSQL
End Sub

'----ȡ�õ���״̬
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

'----���ӱ�������
'----������ϸ������
Private Sub AddBillEx(ByRef oCnn As Connection)
    Dim strSQL As String
    Dim rs As Recordset
    
        '��ϸ������
    '=================
    Set rs = New Recordset
    strSQL = "Select * From G_RuleBillDetail Where B_ObjectID='" & mvarObjectID & "' Order by B_Order"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    '=================
    
    
    
    '----������������
    strSQL = "Insert Into " & m_MainTable
    strSQL = strSQL & " (" & m_MainFieldList & ")"
    strSQL = strSQL & " Select " & m_MainFieldList & " From " & m_DraftMainTable
    strSQL = strSQL & " Where B_ID=" & m_ID
    Debug.Print strSQL
    oCnn.Execute strSQL

    '----������ϸ������
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

'----���ӱ�������
'----������ϸ������

'�Ӳݸ���������ݵ���ʽ��
'������ѡ���ԵĿ���(��ʵ�ֹ���)
Private Sub AddBill(ByRef oCnn As Connection)
    Dim strSQL As String
    Dim rs As Recordset
        
    
    '��ϸ������
    '=================
    Set rs = New Recordset
    strSQL = "Select * From G_RuleBillDetail Where B_ObjectID='" & mvarObjectID & "' Order by B_Order"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    '=================
    
    
    
    '----������������
    strSQL = "Insert Into " & m_MainTable
    strSQL = strSQL & " (" & m_MainFieldList & ")"
    strSQL = strSQL & " Select " & m_MainFieldList & " From " & m_DraftMainTable
    strSQL = strSQL & " Where B_ID=" & m_ID

    oCnn.Execute strSQL
    
    
    
    '����ĵ������òֿ�ĵ��ݱ���ʱ������B_ObjectID��һ��,�ڴ�����
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
    
    
    
    

    '----������ϸ������
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
    '----����Ϊ�������SQL
    'Insert Into MainTable (Fields...) Select (Fields...) From MainTable Where B_ID=m_ID
    Dim rs1 As New Recordset        '----�ݸ��¼
    Dim rs2 As New Recordset        '----�ʼ�¼
    Dim strSQL As String
    Dim i As Integer
    
    Set rs1 = New Recordset
    Set rs2 = New Recordset
    
    '----��ѯ�����ֶ�
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
    
    '----��ѯ��ϸ���ֶ�
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

'----ɾ��ԭ��ϸ��
'----ɾ��ԭ��
Private Sub DeleteOldBillList(ByRef oCnn As Connection)
    Dim strSQL As String
    strSQL = "Delete From " & m_DraftDetailTable & " Where B_ID=" & m_ID
    Debug.Print strSQL
    oCnn.Execute strSQL
    
    strSQL = "Delete From " & m_DraftMainTable & " Where B_ID=" & m_ID
    Debug.Print strSQL
    oCnn.Execute strSQL
End Sub

'----�ж��������Ƿ�������
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
        MsgBox rs("B_ShowChinese"), vbExclamation, "����"
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

