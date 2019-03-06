VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoCreateBL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'���������Զ�����һ�����ݶ���(���������п�ͷΪ12�ĵ��ݶ���)
'������ĳ���ʹ�ü���ӯ��֯

Private A_ObjectID As String   '���������ɵĵ��ݶ���Ķ�����
Private A_TNBill As String     '��ʽ��������
Private A_TNDetail As String   '��ʽ��ϸ������
Private A_TNDraftBill As String '�ݸ���������
Private A_TNDraftDetail As String '�ݸ���ϸ������
Private A_BID As String    '���ݱ�ŵ�Ӣ��ǰ׺
Private A_ID As Long   '������������B_ID
Private A_strItemID As String   'ͨ��CreateOneDetail��ͨ��CreateDetailsByRs�½�����ϸ���ݵ�B_ItemID
                                '�������ӵ����ַ�����

Private strSQL As String
Private A_Inited As Boolean '��ʼ�����


'��ȡ�ױ��Լ�������ʵ����
Public Function GetSuitTable() As dmSuitTables
    Dim o As dmSuitTables
    With o
        .Bill = A_TNBill
        .DraftBill = A_TNDraftBill
        .Detail = A_TNDetail
        .DraftDetail = A_TNDraftDetail
    End With
    
    Set GetSuitTable = o
End Function


'��ʼ������
Public Sub InitCls(ByVal vObjectID As String)
    A_strItemID = ""
    A_ObjectID = vObjectID
    GetObjectPara
    
    A_Inited = True
End Sub

'��ȡĿ������һЩ����
Private Sub GetObjectPara()
    Dim rs As RecordSet
    strSQL = "Select * From G_BL Where B_ObjectID='" & A_ObjectID & "'"
    Set rs = New RecordSet
    'rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        A_TNBill = ""
        A_TNDetail = ""
        A_TNDraftBill = ""
        A_TNDraftDetail = ""
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    
    A_TNBill = rs!B_MainTable
    A_TNDetail = rs!B_DetailTable
    A_TNDraftBill = rs!B_DraftMainTable
    A_TNDraftDetail = rs!B_DraftDetailTable
    A_BID = rs!B_BID
    
    rs.Close
    Set rs = Nothing
End Sub


'���ݸ���������Ϊ��ʽ����
Public Sub Update()
    '1.�Ƚ��ݸ��������ݿ�������ʽ����
    strSQL = "Insert Into " & A_TNBill
    strSQL = strSQL & " Select * From " & A_TNDraftBill & " Where B_ID=" & A_ID
    Gm.cnnTool.cnn.Execute strSQL
    
    '�����ݸ�������ʱ��B_Closed=0����ô����֮����Ҫ�޸�Ϊ1
    strSQL = "Update " & A_TNBill & " set B_Closed=1 where B_ID=" & A_ID
    Gm.cnnTool.cnn.Execute strSQL
    
    
    '2.�ٽ��ݸ���ϸ�������ݿ�������ʽ��ϸ��
    If Right$(A_strItemID, 1) = "," Then
        A_strItemID = Left$(A_strItemID, Len(A_strItemID) - 1)
    End If
    
    strSQL = "Insert Into " & A_TNDetail
    strSQL = strSQL & " Select * From " & A_TNDraftDetail & " Where B_ItemID In"
    strSQL = strSQL & " (" & A_strItemID & ")"
    Debug.Print strSQL
    Gm.cnnTool.cnn.Execute strSQL
    
    
    '3.ɾ���ݸ��������
    strSQL = "Delete From " & A_TNDraftDetail & " Where B_ItemID In (" & A_strItemID & ")"
    Gm.cnnTool.cnn.Execute strSQL
    
    strSQL = "Delete From " & A_TNDraftBill & " Where B_ID=" & A_ID
    Gm.cnnTool.cnn.Execute strSQL
    
End Sub

'�����ɺ�,�Զ��򿪳����뵥
Private Sub OpenBL(ByVal vObjectID As String, ByVal vBillName As String, ByVal vID As String)
    Dim clsCommand1 As New clsCommand
    clsCommand1.InitClass
    clsCommand1.Execute vObjectID, vBillName, "LoadObject", Nothing, vID
End Sub

Public Sub openbill()
    OpenBL A_ObjectID, "����", A_ID
End Sub

'����һ���ݸ�����
'vData���������ַָ��������ݡ����ƣ�B_ClientID="0001",B_Memo="��ʱ��ⵥ"
'       ��һ���ָ������ÿ���ֶΣ��ڶ����ָ�������ֶ����ƺ�VALUE������Щ����д���½�������
'������ݳ�ʼ��ʱ�ṩ��A_ObjectID��ȡ�������У�B_ObjectID��B_CodeID��B_Date����ǰʱ�䣩��B_BID��B_Closed
Public Function CreateOneBill(ByVal vData As String, ByVal vSymFilter01 As String, ByVal vSymFilter02 As String) As Long
    Dim i As Long
    Dim rs As New RecordSet
    Set rs = String2Recordset(vData, vSymFilter01, vSymFilter02)
    
    Dim dtDate As String
    Dim rsBill As RecordSet
    Set rsBill = New RecordSet
    strSQL = "Select * From " & A_TNDraftBill & " Where 1=0"
    rsBill.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    dtDate = Format(Now, "YYYY-MM-DD HH:MM:SS")
    
    
    If rs.RecordCount > 0 Then
        '�����µ�һ���ݸ�����֮ǰ�������2������
        A_strItemID = ""
        A_ID = -999999
        
        rsBill.AddNew
        rsBill!B_ObjectID = A_ObjectID
        rsBill!B_Codeid = GetFrameCodeDetail(A_ObjectID)
        rsBill!B_Date = dtDate
        rsBill!B_BID = A_BID
        rsBill!B_Closed = 1
            
        rs.MoveFirst
        Do While Not rs.EOF
            rsBill(Trim$(rs!B_Field01)) = rs!B_Field02
            rs.movenext
        Loop
        rsBill.Update
        A_ID = rsBill!B_id
    End If
    
    CreateOneBill = A_ID
    
    rsBill.Close
    Set rsBill = Nothing
End Function

'����һ����¼������һ������
'�ڲݸ�����A_TNDraftBill�д���һ�����ݣ����ˣ�B_ObjectID��B_CodeID��B_Date��B_BID��B_Closed
'ϵͳ�����ֶΣ��������ֶΰ��ղ���vData��д������
'������vData���ֶ��ڲݸ屨����Ѱ�Ҷ�Ӧ���ֶ�д������
Public Function CreateOneBillByRs(ByRef vData As RecordSet) As Long
    Dim i As Long
    Dim rs As New RecordSet
    Set rs = vData.Clone
    
    Dim dtDate As String
    Dim rsBill As RecordSet
    Set rsBill = New RecordSet
    strSQL = "Select * From " & A_TNDraftBill & " Where 1=0"
    rsBill.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    dtDate = Format(Now, "YYYY-MM-DD HH:MM:SS")
    
    
    If rs.RecordCount > 0 Then
        '�����µ�һ���ݸ�����֮ǰ�������2������
        A_strItemID = ""
        A_ID = -999999
        
        rsBill.AddNew
        rsBill!B_ObjectID = A_ObjectID
        rsBill!B_Codeid = GetFrameCodeDetail(A_ObjectID)
        rsBill!B_Date = dtDate
        rsBill!B_BID = A_BID
        rsBill!B_Closed = 0
            
            
            
        'ֻ����һ������
        rs.MoveFirst
        For i = 0 To rs.Fields.Count - 1
            If Not IsNull(rs(i)) Then
                rsBill(Trim$(rs(i).name)) = rs(i)
            End If
        Next
        rsBill.Update
        A_ID = rsBill!B_id
    End If
    
    CreateOneBillByRs = A_ID
    
    rsBill.Close
    Set rsBill = Nothing
End Function

'����һ���ݸ���ϸ,���ر���������ϸ��¼��B_ItemID
'һ�㱾������CreateDetailsByRs����ѡ��һʹ��
Public Function CreateOneDetail(ByVal vData As String, ByVal vSymFilter01 As String, _
    ByVal vSymFilter02 As String) As Long
    
    If A_ID <= 0 Then
        Exit Function
    End If
    
    Dim rs As New RecordSet
    Set rs = String2Recordset(vData, vSymFilter01, vSymFilter02)
    If rs.RecordCount <= 0 Then
        Exit Function
    End If
    
    Dim rsdetail As New RecordSet
    strSQL = "Select * From " & A_TNDraftDetail & " Where 1=0"
    rsdetail.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rsdetail.AddNew
    rsdetail!B_id = A_ID
    
    rs.MoveFirst
    Do While Not rs.EOF
        rsdetail(Trim$(rs!B_Field01)) = rs!B_Field02
        rs.movenext
    Loop
    rsdetail.Update
    A_strItemID = A_strItemID & rsdetail!B_itemid & ","
    CreateOneDetail = rsdetail!B_itemid
    rsdetail.Close
    Set rsdetail = Nothing
End Function


'����һ���ݸ���ϸ���ݡ�һ�㱾������CreateOneDetail����ѡ��һʹ��
'vDataRs����Ҫд������ݡ��ֶζ�ӦDB�еı����ֶ�
'����ֵ�����ش�����N����ϸ��¼��B_ItemID�ַ������м���Ӣ�ĵĶ��ż����
Public Function CreateDetailsByRs(ByVal vDataRs As RecordSet) As String
    If A_ID <= 0 Then
        Exit Function
    End If
    
    If vDataRs.State <> adStateOpen Then
        Exit Function
    End If
    
    If vDataRs.RecordCount <= 0 Then
        Exit Function
    End If
    
    
    Dim i As Long
    Dim rsdetail As New RecordSet
    strSQL = "Select * From " & A_TNDraftDetail & " Where 1=0"
    rsdetail.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic

    
    vDataRs.MoveFirst
    Do While Not vDataRs.EOF
        rsdetail.AddNew
        rsdetail!B_id = A_ID
    
        For i = 0 To vDataRs.Fields.Count - 1
            rsdetail(Trim$(vDataRs(i).name)) = vDataRs(i)
            Debug.Print vDataRs(i).name & "=" & vDataRs(i)
        Next
        rsdetail.Update
        A_strItemID = A_strItemID & rsdetail!B_itemid & ","
        vDataRs.movenext
    Loop
    CreateDetailsByRs = A_strItemID
    rsdetail.Close
    Set rsdetail = Nothing
End Function

'��һ����2���ָ������ַ���ת��Ϊһ����¼��
Private Function String2Recordset(ByVal vData As String, ByVal vSymFilter01 As String, ByVal vSymFilter02 As String) As RecordSet
    Dim rs As New RecordSet
    Dim arrArr01, arrArr02
    Dim i As Long, j As Long
    
    rs.Fields.Append "B_Field01", adVarChar, 100
    rs.Fields.Append "B_Field02", adVarChar, 100
    rs.Open
    
    
    arrArr01 = Split(vData, vSymFilter01)
    For i = 0 To UBound(arrArr01)
        arrArr02 = Split(arrArr01(i), vSymFilter02)
        
        rs.AddNew
        rs(0) = arrArr02(0)
        rs(1) = arrArr02(1)
        rs.Update
        
    Next
    
    Set String2Recordset = rs
End Function


'���ɵ��ݱ��B_CodeID
Public Function GetFrameCodeDetail(ByVal m_ObjectID As String) As String
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
    Debug.Print strSQL
    
    mvarm_BID = A_BID
    m_DraftMainTable = A_TNDraftBill
    m_MainTable = A_TNBill
    
    
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
    Debug.Print strTmpBH

    mstrSQL = "Select CASE WHEN ISNULL(P1.B_CodeID,0)>ISNULL(P2.B_CodeID,0) THEN P1.B_CodeID"
    mstrSQL = mstrSQL & " Else P2.B_CodeID End as B_PCodeID"
    mstrSQL = mstrSQL & " From (Select Max(B_CodeID) as B_CodeID"
    mstrSQL = mstrSQL & " From " & m_DraftMainTable
    mstrSQL = mstrSQL & " Where B_CodeID Like '" & Trim(strTmpBH) & "%') as P1,"

    mstrSQL = mstrSQL & " (Select Max(B_CodeID) as B_CodeID"
    mstrSQL = mstrSQL & " From " & m_MainTable
    mstrSQL = mstrSQL & " Where B_CodeID Like '" & Trim(strTmpBH) & "%') as P2"
    Debug.Print mstrSQL
    rs.Open mstrSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    If IsNull(rs(0)) Then
        '���û����ǰ�ļ�¼
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

'�Ѿ����ʵĵ���ͨ����������������ʽ����׷��һ����ϸ��¼
'Private Function AppendDraftDetailByRs(ByVal vID As Long, ByVal vRs As Recordset) As Recordset
'    Dim rtnRs As Recordset
'    Set rtnRs = New Recordset
'    rtnRs.Fields.Append "B_ItemID", adInteger
'    rtnRs.Open
'
'
'
'    Dim i As Long
'    Dim rsDetail As New Recordset
'    strSQL = "Select * From " & A_TNDraftDetail & " Where 1=0"
'    rsDetail.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
'
'
'    vRs.MoveFirst
'    Do While Not vRs.EOF
'        rsDetail.AddNew
'        rsDetail!B_ID = vID
'
'        For i = 0 To vRs.Fields.Count - 1
'            rsDetail(Trim$(vRs(i).name)) = vRs(i)
'            'Debug.Print vRs(i).name & "=" & vRs(i)
'        Next
'        rsDetail.Update
'
'
'        '�����������е�B_ItemID
'        rtnRs.AddNew
'        rtnRs!B_ItemID = rsDetail!B_ItemID
'
'        vRs.MoveNext
'    Loop
'    rsDetail.Close
'    Set rsDetail = Nothing
'
'
'    Set AppendDraftDetailByRs = rtnRs.Clone
'End Function
'
'
'Public Function AppendDetailByRs(ByVal vID As Long, ByVal vRs As Recordset) As Boolean
'
'    If A_Inited = False Then
'        AppendDetailByRs = False
'        MsgBox "��ʼ����ſ���ʹ�ã�", vbOKOnly + vbInformation, "��ʾ"
'        Exit Function
'    End If
'
'    Dim rs As Recordset
'    Set rs = AppendDraftDetailByRs(vID, vRs).Clone
'
'
'End Function
'