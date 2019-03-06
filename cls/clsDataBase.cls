VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'���ഴ����2013��12��26�� 22:42:30

Private strSQL As String
Private A_rsTable As New RecordSet

Private A_LogFile As String
Private A_Logger As New clsFile

'��ʼ������ĺ���
'��ȡ��ǰ���ݿ��µ������û���������
Public Sub initCls(ByVal vDBName As String)
    Set A_rsTable = New RecordSet
    strSQL = "Select * FROM " & vDBName & "..SysObjects Where XType='U' Order By Name"
    A_rsTable.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    
End Sub

'�ж�һ�����ݱ��ڵ�ǰ�����ݿ����Ƿ����
'���ڷ���TRUE,���򷵻�FALSE
Public Function JudgeTableExist(ByVal vTableName As String) As Boolean
    Dim strTip As String
    
    JudgeTableExist = True
    
    If A_rsTable.State <> adStateOpen Then
        JudgeTableExist = False
        strTip = "���ݿ��û�����¼��δ��������!" & vbNewLine
        strTip = strTip & "����ȷ�����Ƿ�Ա������˳�ʼ��!"
        MsgBox strTip, vbOKOnly + vbInformation, "��ʾ"
        Exit Function
    End If
    
    If A_rsTable.RecordCount <= 0 Then
        JudgeTableExist = False
        strTip = "���ݿ��û���������Ϊ0!" & vbNewLine
        strTip = strTip & "����ȷ�����Ƿ�Ա������˳�ʼ��!"
        MsgBox strTip, vbOKOnly + vbInformation, "��ʾ"
        Exit Function
    End If
    
    A_rsTable.Filter = " Name='" & vTableName & "'"
    If A_rsTable.RecordCount > 0 Then
        JudgeTableExist = True
    Else
        JudgeTableExist = False
    End If
    
    A_rsTable.Filter = ""
End Function

'�ж�һ�������Ƿ����ĳ�ֶ�
Public Function JudgeFieldExist(ByVal vTableName As String, ByVal vFieldName As String) As Boolean
    Dim rs As New RecordSet
    
    '�����������ڵ�ʱ��,ֱ���˳�
    If JudgeTableExist(vTableName) = False Then
        JudgeFieldExist = False
        Exit Function
    End If
    
    
    Set rs = New RecordSet
    strSQL = "Select * FROM SysColumns Where id=Object_Id('" & vTableName & "')"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    rs.Filter = " Name='" & vFieldName & "'"
    If rs.RecordCount > 0 Then
        JudgeFieldExist = True
    Else
        JudgeFieldExist = False
    End If
    rs.Close
    Set rs = Nothing
    
End Function


'�ж�һ���ַ����е������ֶ��Ƿ������Ŀ�����
'�����ڵ������ֶζ�������ʾ
Public Function JudgeFieldsExist(ByVal vTableName As String, ByVal vFieldNames As String, ByVal vSymbol As String) As Boolean
    Dim rs As New RecordSet
    Dim szErrTip As String
    Dim curFieldName As String
    
    JudgeFieldsExist = True
    
    '�����������ڵ�ʱ��,ֱ���˳�
    If JudgeTableExist(vTableName) = False Then
        MsgBox "ָ���ı�" & vTableName & "�����ڣ�", vbOKOnly + vbInformation, "��ʾ"
        JudgeFieldsExist = False
        Exit Function
    End If
    
    
    '��û�����ñ������ֶ�ʱ�򷵻�TRUE
    If Len(Trim$(vFieldNames)) <= 0 Then
        MsgBox "û�����ñ������ֶΣ�", vbOKOnly + vbInformation, "��ʾ"
        JudgeFieldsExist = False
        Exit Function
    End If
    
    
    '��ȡָ�����µ������ֶ�
    Set rs = New RecordSet
    strSQL = "Select * FROM SysColumns Where id=Object_Id('" & vTableName & "')"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

    If rs.RecordCount <= 0 Then
        MsgBox "ָ���ı��²������κ��ֶΣ�", vbOKOnly + vbInformation, "��ʾ"
        rs.Close
        Set rs = Nothing
        JudgeFieldsExist = False
        Exit Function
    End If
        
    
    Dim szArr
    Dim i As Long
    szArr = Split(vFieldNames, vSymbol)
    
    
    
    'ѭ������ֶ��Ƿ����
    szErrTip = ""
    For i = 0 To UBound(szArr)
        curFieldName = Trim(szArr(i))
        rs.Filter = " Name='" & curFieldName & "'"
        
        If rs.RecordCount <= 0 Then
            szErrTip = szErrTip & curFieldName & vbNewLine
            JudgeFieldsExist = False
        End If
    Next

    '����в����ڵ��ֶ���ô������ʾ
    If Len(Trim$(szErrTip)) > 0 Then
        szErrTip = "��" & vTableName & "�в������ֶΣ�" & vbNewLine & szErrTip
        MsgBox szErrTip, vbOKOnly + vbInformation, "��ʾ"
    End If
    
    rs.Close
    Set rs = Nothing
End Function

'vNoIdentity:Ĭ�ϲ�������������
Public Function GetTableFields(ByVal vTableName As String, _
    Optional ByVal vNoIdentity As Boolean = True) As String
    
    If vNoIdentity = True Then
        strSQL = "SELECT * FROM syscolumns AS s WHERE s.id=OBJECT_ID('" & vTableName & "') AND COLUMNPROPERTY(s.id,s.name,'IsIdentity')<>1"
    Else
        strSQL = "SELECT * FROM syscolumns AS s WHERE s.id=OBJECT_ID('" & vTableName & "')"
    End If
    Debug.Print strSQL
    Dim rs As New RecordSet
    Dim cls1 As New clsRecordset
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    GetTableFields = cls1.RecordSetToString(rs, "name", ",")
    
    rs.Close
    Set rs = Nothing
End Function

'ִ���ⲿSQL�ű��ļ�
Public Function ExecSQLScript(ByVal vPath As String) As Boolean
    ExecSQLScript = True
    If Len(Trim$(vPath)) <= 0 Then
        ExecSQLScript = False
        MsgBox "ָ����SQL�ű�����·�������ڣ�", vbOKOnly + vbInformation, "��ʾ"
        Exit Function
    End If
    
    A_LogFile = App.Path & "\�����������־.txt"
    A_Logger.WriteFileContent A_LogFile, "��" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "��ʼд�����������־"
    A_Logger.WriteFileContent A_LogFile, "================================================================"



    If Right$(vPath, 1) = "\" Or Right$(vPath, 1) = "/" Then
        vPath = Left$(vPath, Len(vPath) - 1)
    End If
    
    '�޸ĵ�¼���ݿ���û���Ĭ�����ݿ�
    ChangeDBUserDefaultDBCur
     
    Dim rsFile As RecordSet
    Dim clsFile1 As clsFile
    
    Set clsFile1 = New clsFile
    Set rsFile = New RecordSet
    
    '��ȡSQL�����ļ����µ������ļ��ľ���·��
    Set rsFile = clsFile1.TraversalFilesReturnRS(vPath)
    
    
    '����¼�����򣬰���˳��ִ��SQL
    If rsFile.RecordCount <= 0 Then
        rsFile.Close
        Set rsFile = Nothing
        Exit Function
    End If
    
    
    rsFile.Sort = " B_Field0 ASC"
    rsFile.MoveFirst
    
    
    
    Dim rs As New RecordSet
    Do While Not rsFile.EOF
        '����ļ�ִ��SQL
        Set rs = New RecordSet
        strSQL = "master.dbo.xp_cmdshell ' osql -U " & Gm.SysID.DBInfo.DBUser & " -P " & Gm.SysID.DBInfo.DBPassWord & " -i """ & rsFile!B_Field0 & """'"
        rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        WriteLog rs
        
        rs.Close
        Set rs = Nothing
    
        rsFile.MoveNext
    Loop
    
    rsFile.Close
    Set rsFile = Nothing
    
End Function

'�޸����ݿ��û�Ĭ�ϵ����ݿ�
Public Sub ChangeDBUserDefaultDB(ByVal vUser As String, ByVal vDBName As String)
    strSQL = "sp_defaultdb '" & vUser & "','" & vDBName & "'"
    Gm.cnnTool.cnn.Execute strSQL
End Sub

'�޸ĵ�ǰ�����ļ���ָ�������ݿ��¼�û�
'��Ĭ�����ݿ�Ϊ�����ļ���ָ�������ݿ�
Public Sub ChangeDBUserDefaultDBCur()
    ChangeDBUserDefaultDB Gm.SysID.DBInfo.DBUser, Gm.SysID.DBInfo.DBName
End Sub



Private Sub WriteLog(ByRef vRs As RecordSet)
    If vRs.State <> adStateOpen Then
        Exit Sub
    End If
    
    If vRs.RecordCount <= 1 Then
        Exit Sub
    End If
    

    Dim szContent As String
    Do While Not vRs.EOF
        szContent = IIf(IsNull(vRs(0)), "", vRs(0))
        If Len(szContent) > 0 Then
            A_Logger.WriteFileContent A_LogFile, szContent
        End If
        vRs.MoveNext
    Loop
    
End Sub