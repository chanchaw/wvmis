VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private strSQL As String


'��һ����¼���е���������Ϊһ���ı��ĵ�
Public Sub CreateRsFile(ByRef oRs As RecordSet)
    Dim rs As RecordSet
    Dim strLineData As String
    Dim i As Long
    Dim strFilter As String
    
    strFilter = "  |  "
    
    
    If oRs.State <> adStateOpen Then
        MsgBox "Ŀ���¼��δ�򿪣�", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If oRs.RecordCount <= 0 Then
        MsgBox "Ŀ���¼������Ŀ��Ϊ�գ�", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    '1.��¡Դ��¼�����������ڵĵ���ʱ��¼����
    Set rs = oRs.Clone
    
    
    
    '2.����һ���ı��ĵ�
    Dim fso As Object
    Dim strFileName As String
    strFileName = "��¼��" & Format(Now, "YYYY-MM-DD HHMMSS") & ".txt"
    strFileName = App.Path & "\" & strFileName
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateTextFile strFileName
    
    
    '3.д���һ���ֶ�����
    strLineData = ""
    For i = 0 To rs.Fields.Count - 1
        strLineData = strLineData & rs(i).name & strFilter
    Next
    strLineData = Trim$(strLineData)
    strLineData = Left$(strLineData, Len(strLineData) - 1)
    WriteFileContent strFileName, strLineData
    
    
    '4.д�����µ�����
    rs.MoveFirst
    Do While Not rs.EOF
        strLineData = ""
        For i = 0 To rs.Fields.Count - 1
            strLineData = strLineData & rs(i).Value & strFilter
        Next
        WriteFileContent strFileName, strLineData
        rs.movenext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    
    '5.�����洴�����ı��ĵ�
    Dim a&
    a = ShellExecute(0, "open", strFileName, 0, 0, 1)
End Sub


'���ı��ĵ�׷��д���ı���������ļ��������򴴽�����д�룩
'strFileName���ļ���������·��
'strWrite��Ҫд�������
Private Sub WriteFileContent(ByVal strFileName As String, ByVal strWrite As String)
    Dim lngHandle As Long '���
    lngHandle = FreeFile() 'ȡ�þ��
    

    Open strFileName For Append As lngHandle    '���ļ�
    Print #lngHandle, strWrite    'д���ı�
    Close lngHandle    '�ر��ļ�
End Sub

'��һ����¼���е�����д�뵽һ������
Public Function WriteRsToDataBase(ByVal vTableName As String, ByRef vRs As RecordSet) As Boolean
    Dim rs1 As RecordSet
    
    If Len(Trim$(vTableName)) <= 0 Then
        WriteRsToDataBase = False
        Exit Function
    End If
    
    If vRs.State <> adStateOpen Then
        WriteRsToDataBase = False
        Exit Function
    End If
    
    If vRs.RecordCount <= 0 Then
        WriteRsToDataBase = False
        Exit Function
    End If
    
    Set rs1 = New RecordSet
    Set rs1 = vRs.Clone
    
    Dim rs2 As New RecordSet
    
    strSQL = "Select * From " & vTableName
    Set rs2 = New RecordSet
    rs2.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    
    rs1.MoveFirst
    Do While Not rs1.EOF
        rs2.AddNew
        
        rs1.movenext
    Loop
End Function

'����������WriteRsToDataBase�����Ƿ�Public
Private Function GetTableRsCorrespond(ByVal vTableName As String) As RecordSet
    Dim rs As RecordSet
    Set rs = New RecordSet
    rs.Fields.Append "B_Field01", adVarChar
    rs.Fields.Append "B_Field02", adVarChar
    
    
End Function


Public Sub BuildRsFields(ByRef vRs As RecordSet)
    Dim rs As New RecordSet
    Dim rs1 As New RecordSet
    Dim rs2 As New RecordSet
    Dim lID01 As Long   '��G_BuildRsFieldBill��B_ID
    Dim lID02 As Long   '��G_BuildRsFieldDetail_Section��B_ItemID
    Dim lID03 As Long   '��G_BuildRsFieldDetail��B_ItemID
    
    
    strSQL = "Select * From G_BuildRsFieldBill Where B_GroupName='�ⷢ��������'"
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    lID01 = rs!B_id
    rs.Close
    Set rs = Nothing
    
    
    strSQL = "Select * From G_BuildRsFieldDetail_Section Where B_ID=" & lID01 & " Order By B_Order"
    Debug.Print strSQL
    Set rs1 = New RecordSet
    rs1.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs1.RecordCount <= 0 Then
        rs1.Close
        Set rs1 = Nothing
        Exit Sub
    End If
    
    
    Set vRs = New RecordSet
    Do While Not rs1.EOF
        strSQL = "Select * From G_BuildRsFieldDetail Where B_ID=" & rs1!B_ItemID & " Order By B_Order"
        Debug.Print strSQL
        Set rs2 = New RecordSet
        rs2.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
        
        Do While Not rs2.EOF
            'ʹ�õ�ǰ���ݱ�G_BuildRsFieldDetail�����е��ֶν�������¼�����ֶ�
            If rs2!B_CurData = 1 Then
                vRs.Fields.Append rs2!B_FieldName, adVarChar, 100
            End If
            
            'ʹ���ⲿ���ݱ��е�ĳ�ֶ��ϵ�һϵ��ֵ�������
            If rs2!B_CurData = 0 Then
                Set rs = New RecordSet
                strSQL = "Select * From " & rs2!B_ExtendTable & " Where 1=1 Order By B_Order"
                Debug.Print strSQL
                rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
                
                Do While Not rs.EOF
                    vRs.Fields.Append rs(Trim$(rs2!B_ExtendField)), adVarChar, 100
                    rs.movenext
                Loop
                rs.Close
                Set rs = Nothing
            End If
            rs2.movenext
        Loop
        rs2.Close
        Set rs2 = Nothing
        
        rs1.movenext
    Loop
    rs1.Close
    Set rs1 = Nothing
    
    vRs.Open
End Sub


'��һ����¼��ת��Ϊһ���ַ���
'��strFilter���ȴ���0ʱ��ÿ����¼���ø��ַ������
'����Ļ��м䲻�����
'strFieldNameΪĿ���ֶε��ֶ���
Public Function RecordSetToString(ByRef rs As RecordSet, ByVal strFieldName As String, ByVal strFilter As String) As String
    Dim str As String
       
    str = ""
    rs.MoveFirst
    Do While Not rs.EOF
        If Len(Trim(strFilter)) > 0 Then
            str = str & rs(strFieldName) & strFilter
        Else
            str = str & rs(strFieldName)
        End If
        rs.movenext
    Loop
       
    If Len(Trim(strFilter)) > 0 Then
        str = Left(str, Len(str) - 1)
    End If
       
    '�γ��ַ����󣬼�¼���Զ��ƶ�����һ����¼��
    rs.MoveFirst
    RecordSetToString = str
End Function

'����RecordSetToString������ת��
Public Function RecordSetToStringMulti(ByRef vRs As RecordSet, _
    ByVal vDMMulti As domainMultiLevel) As String
        
    If vRs.State <> adStateOpen Then
        Exit Function
    End If
    
    If vRs.RecordCount <= 0 Then
        Exit Function
    End If
    
    Dim str As String
    Dim rs As New RecordSet
    Set rs = vRs.Clone
    
    str = ""
    rs.MoveFirst
    Do While Not rs.EOF
        '������ʽ����һ���ֶ�+�ڶ���������+�ڶ����ֶ�+��һ��������
        '����vDMMulti.Field1+vDMMulti.Symb2+vDMMulti.Field2+vDMMulti.Symb1
        str = str & rs(vDMMulti.Field1) & vDMMulti.Symb2
        str = str & rs(vDMMulti.Field2) & vDMMulti.Symb1
        rs.movenext
    Loop
       
    str = Left(str, Len(str) - 1)

    RecordSetToStringMulti = str
End Function


'����ѭ��vRs2�ļ�¼
'��vRs1�н��ֶ�vField1��VALUE��vRs2һ�����滻Ϊĳһ�ַ������������ѭ����������N������ͬ�ļ�¼��
'ʹ��vReplaceStr���滻
Public Function replaceRecord(ByRef vRs1 As RecordSet, ByRef vRs2 As RecordSet, _
    ByVal vField1 As String, ByVal vFiedl2 As String, ByVal vReplaceStr As String)
    
    Dim szTemp As String

    If vRs1.State <> adStateOpen Then
        Exit Function
    End If
    
    If vRs2.State <> adStateOpen Then
        Exit Function
    End If
    
    If vRs1.RecordCount <= 0 Then
        Exit Function
    End If
    
    If vRs2.RecordCount <= 0 Then
        Exit Function
    End If
    
    
    vRs2.MoveFirst
    Do While Not vRs2.EOF
        vRs1.Filter = " " & vField1 & "='" & vRs2(Trim$(vFiedl2)) & "'"
        If vRs1.RecordCount > 0 Then
            Do While Not vRs1.EOF
                vRs1(Trim$(vField1)) = vReplaceStr
                vRs1.movenext
            Loop
        End If
        vRs2.movenext
    Loop
    
    
    vRs1.Filter = ""
    If vRs1.RecordCount > 0 Then
        vRs1.MoveFirst
        Do While Not vRs1.EOF
            szTemp = IIf(IsNull(vRs1(Trim$(vField1))), "", vRs1(Trim$(vField1)))
            If Len(Trim$(szTemp)) <= 0 Then
                vRs1.delete
                vRs1.Update
            End If
            vRs1.movenext
        Loop
    End If
End Function

Public Sub FillUnConnectRecordSet(ByRef sRs As RecordSet, ByRef tRs As RecordSet, _
    Optional ByVal vFillData As Boolean = True)
    
    On Error Resume Next
    Dim i As Long
       
    Set tRs = New RecordSet
    For i = 0 To sRs.Fields.Count - 1
        tRs.Fields.Append sRs.Fields(i).name, sRs.Fields(i).Type, sRs.Fields(i).DefinedSize, sRs.Fields(i).Attributes
    Next
       
    tRs.Open
    
    '�������
    If vFillData = True Then
        Do While Not sRs.EOF
            tRs.AddNew
            For i = 0 To sRs.Fields.Count - 1
                tRs.Fields(Trim(sRs(i).name)).Value = IIf(IsNull(sRs.Fields(Trim(sRs.Fields(i).name)).Value), Null, sRs.Fields(Trim(sRs.Fields(i).name)).Value)
            Next
            tRs.Update
            sRs.movenext
        Loop
          
        tRs.MoveFirst
    
    End If
End Sub

'�ж�������¼����ĳ�ֶ����Ƿ���ȫһ�£�������¼����ָ���ֶ��µ�VALUE��
'��ȫһ���Ļ��򷵻�TRUE�����򷵻�FALSE
Public Function JudgeTheSame(ByVal vRs1 As RecordSet, _
    ByVal vRs2 As RecordSet, ByVal vField1 As String, _
    ByVal vField2 As String) As Boolean

    If vRs1.State <> adStateOpen Then
        JudgeTheSame = False
        Exit Function
    End If
    
    If vRs2.State <> adStateOpen Then
        JudgeTheSame = False
        Exit Function
    End If
    
    If vRs1.RecordCount <= 0 Then
        JudgeTheSame = False
        Exit Function
    End If
    
    If vRs2.RecordCount <= 0 Then
        JudgeTheSame = False
        Exit Function
    End If
    
    vRs1.Filter = ""
    vRs2.Filter = ""
    
    If vRs1.RecordCount <> vRs2.RecordCount Then
        JudgeTheSame = False
        Exit Function
    End If
    
    vRs1.Sort = " " & vField1 & " asc"
    vRs2.Sort = " " & vField2 & " asc"
    
    vRs1.MoveFirst
    vRs2.MoveFirst
    
    Do While Not vRs1.EOF
        If vRs1(Trim$(vField1)) <> vRs2(Trim$(vField2)) Then
            JudgeTheSame = False
            Exit Function
        End If
        
        vRs1.movenext
        vRs2.movenext
    Loop
    
    JudgeTheSame = True
End Function

'����һ�����������ֶΣ��������ֶγ���
Public Function getTableFieldsWithoutIncream(ByVal vTableName As String) As String
    Dim strSQL As String
    Dim rs As RecordSet
    Dim i As Long
    Dim szReturn As String
    
    Set rs = New RecordSet
    strSQL = "Select top 1 * from " & vTableName

    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.Fields.Count <= 0 Then
        getTableFieldsWithoutIncream = ""
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    
    For i = 0 To rs.Fields.Count - 1
        Debug.Print rs.Fields(i).name
        If rs.Fields(i).Properties("ISAUTOINCREMENT").Value = False Then
            szReturn = szReturn & rs.Fields(i).name & ","
        End If
    Next
    
    szReturn = Left$(szReturn, Len(szReturn) - 1)
    rs.Close
    Set rs = Nothing
    
    getTableFieldsWithoutIncream = szReturn
End Function

'��ȡһ���ֶ��������ڵļ�¼���е�����ţ��������-1��ʾ������
Public Function GetColIndex(ByRef vRs As RecordSet, ByVal vFieldName As String) As Long
    Dim rs As New RecordSet
    Set rs = vRs.Clone
    Dim i As Long
    
    
    GetColIndex = -1
    For i = 0 To rs.Fields.Count
        If rs(i).name = vFieldName Then
            GetColIndex = i
            Exit Function
        End If
    Next
End Function

Public Sub PrintFields(ByRef vRs As RecordSet)
    Dim rs As New RecordSet
    Dim i As Long
    Dim szContent As String
    
    Set rs = vRs.Clone
    For i = 0 To rs.Fields.Count - 1
        szContent = ""
        szContent = rs(i).name & "," & rs(i).Type
        Debug.Print szContent
    Next
    
    rs.Close
    Set rs = Nothing
End Sub

'��һ���ַ���ͨ���̶��ļ������ת��Ϊһ����¼������
'vString���й̶�������ŵ�һ���ַ���
'vFilterSymbol���������
'vFieldName���γɵļ�¼�����ֶε����� - ���ɵļ�¼��ֻ��һ���ֶ�
Public Function String2RS(ByVal vString As String, ByVal vFilterSymbol As String, _
    ByVal vFieldName As String) As RecordSet
    
    Dim i As Long
    Dim rs As New RecordSet
    
    vFieldName = Trim$(vFieldName)
    rs.Fields.Append vFieldName, adVarChar, 100
    rs.Open
    
    Dim arrString() As String
    arrString = Split(vString, vFilterSymbol)
    For i = 0 To UBound(arrString)
        rs.AddNew
        rs(vFieldName) = arrString(i)
    Next
    
    Set String2RS = rs.Clone
    rs.Close
    Set rs = Nothing
End Function

'�ڶ�ά������ѭ����ӡ����
Public Sub PrintData(ByRef vRs As RecordSet)
    Dim rs As New RecordSet
    Dim i As Long
    Dim szTemp As String
    
    Set rs = vRs.Clone
    If rs.RecordCount <= 0 Then
        Exit Sub
    End If
    
    '1.��ӡ�ֶ�
    szTemp = ""
    For i = 0 To rs.Fields.Count - 1
        szTemp = szTemp & rs.Fields(i).name & ","
    Next
    szTemp = Left$(szTemp, Len(szTemp) - 1)
    
    Debug.Print szTemp
    
    
    '2.��ӡ����
    rs.MoveFirst
    Do While Not rs.EOF
        szTemp = ""
        For i = 0 To rs.Fields.Count - 1
            szTemp = szTemp & rs.Fields(i).Value & ","
        Next
        szTemp = Left$(szTemp, Len(szTemp) - 1)
        Debug.Print szTemp
        rs.movenext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub


Public Function GetRS(ByVal vSQL As String, Optional ByVal vCanUpdate As Boolean = True) As RecordSet
    Dim rs As New RecordSet
    If vCanUpdate = True Then
        rs.Open vSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Else
        rs.Open vSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    End If
    
    Set GetRS = rs.Clone
    rs.Close
    Set rs = Nothing
End Function

'�ڼ�¼�����ж�һ���ֶ��Ƿ����
Public Function JudgeFieldExists(ByRef vRs As RecordSet, ByVal vFieldName As String) As Boolean
    JudgeFieldExists = False
    
    If vRs.State <> adStateOpen Then
        JudgeFieldExists = False
        Exit Function
    End If
    
    Dim i As Long
    For i = 0 To vRs.Fields.Count - 1
        If vRs.Fields(i).name = vFieldName Then
            JudgeFieldExists = True
        End If
    Next
End Function


Public Function GetFieldValue(ByRef vField As ADODB.Field) As String
    Select Case vField.Type
        Case adVarChar, adChar
            GetFieldValue = IIf(IsNull(vField.Value), "", vField.Value)       ' 1--String��
        Case adInteger, adSingle, adSmallInt, adDouble, adNumeric
            GetFieldValue = IIf(IsNull(vField.Value), 0, vField.Value)   ' 2--��ֵ��
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            GetFieldValue = IIf(IsNull(vField.Value), Format(Now, "YYYY-MM-DD"), Format(vField.Value, "YYYY-MM-DD"))   ' 3--������
        Case Else
            GetFieldValue = IIf(IsNull(vField.Value), "", vField.Value)       ' 1--String��
    End Select
End Function

'�ж�һ����¼���е�ĳ�ֶ��ϵ�ֵ�Ƿ�Ψһ
'Ψһ�򷵻�TRUE�����򷵻�FALSE
Public Function JudgeFieldValueUnique(ByRef vRs As RecordSet, _
    ByVal vField As String) As Boolean
    
    
    JudgeFieldValueUnique = True
    If vRs.State <> adStateOpen Then
        JudgeFieldValueUnique = True
        Exit Function
    End If
    
    If vRs.RecordCount <= 0 Then
        JudgeFieldValueUnique = True
        Exit Function
    End If
    
    Dim rs As New RecordSet
    Dim szTemp As String
    Dim szChecked As String
    
    Set rs = vRs.Clone

    Do While Not rs.EOF
        If Len(szTemp) <= 0 Then
            szTemp = rs(Trim$(vField))
        Else
            szChecked = rs(Trim$(vField))
            If szChecked <> szTemp Then
                JudgeFieldValueUnique = False
                Exit Function
            End If
        End If
        
        rs.movenext
    Loop
    
    rs.Close
    Set rs = Nothing
End Function

'���ع������ֶε�û�����ݵļ�¼��
Public Function buildRsNoData(ByRef vRs As RecordSet) As RecordSet
'    Dim rs As New RecordSet
'    Dim i As Long
'    For i = 0 To vRs.Fields.Count - 1
'        rs.Fields.Append vRs.Fields(0).name, vRs.Fields(0).Type, vRs.Fields(0).DefinedSize, vRs.Fields(0).Attributes
'    Next
'    rs.Open
'
'    buildRsNoData = rs.Clone
End Function


'�����ɸѡ�����ļ�¼��
Public Function buildRsWithData(ByRef vRs As RecordSet, ByVal vFilterString As String) As RecordSet
    Dim rs As New RecordSet
    Dim i As Long
    For i = 0 To vRs.Fields.Count - 1
        rs.Fields.Append vRs.Fields(i).name, vRs.Fields(i).Type, vRs.Fields(i).DefinedSize, vRs.Fields(i).Attributes
    Next
    rs.Open
    
    vRs.Filter = vFilterString
    Do While Not vRs.EOF
        rs.AddNew
        For i = 0 To vRs.Fields.Count - 1
            rs(i) = vRs(i)
        Next
        vRs.movenext
    Loop
    
   Set buildRsWithData = rs.Clone
End Function
