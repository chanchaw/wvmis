VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'�����Excel�������ã�
'Microsoft Excel 11.0 Object Library
'һ��·��Ϊ��X\Program File\Microsoft Office\OFFICE11\EXCEL.EXE

'Dim xlApp As Excel.Application '����EXCEL��
'Dim xlBook As Excel.Workbook '���幤������
'Dim xlSheet As Excel.Worksheet '���幤������

Dim xlApp As Object '����EXCEL��
Dim xlBook As Object '���幤������
Dim xlSheet As Object '���幤������

Private strSQL As String
Private A_rsRules As New RecordSet
Private A_rsMappingDetail As New RecordSet
Public W_LogFileName As String
Private A_TableName As String   '���������ݵı���

Private clsLog As New clsFile


'����Ϊ�ļ����������


'�����������ļ�
'��һ������Ϊ����·�����ļ���
'�ڶ�������Ϊ�������ļ�������
Public Sub CreateFileBat(ByVal m_FileName As String, ByVal m_Content As String)
    
    Open m_FileName For Append As #1
    Print #1, m_Content
    Close #1
    
End Sub


'���ı��ĵ�׷��д���ı���������ļ��������򴴽�����д�룩
'strFileName���ļ���������·��
'strWrite��Ҫд�������
Public Sub WriteFileContent(ByVal strFileName As String, ByVal strWrite As String)
    Dim lngHandle As Long '���
    lngHandle = FreeFile() 'ȡ�þ��
    

    Open strFileName For Append As lngHandle    '���ļ�
    Print #lngHandle, strWrite    'д���ı�
    Close lngHandle    '�ر��ļ�
End Sub

'���ı��ĵ�׷��д���ı���������ļ��������򴴽�����д�룩
'strFileName���ļ���������·��
'strWrite��Ҫд�������
'vWriteTo��1��ʾҪд�룬0��ʾ��Ҫд��
Public Sub WriteFileContent_WithSwitch(ByVal strFileName As String, _
    ByVal strWrite As String, ByVal vWriteTo As Long)
    
    If vWriteTo = 0 Then
        Exit Sub
    End If
    
    Dim lngHandle As Long '���
    lngHandle = FreeFile() 'ȡ�þ��
    

    Open strFileName For Append As lngHandle    '���ļ�
    Print #lngHandle, strWrite    'д���ı�
    Close lngHandle    '�ر��ļ�
End Sub


'��һ���ļ�
'vFileName���ļ���������·����
Public Sub OpenFile(ByVal vFileName As String)
    Dim a&
    a = ShellExecute(0, "open", vFileName, 0, 0, 1)
End Sub


'�ж�ĳ�ļ��Ƿ����
'�����ڷ���false
'���ڵĻ�����true
Public Function JudgeFileExist(ByVal m_FileName As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    JudgeFileExist = fso.FileExists(m_FileName)
       
    
    Set fso = Nothing
End Function


'ɾ���ļ�
Public Sub DeleteFile(ByVal m_FileName As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.DeleteFile (m_FileName)
End Sub


'���һ���ı��ĵ��е�����
Public Sub ClearTxtContent(ByVal vFileName As String)
    Open vFileName For Output As #1
    Close #1
End Sub


'���д��ı��ĵ��ж�ȡ����
Public Function ReadTXT(TXTPath As String) As String
    Dim TXTCon, str As String, i As Long
    
    Dim lngHandle As Long '���
    lngHandle = FreeFile() 'ȡ�þ��
    
    
    Open TXTPath For Input As lngHandle
    Do While Not EOF(lngHandle)
        'Line Input lngHandle,TXTCon
        Line Input #lngHandle, TXTCon
        'Str = Str & TXTCon & vbCrLf
        str = str & TXTCon & vbNewLine
    Loop
    Debug.Print str
    ReadTXT = str
    
    Close lngHandle
End Function

'���ж�ȡ�ı��ĵ��е����ݣ�ÿ��֮��ʹ��Ӣ�ĵĶ��ż��
Public Function ReadTXT_InsertSym(TXTPath As String) As String
    Dim TXTCon, str As String, i As Long
    
    Dim lngHandle As Long '���
    lngHandle = FreeFile() 'ȡ�þ��
    
    
    Open TXTPath For Input As lngHandle
    Do While Not EOF(lngHandle)
        Line Input #lngHandle, TXTCon
        str = str & TXTCon & ","
    Loop
    ReadTXT_InsertSym = Left(str, Len(str) - 1)
    
    Close lngHandle
End Function


'���ı��ĵ��ж�ȡ���һ�е�����
'����Ϊ�ļ���ȫ·���ļ���
Private Function ReadTXTLast(ByVal vFileName As String) As String
    Dim TXTCon, str As String, i As Long

    Open vFileName For Input As #1
    Do While Not EOF(1)
        Line Input #1, TXTCon
    Loop
    ReadTXTLast = TXTCon
End Function

'ͨ���������ļ��Ի�������ȡһ���ļ��ľ���·��
'�����ļ��ľ���·��
Public Function OpenFile_Dialog() As String
    Dim szFileName As String
    Dim o As Object
    Set o = CreateObject("MSComDlg.CommonDialog")
    With o
        .Filter = "Excel�ļ� (*.xls)|*.xls|�����ļ�(*.*)|*.*|Excel�ļ� (*.xlsx)|*.xlsx"
        .ShowOpen
        szFileName = .FileName
    End With
    
    OpenFile_Dialog = Trim(szFileName)
    
    Set o = Nothing
End Function


Public Function OpenFile_DialogAllEt() As String
    Dim szFileName As String
    Dim o As Object
    Set o = CreateObject("MSComDlg.CommonDialog")
    With o
        .Filter = "�����ļ�(*.*)|*.*"
        .ShowOpen
        szFileName = .FileName
    End With
    
    OpenFile_DialogAllEt = Trim(szFileName)
    
    Set o = Nothing
End Function


'��ȡһ��EXCEL�ļ��е����ݵ���¼����
'���ݽ����Ĳ������ļ��ľ���·����ȫ�ļ�����
'vRows:��EXCEL�ļ��ж�ȡ������
'vCols:��EXCEL�ļ��ж�ȡ������
Public Function ReadExcelToRecordSet(ByVal vFileName As String, _
    ByVal vRows As Long, ByVal vCols As Long, ByVal vExcludeRows As Long, _
    ByVal vExcludeCols As Long) As RecordSet
    
    On Error GoTo IFERR
    
    Dim i As Long, J As Long
    Dim retRecordSet As RecordSet
    Dim szErrExcelFileName As String
    Dim szTemp As String
    Dim szUnit As String
    Dim clsString1 As New clsString
    
    If Len(Trim$(vFileName)) <= 0 Then
        Exit Function
    End If
    
    If vRows <= 0 Then
        Exit Function
    End If
    
    If vCols <= 0 Then
        Exit Function
    End If
    
    '��ȡ�����ļ��ľ���·��+�ļ�����
    '��2015��10��8�� 13:25:01�������������ݵ��뵽һ��EXCEL�ļ��м�¼����
    '��������������on error resume next������ʵ�֡��콫�����һ�������ע�͵���
'    szErrExcelFileName = Format(Now, "YYYYMMDD HHMMSS") & "����ʧ�ܵ�����.xls"
'    If Right$(g_ErrFilePath, 1) <> "\" Then
'        g_ErrFilePath = g_ErrFilePath & "\"
'    End If
'    szErrExcelFileName = g_ErrFilePath & szErrExcelFileName
    
    
    writeLogImportLog "��ʼ����ExcelӦ�ó���"
    Set xlApp = CreateObject("Excel.Application") '����EXCELӦ����
    xlApp.Visible = False '����EXCEL���ɼ�
    
    Set xlBook = xlApp.Workbooks.Open(vFileName) '��EXCEL������
    writeLogImportLog "�����������"
    Set xlSheet = xlBook.Worksheets(1) '��EXCEL������
    writeLogImportLog "sheet�����"
    
    Set retRecordSet = New RecordSet
    For i = 1 To vCols
        retRecordSet.Fields.Append "B_Field" & i, adVarChar, 100
    Next
    retRecordSet.Open
    writeLogImportLog "�ռ�¼����ʼ�����"
    
    writeLogImportLog "����������������£�////////////////"
    For i = (1 + vExcludeRows) To (vRows + vExcludeRows)
        retRecordSet.AddNew
        szTemp = ""
        For J = (1 + vExcludeCols) To (vCols + vExcludeCols)
            szUnit = clsString1.KillCrlfSpace(xlSheet.Cells(i, J))
            retRecordSet(J - 1) = szUnit
            szTemp = szTemp & szUnit & ","
        Next
        szTemp = Left$(szTemp, Len(szTemp) - 1)
        writeLogImportLog szTemp
    Next
    writeLogImportLog "���ݵ�����ϣ��ȴ��ر�/////////////"
    
    
    xlBook.Close (True) '�ر�EXCEL������
    xlApp.Quit '�ر�EXCEL
    Set xlApp = Nothing '�ͷ�EXCEL����
    
    writeLogImportLog "��ʽ�ر���ϣ��ȴ�����"
    Set ReadExcelToRecordSet = retRecordSet.Clone
    
    Exit Function
IFERR:
    Dim szErr As String
    szErr = "��ReadExcelToRecordSet�г�����" & vbNewLine & Err.Description
    writeLogImportLog szErr
    MsgBox szErr, vbOKOnly + vbInformation, "��ʾ"
End Function

'���ݱ�G_ExcelMappingTableMain��G_ExcelMappingTableDetail
'������ļ�¼���е�����д�뵽��Ӧ�ı���
'oRs:����ļ�¼�������д�������
'vMappingIndex����Ӧ�����ݱ�G_ExcelMappingTableMain��B_ID�������������ŵ��趨�����뵽ĳ������
'���Ұ���G_ExcelMappingTableDetail���趨���ֶζ�Ӧ����
'Ҫʹ�ñ��������ա���徫����װ��ӡϵͳ���е�3�������ױ���
'G_ExcelMappingTableMain��G_ExcelMappingTableDetail��G_ExcelMappingTableRules
Private Function WriteRSToTable_ByMappingTable(ByRef oRs As RecordSet, ByVal vMappingIndex As Long) As Boolean
    On Error Resume Next
    Dim rsMap As RecordSet
    Dim rsData As RecordSet
    Dim rsTable As RecordSet
    Dim i As Long
    
    WriteRSToTable_ByMappingTable = True
    

    
    '��ȡ��G_ExcelMappingTableMain��G_ExcelMappingTableDetailָ��B_ID������
    strSQL = "exec dbo.P_GetMappingTable " & vMappingIndex
    Debug.Print strSQL
    Set rsMap = New RecordSet
    rsMap.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rsMap.RecordCount <= 0 Then
        WriteRSToTable_ByMappingTable = False
        rsMap.Close
        Set rsMap = Nothing
        Exit Function
    End If
    
    '��ȡ�������������ݵı�����
    A_TableName = IIf(IsNull(rsMap!B_TableName), "", rsMap!B_TableName)
    
    If oRs.State <> adStateOpen Then
        WriteRSToTable_ByMappingTable = False
        rsMap.Close
        Set rsMap = Nothing
        Exit Function
    End If
    
    If oRs.RecordCount <= 0 Then
        WriteRSToTable_ByMappingTable = False
        rsMap.Close
        Set rsMap = Nothing
        Exit Function
    End If
    
    Set rsData = New RecordSet
    Set rsData = oRs.Clone
    
    
    Set rsTable = New RecordSet
    strSQL = "Select * From " & rsMap!B_TableName
    Debug.Print strSQL; strSQL
    rsTable.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    'rsData:��EXCEL��ȡ������
    'rsTable�����ݵ����Ŀ�ı��л�ȡ�ļ�¼��
    rsData.MoveFirst
    Do While Not rsData.EOF
        '�����ж�ĳ�ֶ��ϵ�ֵΨһ�������
        '�����Ҫ�жϣ�ͬʱ�жϵó�û���ظ���¼�����ADDNEW
        '�����Ҫ�жϣ�ͬʱ�жϵó����ظ��ļ�¼����������һ��EXCEL�еļ�¼����ִ�С�
        If JudgeRules_Repeat(vMappingIndex, rsData, rsTable) = True Then
    
            rsTable.AddNew
            For i = 1 To rsData.Fields.Count
                rsMap.Filter = " B_ExcelColIndex='" & i & "'"
                If rsMap.RecordCount > 0 Then
                    rsTable(Trim(rsMap!B_FieldName)) = rsData(i - 1)
                End If
            Next
            
            '�����ǰ�в���������������Ҫȡ����
            If JudgeRules(vMappingIndex, rsTable) = False Then
                rsTable.CancelUpdate
            Else
                rsTable.Update
            End If
        
        End If
        
        rsData.movenext
    Loop
    
    rsTable.Close
    Set rsTable = Nothing
    
    
    rsData.Close
    Set rsData = Nothing
    
    rsMap.Close
    Set rsMap = Nothing
End Function

'ͨ������һ�����壬������������ԴExcel�еĵ������ݵ�������������
'�ų���ǰN�У��ų���ǰN�С�������һ���������ݵ�����
'���صĽṹ����modStruecture������
'����ʹ�õ��Ľṹ��λ��soft/������/modStructure
Private Function getImportExcelRect() As importExcelRect
    Dim frm1 As New frmSetGridRowsCols
    Dim rtn As importExcelRect
    frm1.Show vbModal
    
    rtn.Rows = frm1.w_Rows
    rtn.cols = frm1.w_Cols
    rtn.excludeRows = frm1.w_ExcludeRows
    rtn.excludeCols = frm1.w_ExcludeCols
    
    getImportExcelRect = rtn
    
    Unload frm1
End Function

Public Function WriteRSToTable() As Boolean
    Dim szFileName As String
    Dim rsExcel As RecordSet
    
    Dim i As Long, J As Long, m As Long, n As Long
    
    Dim szTemp As String
    
    WriteRSToTable = True
    
    szFileName = OpenFile_Dialog
    If Len(Trim(szFileName)) <= 0 Then
        WriteRSToTable = False
        MsgBox "�ļ����ƷǷ���", vbOKOnly + vbInformation, "��ʾ"
        Exit Function
    End If
    
    

    '��ȡExcel�ļ�������Դ����
    writeLogImportLog "��" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "��ʼ�ĵ������======================="
    writeLogImportLog "��ʼ������Դ����rect"
    
    Dim RECT As importExcelRect
    RECT = getImportExcelRect
    i = RECT.Rows
    J = RECT.cols
    m = RECT.excludeRows
    n = RECT.excludeCols
    
    writeLogImportLog "��ȡ������Դ����Ϊ������=" & i & ",����=" & J & ",�ų�ǰ" & m & "��,�ų�ǰ" & n & "��"
    
    
    
    If i <= 0 Or J <= 0 Then
        WriteRSToTable = False
        MsgBox "�������������Ƿ���", vbOKOnly + vbInformation, "��ʾ"
        Exit Function
    End If
    
    writeLogImportLog "��ʼ����ReadExcelToRecordSet"
    Set rsExcel = ReadExcelToRecordSet(szFileName, i, J, m, n)
    writeLogImportLog "��ReadExcelToRecordSetִ����ϣ�"
    
    WriteRSToTable_ByMappingTable rsExcel, 1
    writeLogImportLog "WriteRSToTable_ByMappingTableִ�����"
End Function


'��ʢ���뿼�������õ�.����֮ǰ��ɾ��֮ǰ����������
Public Function WriteRSToTable_ByGroupName(ByVal vGroupName As String) As Boolean
    Dim szFileName As String
    Dim rsExcel As RecordSet
    
    Dim i As Long, J As Long, m As Long, n As Long
    
    WriteRSToTable_ByGroupName = True
    
    szFileName = OpenFile_Dialog
    If Len(Trim(szFileName)) <= 0 Then
        WriteRSToTable_ByGroupName = False
        MsgBox "�ļ����ƷǷ���", vbOKOnly + vbInformation, "��ʾ"
        Exit Function
    End If
    
    
    Dim frm1 As New frmSetGridRowsCols
    'frm1.W_ExcelImportDefaultGroupName = vGroupName   '��徫��û�����ñ�G_ExcelImportDefault.��ʢ����ϵͳ��
    frm1.Show vbModal
    
    i = frm1.w_Rows
    J = frm1.w_Cols
    m = frm1.w_ExcludeRows
    n = frm1.w_ExcludeCols
    
    Unload frm1
    
    If i <= 0 Or J <= 0 Then
        WriteRSToTable_ByGroupName = False
        MsgBox "�������������Ƿ���", vbOKOnly + vbInformation, "��ʾ"
        Exit Function
    End If
    
    '��ɾ��֮ǰ����������
    DelBeforeImport vGroupName
    
    Set rsExcel = ReadExcelToRecordSet(szFileName, i, J, m, n)
    
    WriteRSToTable_ByMappingTable rsExcel, 1
End Function

'�ڵ�������ǰ����ϸ���е�����ȫ��ɾ��
Private Sub DelBeforeImport(ByVal vGroupName As String)
    Dim rs As RecordSet
    strSQL = "Select * From G_ExcelImportDefault Where B_GroupName='" & vGroupName & "'"
    Set rs = New RecordSet
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    If IIf(IsNull(rs!B_DelBeforeImport), 0, rs!B_DelBeforeImport) = 1 Then
        strSQL = "Truncate table G_KQDetail"
        Gm.cnnTool.cnn.Execute strSQL
    End If
    rs.Close
    Set rs = Nothing
End Sub


'����TRUE��ʾ����ǰһ�����ݿ��Խ��б���
'����FALSE��ʾ��ǰһ��������Ҫͨ��CANCLEȡ����
Private Function JudgeRules(ByVal vMappingIndex As Long, ByRef rsTable As RecordSet) As Boolean
    On Error GoTo IFERR
    Dim rsRules As RecordSet
    Dim szGZFieldName As String
    Dim szGZFieldValue As String
    Dim szFilterString As String
    
    JudgeRules = True
    
    Set rsRules = New RecordSet
    strSQL = "Select * From G_ExcelMappingTableRules Where B_ID=" & vMappingIndex
    rsRules.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rsRules.RecordCount <= 0 Then
        rsRules.Close
        Set rsRules = Nothing
        Exit Function
    End If
    
    Do While Not rsRules.EOF
        '��ֵ����0
        If IIf(IsNull(rsRules!B_SZDY0), 0, rsRules!B_SZDY0) = 1 Then
            If IIf(IsNull(rsTable(Trim$(rsRules!B_FieldName))), 0, rsTable(Trim$(rsRules!B_FieldName))) <= 0 Then
                JudgeRules = False
                rsRules.Close
                Set rsRules = Nothing
                Exit Function
            End If
        End If
        
        '���ȴ���0
        If IIf(IsNull(rsRules!B_CDDY0), 0, rsRules!B_CDDY0) = 1 Then
            If Len(Trim(IIf(IsNull(rsTable(Trim$(rsRules!B_FieldName))), "", rsTable(Trim$(rsRules!B_FieldName))))) <= 0 Then
                JudgeRules = False
                rsRules.Close
                Set rsRules = Nothing
                Exit Function
            End If
        End If
        
        
        'ĳ�ֶ��ϵ�ֵΨһ
'        If IIf(IsNull(rsRules!B_OnlyOne), 0, rsRules!B_OnlyOne) = 1 Then
'            szGZFieldName = Trim$(rsRules!B_FieldName)
'            szGZFieldValue = rsTable(Trim$(rsRules!B_FieldName))
'            szFilterString = " " & szGZFieldName & "='" & szGZFieldValue & "'"
'            Debug.Print szFilterString
'            rsTable.Filter = szFilterString
'            If rsTable.RecordCount > 0 Then
'                JudgeRules = False
'                rsRules.Close
'                Set rsRules = Nothing
'                Exit Function
'            End If
'        End If
        
        
        rsRules.movenext
    Loop
    
    rsRules.Close
    Set rsRules = Nothing
    
    Exit Function
IFERR:
    Dim szErr As String
    szErr = Err.Description & vbNewLine
    szErr = szErr & "������JudgeRules��"
    MsgBox Err.Description, vbOKOnly + vbInformation, "��ʾ"
End Function

'��EXCEL�������Ʒ�������
Public Function ImportExcelData_Goods() As Boolean
    Dim szFileName As String
    Dim rsExcel As RecordSet
    
    Dim i As Long, J As Long, m As Long, n As Long
    
    ImportExcelData_Goods = True
    
    szFileName = OpenFile_Dialog
    If Len(Trim(szFileName)) <= 0 Then
        ImportExcelData_Goods = False
        MsgBox "�ļ����ƷǷ���", vbOKOnly + vbInformation, "��ʾ"
        Exit Function
    End If
    
    
    Dim frm1 As New frmSetGridRowsCols
    frm1.Show vbModal
    
    i = frm1.w_Rows
    J = frm1.w_Cols
    m = frm1.w_ExcludeRows
    n = frm1.w_ExcludeCols
    
    Unload frm1
    
    If i <= 0 Or J <= 0 Then
        ImportExcelData_Goods = False
        MsgBox "�������������Ƿ���", vbOKOnly + vbInformation, "��ʾ"
        Exit Function
    End If
    
    
    Set rsExcel = ReadExcelToRecordSet(szFileName, i, J, m, n)
    
    WriteRSToTable_ByMappingTable rsExcel, 2
End Function

'�жϹ����G_ExcelMappingTableRules����B_OnlyOne
'oRsMap:���ݽ���rsMap����Excel�����ݱ�֮���ӳ���
Private Function JudgeOnlyOne(ByRef oRsMap As RecordSet, ByRef oRsData As RecordSet) As Boolean
    Dim rsMap As RecordSet
    Set rsMap = oRsMap.Clone
    
    Dim rsData As RecordSet
    Set rsData = oRsData.Clone
    
    If rsMap.State <> adStateOpen Then
        JudgeOnlyOne = True
        Exit Function
    End If
    
    If rsMap.RecordCount <= 0 Then
        JudgeOnlyOne = True
        Exit Function
    End If
    
    rsMap.MoveFirst
    Do While Not rsMap.EOF
        If IIf(IsNull(rsMap!B_OnlyOne), 0, rsMap!B_OnlyOne) = 1 Then
            
        End If
        rsMap.movenext
    Loop
End Function

'�жϵ�ǰ�ֶ��Ƿ���Ҫ�ж��ظ�ֵ������
'�ظ��Ļ�����FALSE
'���ظ��Ļ�����TRUE
Private Function JudgeRules_Repeat(ByVal vMappingIndex As Long, ByRef vRsDataExcel As RecordSet, ByRef vRsDataTable As RecordSet) As Boolean
    On Error GoTo IFERR
    Dim szFilter As String
    JudgeRules_Repeat = True
    If A_rsRules.State <> adStateOpen Then
        JudgeRules_Repeat = True
        Exit Function
    End If
    
    If A_rsRules.RecordCount <= 0 Then
        JudgeRules_Repeat = True
        Exit Function
    End If
    
    Dim rsDataTable As RecordSet
    Dim szRuleFieldName As String
    Dim szRuleFieldValue As String
    Dim szFilterString As String
    
    
    szFilter = ""
    szFilter = " B_ID=" & vMappingIndex
    A_rsRules.Filter = szFilter
    If A_rsRules.RecordCount <= 0 Then
        A_rsRules.Filter = ""
        JudgeRules_Repeat = True
        Exit Function
    End If
    
    szFilter = szFilter & " And B_OnlyOne=1"
    'A_rsRules.Filter = " B_OnlyOne=1"
    Debug.Print szFilter
    A_rsRules.Filter = szFilter
    If A_rsRules.RecordCount <= 0 Then
        A_rsRules.Filter = ""
        JudgeRules_Repeat = True
        Exit Function
    End If
    
    
    Dim szFieldCN As String
    Dim szFieldValue As String
    Dim szTip As String
    
    
    Set rsDataTable = vRsDataTable.Clone
    A_rsMappingDetail.Filter = " B_ID=" & vMappingIndex
    Do While Not A_rsRules.EOF
        A_rsMappingDetail.Filter = " B_FieldName='" & Trim$(A_rsRules!B_FieldName) & "'"
        If A_rsMappingDetail.RecordCount > 0 Then
            szRuleFieldName = Trim$(A_rsRules!B_FieldName)
            szRuleFieldValue = Trim(vRsDataExcel(Val(A_rsMappingDetail!B_ExcelColIndex) - 1))
            szFilterString = " " & szRuleFieldName & "='" & szRuleFieldValue & "'"
            
            rsDataTable.Filter = szFilterString
            If rsDataTable.RecordCount > 0 Then
                JudgeRules_Repeat = False
                szFieldCN = Get_FieldCN_110901(Trim$(A_rsRules!B_FieldName))
                szTip = szFieldCN & ":" & szRuleFieldValue & "�ظ�" & vbNewLine
                szTip = szTip & "���¼�±������ţ����ȷ��Ҫ����������һ�����ݣ�"
                MsgBox szTip, vbOKOnly + vbInformation, "��ʾ"
                Exit Function
            End If
            
        End If
        A_rsRules.movenext
    Loop
    
    
    Exit Function
IFERR:
    Dim szErr As String
    szErr = Err.Description & vbNewLine
    szErr = szErr & "������JudgeRules_Repeat��"
    MsgBox szErr, vbOKOnly + vbInformation, "��ʾ"
End Function


Private Sub InitRules()
    Set A_rsRules = New RecordSet
    strSQL = "Select * From G_ExcelMappingTableRules"
    A_rsRules.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
End Sub

Private Sub InitMappingDetail()
    InitMappingDetail
    InitRules
    
    
    Set A_rsMappingDetail = New RecordSet
    strSQL = "Select * From G_ExcelMappingTableDetail"
    A_rsMappingDetail.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
End Sub


'��ȡ����110901���ֶε���������
Private Function Get_FieldCN_110901(ByVal vFieldName As String) As String
    Dim rs As New RecordSet
    Dim szReturn As String
    
    strSQL = "Select * From dbo.G_BLSField Where B_ObjectID='110901'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        szReturn = ""
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    
    rs.Filter = " B_FieldName='" & vFieldName & "'"
    If rs.RecordCount > 0 Then
        szReturn = rs!B_CnName
    Else
        szReturn = ""
    End If
    rs.Close
    Set rs = Nothing
    
    Get_FieldCN_110901 = szReturn
End Function

'��·���л�ȡ�ļ���
Public Function Get_FileName_FromPath(ByVal vStrPath As String) As String
    Dim i As Long
    i = InStrRev(vStrPath, "\")
    Get_FileName_FromPath = Mid$(vStrPath, i + 1)
End Function

'����һ���ļ�
Public Function CopyFile(ByVal voPath As String, vtPath As String) As Boolean
    Dim szFileName As String
    Dim szTPathFileName As String
    
    '���Ŀ��·���Ƿ���ڣ� �������򴴽�
    Judge_Folder_Exists_And_AddNew vtPath
    
    szFileName = Get_FileName_FromPath(voPath)
    If Right$(vtPath, 1) = "\" Then
        szTPathFileName = vtPath & szFileName
    Else
        szTPathFileName = vtPath & "\" & szFileName
    End If
    Debug.Print szTPathFileName
    
    FileCopy voPath, szTPathFileName
End Function

'�ж�һ��·�����ļ��У��Ƿ����
Public Function Judge_Folder_Exists(ByVal vPath As String) As Boolean
    If Dir(vPath, vbDirectory) = "" Then
        Judge_Folder_Exists = False
    Else
        Judge_Folder_Exists = True
    End If
End Function

'һ���ļ�·�����ӷ�����Ŀ¼��ʼ��⣬����ļ��м�⣬·�������ڵĻ�
'����㴴��
Private Sub Create_Folder(ByVal vPath As String)
    Dim arr01
    Dim szFilterSym As String
    Dim i As Long, J As Long
    Dim szFilePath As String
    
    
    szFilterSym = "\"
    arr01 = Split(vPath, szFilterSym)
    
    
    If UBound(arr01) >= 1 Then
        szFilePath = arr01(0)
        For i = 1 To UBound(arr01)
            szFilePath = szFilePath & "\" & arr01(i)
            If Judge_Folder_Exists(szFilePath) = False Then
                MkDir szFilePath
            End If
        Next
    End If

End Sub

'�ж�һ��·���Ƿ���ڣ��������򴴽�
Public Sub Judge_Folder_Exists_And_AddNew(ByVal vPath As String)
    '�����һ��\ȥ����΢����ͬ�ĸ�ʽ��
    If Right$(vPath, 1) = "\" Then
        vPath = Left$(vPath, Len(vPath) - 1)
    End If
    
    
    If Judge_Folder_Exists(vPath) = False Then
        Create_Folder vPath
    End If
End Sub

'���������ļ��ĶԻ����������ļ�
Public Function SaveFile_PopUpDialog() As String
    Dim szFileName As String
    Dim o As Object
    Set o = CreateObject("MSComDlg.CommonDialog")
    With o
        .Filter = "�ı��ļ� (*.txt)|*.txt|�����ļ�(*.*)|*.*"
        .ShowSave
        szFileName = .FileName
    End With
    
    SaveFile_PopUpDialog = Trim(szFileName)
    
    Set o = Nothing
End Function

'���õ�ʱ���Զ���������ļ�����
Public Function OpenFile_Dialog_Filter(ByVal vFilterString As String) As String
    Dim szFileName As String
    Dim o As Object
    Set o = CreateObject("MSComDlg.CommonDialog")
    With o
        '.Filter = "Excel�ļ� (*.xls)|*.xls|�����ļ�(*.*)|*.*|Excel�ļ� (*.xlsx)|*.xlsx"
        .Filter = vFilterString
        .ShowOpen
        szFileName = .FileName
    End With
    
    OpenFile_Dialog_Filter = Trim(szFileName)
    
    Set o = Nothing
End Function


Public Sub WriteLog(ByVal vContent As String)
    If Len(Trim$(W_LogFileName)) <= 0 Then
        MsgBox "���ȳ�ʼ����־�ļ����ƣ�", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    WriteFileContent W_LogFileName, Now & ":"
    WriteFileContent W_LogFileName, vContent & vbNewLine
End Sub


'��Ҫ�ڹ���->����  �����á�Microsoft Scripting Runtime��
'vFolderPath���ļ���·������ѯ���ļ����µ��������ļ���
'vBSubFolders���Ƿ���������ö�����ļ���
Private Sub TraversalFolders(ByVal vFolderPath As String, ByVal vBSubFolders As Boolean)
    Dim fs, F, f1, S, sf
    Dim hs, H, H1, hf
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set F = fs.GetFolder(vFolderPath)
    
    Set sf = F.SubFolders
    
    For Each f1 In sf
        'List1.AddItem vFolderPath & "\" & f1.name
        If vBSubFolders = True Then
            Call TraversalFolders(vFolderPath & "\" & f1.Name, vBSubFolders)
        End If
    Next
End Sub



'����һ���ļ����µ������ļ����ҷ���һ����¼��
'�ü�¼��ֻ��һ���ֶ�B_Field0�������ļ��ľ���·��
Public Function TraversalFilesReturnRS(ByVal vFolderPath As String) As RecordSet
    Dim fs, F, f1, fc, S
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set F = fs.GetFolder(vFolderPath)
    Set fc = F.Files
    
    
    Dim rs As New RecordSet
    rs.Fields.Append "B_Field0", adVarChar, 1000
    rs.Open
    
    For Each f1 In fc
        rs.AddNew
        rs(0) = vFolderPath & "\" & f1.Name
    Next
    
    Set TraversalFilesReturnRS = rs
End Function

'vPathAndFileName:����·��+�ļ����ƣ����к�׺����
'����һ���ļ����������·���������򴴽�
'������ļ��Ѿ����ڣ��򲻻��ٴ������ļ����ҷ���FALSE
'����ļ������ڣ��򴴽����ҷ���TRUE
Public Function CreateFile(ByVal vPathAndFileName As String) As Boolean
    Dim lngHandle As Long '���
    Dim cslString1 As New clsString
    
    CreateFile = False
    
    '��ʽ������
    vPathAndFileName = cslString1.KillCrlfSpace(vPathAndFileName)
    
    '�ж�·���Ƿ�Ϊ��
    If Len(Trim$(vPathAndFileName)) <= 0 Then
        CreateFile = False
        Exit Function
    End If
    
    lngHandle = FreeFile() 'ȡ�þ��
    
    Dim szPath As String
    Dim i As Long
    i = InStrRev(vPathAndFileName, "\") '��ȡ���һ��б�ߵ�INDEX
    If Mid$(vPathAndFileName, i - 1, 1) = ":" Then
        '��������̷��µĸ�Ŀ¼
        szPath = Left(vPathAndFileName, i)
    Else
        '�Ƿ����ĸ�Ŀ¼
        szPath = Left$(vPathAndFileName, i - 1)
    End If
    
    '���·���Ƿ���ڣ��������������㴴��
    Judge_Folder_Exists_And_AddNew szPath
    
    If JudgeFileExist(vPathAndFileName) = False Then
        '����һ���ļ�
        Dim fso As New FileSystemObject
        Dim fileStream As TextStream
        Set fileStream = fso.CreateTextFile(vPathAndFileName)
        fileStream.Close
        
        Exit Function
    End If

    CreateFile = False
End Function

'��2015��10��8�� 13:26:35����
'д��֮ǰ�жϵ�ǰ�ļ��Ƿ���ڣ�����������򴴽�
'������ʱ��Ҫд����ͷ���Ǵ�����ʱ��ֻҪд�����ݼ��ɡ�
'xlSheet���Ӹò����п��Ի�ȡ��һ�У�����ͷ��
'vExcludeCols���ų�ǰ������
'vCols��Ҫ������е���Ŀ
'vErrFileName��ÿ�ε��õ��뺯����ʱ����ļ����ƶ���һ�����жϵ�ǰ�ļ����Ƿ����
'vRowIndex�������ʱ��������кš��Ὣ���е���������д�������־��
Private Sub WriteErrExcel(ByRef xlSheet As excel.Worksheet, _
    ByVal vExcludeCols As Long, ByVal vCols As Long, _
    ByVal vErrFileName As String, ByVal vRowIndex As Long)
    
    
    
    Dim clsExcel1 As New clsExcel
    Dim i As Long
    Dim szTemp As String
    
    '�����������ݼ�¼�ļ�
    If JudgeFileExist(vErrFileName) = False Then
        CreateFile vErrFileName
        
        '������ͬʱҪд����ͷ
        For i = (1 + vExcludeCols) To (vCols + vExcludeCols)
            szTemp = xlSheet.Cells(1, i)
            clsExcel1.WriteContentAtCell 1, i, szTemp
        Next
    End If
    
    
    'д��û�е���ɹ�������
    For i = (1 + vExcludeCols) To (vCols + vExcludeCols)
         
    Next
End Sub

Private Sub writeLogImportLog(ByVal vContent As String)
    Dim g_szFileName_ImportLog As String
    Dim g_lImportLog As Long
    clsLog.WriteFileContent_WithSwitch g_szFileName_ImportLog, vContent, g_lImportLog
End Sub

'��Ҫ��ͼƬ��ʾ��PictureBox�ؼ��У��ſ��Գɹ��ϴ���DB
'��ȡ��ʱ�����ʹ��PerpertyBag_DownloadFromDB������ȷ���ز�����ʾ��PictureBox��
Public Function PerpertyBag_Upload2DB(ByRef oPicture As Object, ByVal vTable As String, _
    ByVal vKeyField As String, ByVal vKeyFieldValue As String, ByVal vPictureField As String) As Boolean

    On Error GoTo IFERR
    Dim strSQL As String
    Dim rs As New RecordSet
    strSQL = "Select * From " & vTable & " where " & vKeyField & " ='" & vKeyFieldValue & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    Dim Bag As PropertyBag
    Dim buff() As Byte
    Set Bag = New PropertyBag
    Bag.WriteProperty "Image", oPicture.Picture
    ReDim buff(LenB(Bag.Contents))
    buff = Bag.Contents
    
    rs(Trim$(vPictureField)) = buff
    rs.Update
    Set rs = Nothing
    Set Bag = Nothing
    
    PerpertyBag_Upload2DB = True
    Exit Function
IFERR:
    Dim szErr As String
    PerpertyBag_Upload2DB = False
    szErr = "���������ϴ��С�" & vbNewLine & Err.Description
    'MsgBox szErr, vbOKOnly + vbInformation, "��ʾ"
End Function

'��������ʹ��PerpertyBag_Upload2DB���ſ�����ȷ���ز�����ʾͼƬ
Public Function PerpertyBag_DownloadFromDB(ByRef oPicture As Object, ByVal vTable As String, _
    ByVal vKeyField As String, ByVal vKeyFieldValue As String, ByVal vPictureField As String) As Boolean
    
    On Error Resume Next
    Dim strSQL As String
    Dim rs As New RecordSet
    strSQL = "Select * From " & vTable & " where " & vKeyField & " ='" & vKeyFieldValue & "'"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    Dim buff() As Byte
    Dim Bag As New PropertyBag
    

    buff = rs.Fields(Trim$(vPictureField)).Value
    
    
    Set Bag = New PropertyBag
    Bag.Contents = buff
    Call Bag.WriteProperty("Image", buff)
    
    '�ڼ���ͼƬǰ����տؼ��ڵ�ͼƬ
    oPicture.Picture = LoadPicture("")
    Set oPicture.Picture = Bag.ReadProperty("Image")
    
    
    Set rs = Nothing
    Set Bag = Nothing
    
    PerpertyBag_DownloadFromDB = True
    Exit Function
IFERR:
    Dim szErr As String
    PerpertyBag_DownloadFromDB = False
    szErr = "�������������С�" & vbNewLine & Err.Description
    'MsgBox szErr, vbOKOnly + vbInformation, "��ʾ"
End Function

Public Sub DownloadPic(ByRef vFld As ADODB.Field, ByVal vPath As String)
    Dim Stream As ADODB.Stream
    Set Stream = New ADODB.Stream
    
    If Not IsNull(vFld) Then
        Stream.Type = adTypeBinary
        'Stream.Mode = adModeReadWrite
        Stream.Open
        Stream.Write vFld
        Stream.SaveToFile vPath, adSaveCreateOverWrite
        Stream.Close
    End If
End Sub

'�������ܣ�
'vRs��������ͼƬ��Դ������Դ
'vPicField������ͼƬ�ļ����ֶ���
'oCtl��������ʾ�Ŀؼ���PictureBox��Image
Public Sub PicShowToCtl(ByRef vFld As ADODB.Field, ByRef oCtl As Object)
    On Error GoTo IFERR
    Dim Stream As ADODB.Stream
    Set Stream = New ADODB.Stream
    

    oCtl.Picture = LoadPicture("")
    If Not IsNull(vFld) Then
        Stream.Type = adTypeBinary
        Stream.Open
        Stream.Write vFld
        Stream.SaveToFile "filename", adSaveCreateOverWrite
        oCtl.Picture = LoadPicture("filename")
        Stream.Close
    End If
    
    Set Stream = Nothing
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "������������ͼƬ�У�" & Err.Description
    
End Sub


'fld����¼���е��ֶ�
'vFilePath��ͼƬ�ļ��ľ���·��������ͼƬ�ļ�������չ��
Public Sub PicSaveToDB(ByRef fld As ADODB.Field, ByVal vFilePath As String)
    Const blocksize = 4096
    Dim bytedata() As Byte
    Dim numblocks As Long
    Dim filelength As Long
    Dim leftover As Long
    Dim sourcefile As Long
    Dim i As Long
    
    '���û���ļ�����ôֱ���˳�
    If Len(vFilePath) <= 0 Then
        Exit Sub
    End If
    
    sourcefile = FreeFile
    
    Open Trim(vFilePath) For Binary Access Read As sourcefile
    filelength = LOF(sourcefile)
    
    If filelength = 0 Then
        Close sourcefile
        'MsgBox Trim(vFilePath) & "�����ݻ򲻴��ڣ�"
    Else
        numblocks = filelength \ blocksize
        leftover = filelength Mod blocksize
        fld.Value = Null
        
        ReDim bytedata(blocksize)
        
        For i = 1 To numblocks
            Get sourcefile, , bytedata()
            fld.AppendChunk bytedata()
        Next
        
        ReDim bytedata(leftover)
        Get sourcefile, , bytedata()
        fld.AppendChunk bytedata()
        Close sourcefile
    End If
End Sub


Public Sub WriteFileContentChoice(ByVal strFileName As String, ByVal strWrite As String, ByVal vChoice As Long)
    Dim lngHandle As Long '���
    
    If vChoice = 0 Then
        Exit Sub
    End If
    
    lngHandle = FreeFile() 'ȡ�þ��
    

    Open strFileName For Append As lngHandle    '���ļ�
    Print #lngHandle, strWrite    'д���ı�
    Close lngHandle    '�ر��ļ�
End Sub


Public Function ShowSaveFileDialog(Optional ByVal vFilterString As String = "�����ļ�(*.*)|*.*") As String
    Dim szFileName As String
    Dim o As Object
    Set o = CreateObject("MSComDlg.CommonDialog")
    With o
        .Filter = vFilterString
        .ShowSave
        szFileName = .FileName
    End With
    
    ShowSaveFileDialog = Trim(szFileName)
    
    Set o = Nothing
End Function