VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Link_Data_Ctl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'����������ʹ�÷�����
'��һ����¼���󶨵�ComboBox�У�����ʾ���ݺ�Key���ݷ��롣
'������ʾ��Caption��������ȡ��Ӧ��Key�ֶε�VALUE����ȡ��

Private oCombo As Object 'ע���ComboBox�ؼ������ǿؼ����顣����ʹ��oCombo.clear�Ƿ�
Private A_rs As Recordset
Private strSQL As String

Private A_FieldNameDisplay As String
Private A_FieldNameGet As String
Private A_FieldNameDefault As String   'Ĭ����ʾ���ֶ�����

Private A_bIsInit As Boolean
Private A_TipInit As String

Private WithEvents Combo1 As ComboBox
Attribute Combo1.VB_VarHelpID = -1




'�������ܣ�
'oCtl�����û�UI��Ҫ��ʾ���ݵĿؼ�����ַ
'vSQL����ȡ���ݵ�SQL
'vFieldNameDisplay����ʾ�ڿؼ��ϵ��ֶ�
'vFieldNameGet���û�ѡ��֮����ʵ�ʷ��ص��ֶ�����
'vFieldNameDefault:����Ĭ�ϱ�־���ֶε����ƣ���Ϊ�գ�
'vTip_Init: �ڳ�ʼ��ʧ�ܺ����ʾ��Ϣ (����Ϊ��)
Public Sub InitCls(ByRef oCtl As Object, ByVal vSQL As String, _
    ByVal vFieldNameDisplay As String, ByVal vFieldNameGet As String, _
    ByVal vFieldNameDefault As String, ByVal vTip_Init As String)
    
    Set oCombo = oCtl
    Set A_rs = New Recordset
    Debug.Print vSQL
    A_rs.Open vSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Debug.Print vSQL
    A_FieldNameDisplay = vFieldNameDisplay
    A_FieldNameGet = vFieldNameGet
    A_FieldNameDefault = vFieldNameDefault
    A_TipInit = vTip_Init
    
    
    A_bIsInit = True
End Sub

'���ݳ�ʼ�������еĲ�ѯSQL��ȡ��¼����CAPTION�ֶ���ʾ��UI�ϣ�����KEY�ֶ�
Public Sub ShowData()
    If A_bIsInit = False Then
        MsgBox A_TipInit, vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If A_rs.RecordCount <= 0 Then
        oCombo.Clear
        'MsgBox "��¼��Ϊ�գ�", vbOKOnly + vbInformation, "��ʾ"
        Exit Sub
    End If
    
    
    oCombo.Clear
    A_rs.MoveFirst
    oCombo.AddItem ""
    Do While Not A_rs.EOF
        oCombo.AddItem A_rs(Trim$(A_FieldNameDisplay))
        A_rs.MoveNext
    Loop
    
    
    '��ʾĬ��
    If Len(A_FieldNameDefault) > 0 Then
        A_rs.Filter = " " & A_FieldNameDefault & "=1"
        If A_rs.RecordCount > 0 Then
            oCombo.Text = A_rs(Trim$(A_FieldNameDisplay))
        End If
    End If
End Sub




'���ݵ�ǰUI����ʾ��CAPTION��ȡ��Ӧ��KEY�ֶε�VALUE
Public Function GetData() As String
    On Error Resume Next
    If A_bIsInit = False Then
        MsgBox A_TipInit, vbOKOnly + vbInformation, "��ʾ"
        Exit Function
    End If
    
    A_rs.Filter = ""
    If A_rs.RecordCount <= 0 Then
        GetData = ""
        Exit Function
    End If
    
    If Len(Trim$(oCombo.Text)) <= 0 Then
        GetData = ""
        Exit Function
    End If
    
    A_rs.Filter = ""
    '��2016��12��Ϊ���������ȡ���ݵ�ʱ�������˵�һ��Ϊ������
    '��ô֮���ȡѡ��VALUEʱ��Ͳ���+1�ˡ�
    'A_rs.Bookmark = Val(oCombo.ListIndex) + 1
    A_rs.Bookmark = Val(oCombo.ListIndex)
    Debug.Print A_rs(Trim$(A_FieldNameGet))
    GetData = A_rs(Trim$(A_FieldNameGet))
End Function

'��ȡComboBox�İ󶨵ļ�¼���У�������ʾ���û�����CAPTION�ֶκ������ֶ�֮�������ĳһ�ֶ�
Public Function GetData_OtherField(ByVal vFieldName As String) As String
    If A_bIsInit = False Then
        MsgBox A_TipInit, vbOKOnly + vbInformation, "��ʾ"
        Exit Function
    End If
    
    A_rs.Filter = ""
    If A_rs.RecordCount <= 0 Then
        Exit Function
    End If
    
    If Len(Trim$(oCombo.Text)) <= 0 Then
        GetData_OtherField = ""
        Exit Function
    End If
    
    A_rs.Filter = ""
    '��2017��4��6��Ϊ�´���¿ͻ���ʱ����index+1��Ĵ�����������
    'A_rs.Bookmark = Val(oCombo.ListIndex) + 1
    A_rs.Bookmark = Val(oCombo.ListIndex)
    GetData_OtherField = A_rs(Trim$(vFieldName))
End Function

'����KEY�ֶε�VALUE��vDesignation��֮����ʾ��CAPTION�ֶε�VALUE��UI��
'���ڴӳ־ò��ȡ���ݺ���ʾ�û�VALUE
'����OpenBill�У����־û����������ʾ���û�UI��
Public Sub ShowData_Designation(ByVal vDesignation As String)
    On Error Resume Next
    Dim iListIndex As Long
    
    A_rs.Filter = ""
    A_rs.MoveFirst
    oCombo.ListIndex = -1
    Do While Not A_rs.EOF
        If A_rs(Trim$(A_FieldNameGet)) = vDesignation Then
            iListIndex = Val(A_rs.Bookmark) - 1
            oCombo.Text = A_rs(Trim$(A_FieldNameDisplay))
            Exit Do
        End If
        A_rs.MoveNext
    Loop
    
End Sub

