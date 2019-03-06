VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim rs As New Recordset
Private A_FileName As String
Private A_UpgradeForced As Long   '1��ʾǿ��Ҫ����,0��ʾ����ǿ�����û�ѡ��������
Private clsLog As New clsFile

Private Function NeedUpdate() As Boolean
    Dim strSQL As String
    Dim Version() As String


    NeedUpdate = False
    

    A_FileName = App.EXEName
    Set rs = New Recordset
    
    strSQL = "exec dbo.P_GetLastVersion '" & A_FileName & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        '��ǰ�������Ƿ�ǿ��ִ��.Ĭ��Ϊǿ��
        A_UpgradeForced = IIf(IsNull(rs!B_UpgradeForced), 1, rs!B_UpgradeForced)
    End If
    
    
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            Version = Split(rs("B_Version"), ".")
            Select Case Val(Version(0)) - Val(App.Major)
                Case Is > 0
                    NeedUpdate = True
                    Exit Function
                Case 0
                    Select Case Val(Version(1)) - Val(App.Minor)
                        Case Is > 0
                            NeedUpdate = True
                            Exit Function
                        Case 0
                            Select Case Val(Version(2)) - Val(App.Revision)
                                Case Is > 0
                                    NeedUpdate = True
                                    Exit Function
                                Case Is <= 0
                                    Exit Function
                            End Select
                        Case Is < 0
                            Exit Function
                    End Select
                Case Is < 0
                    Exit Function
            End Select
        End If
    End If
End Function

Public Sub CheckNewVersion(Handle As Long)
    Dim sHandle As String
    On Error Resume Next
    If Len(Dir(App.Path & "\*.tmp")) <> 0 Then Kill App.Path & "\*.tmp"
    sHandle = Trim(str(Handle))
    If NeedUpdate = True Then

        If A_UpgradeForced = 0 Then
            'ѡ��������
            If MsgBox("�����°汾���Ƿ���������", vbYesNo + vbQuestion, "������ʾ") = vbNo Then
                Exit Sub
            Else
                Call Shell(App.Path & "\�Զ����¿ͻ���.exe " & A_FileName)
                End
            End If
        Else
            'ǿ�ƽ�������
            Call Shell(App.Path & "\�Զ����¿ͻ���.exe " & A_FileName)
            End
        End If
    Else

    End If
End Sub


Private Sub Class_Terminate()
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
End Sub