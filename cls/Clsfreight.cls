VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clsfreight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim rs As RecordSet

'�ж������Ƿ���Ԫ��
Public Function Freight(ByVal vArr As Variant) As Boolean

    Set rs = New RecordSet
    Dim sql As String
    sql = "select * from G_Freightdetail where B_Codeid='" & vArr & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        Freight = False
        Exit Function
    End If
    Freight = True

End Function

'�жϸ߳��ƻ��ĸ�����
Public Function blv(ByVal a As String, ByVal b As Long) As Boolean
    blv = True
    Dim sql As String
    Dim rs As RecordSet
    Set rs = New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet
    
    sql = "SELECT SUM(B_kg) as  B_kg FROM G_billdetailorder WHERE B_ID='" & a & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    sql2 = "select * from G_Config_OneFloat where B_groupname='֯��ϵͳ_�ƻ�������'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Debug.Print sql2
    If rs!B_kg > 0 Then
        If (b - rs!B_kg) / rs!B_kg > rs2!B_Value Then
            blv = False
        End If
    End If
End Function
