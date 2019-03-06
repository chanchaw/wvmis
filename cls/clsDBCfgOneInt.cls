VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBCfgOneInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'DB端的配置数据
Private A_rs As New RecordSet
Private A_UsingFineGrainedAuthority As Boolean

Public Property Get UsingFineGrainedAuthority() As Boolean
    UsingFineGrainedAuthority = A_UsingFineGrainedAuthority
End Property


Private Sub Class_Initialize()
    Dim ltemp As Long
    
    strSQL = "SELECT * FROM G_Config_OneInt WHERE 1=1"
    Set A_rs = New RecordSet
    A_rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    
    '是否启用细颗粒度权限系统
    '1表示启用，其他数值都是不使用
    A_rs.Filter = " B_GroupName='系统权限_启用细颗粒度权限系统'"
    If A_rs.RecordCount > 0 Then
        ltemp = IIf(IsNull(A_rs!B_Value), 0, A_rs!B_Value)
        
        If ltemp = 1 Then
            A_UsingFineGrainedAuthority = True
        Else
            A_UsingFineGrainedAuthority = False
        End If
        
    Else
        A_UsingFineGrainedAuthority = False
    End If
    
    
    
    A_rs.Close
    Set A_rs = Nothing
End Sub

