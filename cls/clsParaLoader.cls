VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParaLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'参数加载器
Public Sub LoadPara()
    '获取计算机名、用户名、IP地址
    GetHardWareID
    
    '加载用户权限
    LoadAuthority
    
    LoadFrmCache
    
    '加载DB端配置文件G_Config_OneInt中的数据
    LoadDBCfgOneInt
End Sub

Private Sub GetHardWareID()
    Dim cls1 As New clsHardWareID
    
    With Gm.SysID
        .ComputerName = cls1.GetComputerName
        .ComputerUserName = cls1.GetComputerUserName
        .iP = cls1.GetIP
    End With
End Sub

Private Sub LoadFrmCache()
    Dim oCache As New CacheFrm
    Set Gm.CacheFrms = oCache
End Sub

'加载用户权限
Private Sub LoadAuthority()
    Dim cls1 As New clsCommand
    cls1.InitClass
    
    Set Gm.Authority = cls1
    
    
    Dim cls2 As New clspI
    cls2.InitCls
    
    Set Gm.PI = cls2
End Sub

'加载DB端的配置数据
Private Sub LoadDBCfgOneInt()
    Dim cls1 As New clsDBCfgOneInt
    Set Gm.DBCfgOneInt = cls1
End Sub




