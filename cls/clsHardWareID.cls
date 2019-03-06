VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHardWareID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'获取当前登录电脑的用户名
Public Function GetComputerUserName() As String
    GetComputerUserName = Environ("username")
End Function

'获取电脑名
Public Function GetComputerName() As String
    GetComputerName = Environ("computername")
End Function

'获取电脑IP地址
Public Function GetIP() As String
    GetIP = GetIPAddress
End Function
