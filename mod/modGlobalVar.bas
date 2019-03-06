Attribute VB_Name = "modGlobalVar"
Option Explicit

'全局变量
Public Gm As New clsGlobalManager
Public Lanucher As New clsLanucher
Public ParaLoader As New clsParaLoader
Public strSQL As String
Public CompanyInfo_Name4Report As String
Public CompanyInfo_AppTitle As String
Public iAnimate  As Long '是否开启动画
Public g_FunctTool As New FunctTool

Public g_CJSuite As New clsCJSuite 'Codejock套件工具类
Public g_SwitchLog4Runtime As Long
Public g_LogFile_Runtime As String
