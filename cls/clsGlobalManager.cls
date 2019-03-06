VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobalManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private A_dmHardWareID As dmHardWareID
Private A_SysID As dmSysID
Private A_CompanyInfo As dmCompanyInfo
Private A_IniFile As clsIniFile

Private A_Connection As clsDBConnect   '本年度账套
Private A_ConnectionSOB As clsDBConnect '账套集数据库链接
Private A_ConnectionImage As clsDBConnect '图片数据库链接

Private A_frmMain As frmMain
Private A_CacheFrms As CacheFrm
Private A_Authority As clsCommand
Private A_iAnimate As Integer
Private A_PI As clspI
Private A_Business As New clsBusiness

Private A_DBCfgOneInt As clsDBCfgOneInt

Private A_OnlyDataBreak As Long

Public Property Get OnlyDataBreak() As Long
    OnlyDataBreak = A_OnlyDataBreak
End Property

Public Property Let OnlyDataBreak(ByVal vData As Long)
    A_OnlyDataBreak = vData
End Property

Public Property Set Business(ByVal vData As clsBusiness)
    Set A_Business = vData
End Property

Public Property Get Business() As clsBusiness
    Set Business = A_Business
End Property



Public Property Set DBCfgOneInt(ByVal vData As clsDBCfgOneInt)
    Set A_DBCfgOneInt = vData
End Property

Public Property Get DBCfgOneInt() As clsDBCfgOneInt
    Set DBCfgOneInt = A_DBCfgOneInt
End Property



Public Property Set PI(ByVal vData As clspI)
    Set A_PI = vData
End Property

Public Property Get PI() As clspI
    Set PI = A_PI
End Property



Public Property Let iAnimate(ByVal vData As Integer)
    A_iAnimate = vData
End Property

Public Property Get iAnimate() As Integer
    iAnimate = A_iAnimate
End Property



Public Property Set Authority(ByVal vData As clsCommand)
    Set A_Authority = vData
End Property

Public Property Get Authority() As clsCommand
    Set Authority = A_Authority
End Property


Public Property Set CacheFrms(ByVal vData As CacheFrm)
    Set A_CacheFrms = vData
End Property

Public Property Get CacheFrms() As CacheFrm
    Set CacheFrms = A_CacheFrms
End Property


Public Property Set frmMain(ByVal vData As frmMain)
    Set A_frmMain = vData
End Property

Public Property Get frmMain() As frmMain
    Set frmMain = A_frmMain
End Property



Public Property Set cnnTool(ByVal vData As clsDBConnect)
    Set A_Connection = vData
End Property

Public Property Get cnnTool() As clsDBConnect
    Set cnnTool = A_Connection
End Property

Public Property Set cnnToolSOB(ByVal vData As clsDBConnect)
    Set A_ConnectionSOB = vData
End Property

Public Property Get cnnToolImage() As clsDBConnect
    Set cnnToolImage = A_ConnectionImage
End Property

Public Property Set cnnToolImage(ByVal vData As clsDBConnect)
    Set A_ConnectionImage = vData
End Property

Public Property Get cnnToolSOB() As clsDBConnect
    Set cnnToolSOB = A_ConnectionSOB
End Property

Public Property Set IniFile(ByVal vData As clsIniFile)
    Set A_IniFile = vData
End Property

Public Property Get IniFile() As clsIniFile
    Set IniFile = A_IniFile
End Property



Public Property Set HardWareID(ByVal vData As dmHardWareID)
    Set A_dmHardWareID = vData
End Property

Public Property Get HardWareID() As dmHardWareID
    Set HardWareID = A_dmHardWareID
End Property


Public Property Set SysID(ByVal vData As dmSysID)
    Set A_SysID = vData
End Property

Public Property Get SysID() As dmSysID
    Set SysID = A_SysID
End Property

Public Property Set CompanyInfo(ByVal vData As dmCompanyInfo)
    Set A_CompanyInfo = vData
End Property

Public Property Get CompanyInfo() As dmCompanyInfo
    Set CompanyInfo = A_CompanyInfo
End Property


Public Function IsDeveloper() As Boolean
    IsDeveloper = False
    
    Exit Function
    
    If UCase(Gm.SysID.ComputerName) = "CHANCHAW-LENOVO" Then
        IsDeveloper = True
    End If
    
    If Gm.SysID.SystemUser = "管理员" Then
        IsDeveloper = True
    End If
End Function

Public Sub RefreshCnn()
    A_Connection.IniConnection8DM A_SysID.DBInfo
End Sub



'判断一个数组是否为空
'返回TRUE表示数组有元素
'反之则没有
Public Function ArrNotEmpty(ByVal vArr As Variant) As Boolean
    Dim cls1 As New clsArray
    ArrNotEmpty = cls1.ArrNotEmpty(vArr)
End Function

'打开一个对象
Public Sub OpenBL(ByVal m_ObjectID As String, ByVal m_BillName As String, ByVal vID As String)
    A_Authority.Execute m_ObjectID, m_BillName, "EditObject", Nothing, vID
End Sub


Public Sub log4Runtime(ByVal vContent As String)
    If g_SwitchLog4Runtime <> 1 Then
        Exit Sub
    End If
    
    
    Dim cls1 As New clsFile
    cls1.WriteFileContent g_LogFile_Runtime, vContent
End Sub

