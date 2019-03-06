VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)

Private mvarSysDir As String
Private mvarWinDir As String
Private mvarComputerName As String
Private mvarUserName As String
Private mvarSysRunTime As String
Private mvarResRat As String
Private mvarPhyMemory As String
Private m_IP As String
Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type

Public Property Get SysDir() As String '得到系统目录
    Dim buffer As String * 256
    Call GetSystemDirectory(buffer, 256)
    mvarSysDir = Left(buffer, InStr(1, buffer, vbNullChar) - 1)
    SysDir = mvarSysDir
End Property

Public Property Get WinDir() As String '得到Windows安装目录
    Dim buffer As String * 256
    Call GetWindowsDirectory(buffer, 256)
    mvarWinDir = Left(buffer, InStr(1, buffer, vbNullChar) - 1)
    WinDir = mvarWinDir
End Property

Public Property Get ComputerName() As String '得到计算机名字
    Dim buffer As String * 256
    Call GetComputerName(buffer, 256)
    mvarComputerName = Left(buffer, InStr(1, buffer, vbNullChar) - 1)
    ComputerName = mvarComputerName
End Property

Public Property Get UserName() As String '得到当前登录用户
    Dim buffer As String * 256
    Call GetUserName(buffer, 256)
    mvarUserName = Left(buffer, InStr(1, buffer, vbNullChar) - 1)
    UserName = mvarUserName
End Property

Public Property Get SysRunTime() As String '得到系统运行时间
    Dim RunH As String
    Dim RunM As String
    Dim RunS As String
    RunH = Int(GetTickCount \ 1000 \ 60 \ 60)
    RunM = Int(GetTickCount \ 1000 \ 60) - Int(GetTickCount \ 1000 \ 60 \ 60) * 60
    RunS = GetTickCount \ 1000 - RunH * 3600 - RunM * 60
    mvarSysRunTime = RunH & "小时" & RunM & "分钟" & RunS & "秒"
    SysRunTime = mvarSysRunTime
End Property

Public Property Get GetResRat() As String '得到屏幕分辨率
    mvarResRat = Screen.Width / Screen.TwipsPerPixelX & "*" & Screen.Height / Screen.TwipsPerPixelY
    GetResRat = mvarResRat
End Property

Public Property Get TotalPhyMemory() As String '得到计算机上安装的物理内存
    Dim TotalMemory As MEMORYSTATUS
    Call GlobalMemoryStatus(TotalMemory)
    mvarPhyMemory = TotalMemory.dwTotalPhys / 1024 & "KB"
    TotalPhyMemory = mvarPhyMemory
End Property

Public Function GetFileVerInfo(ByVal FullFileName As String) As String '得到文件版本号
   Dim rc As Long, lDummy As Long, sBuffer() As Byte
   Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
   Dim lVerbufferLen As Long

   '*** Get size ****
   lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
   If lBufferLen < 1 Then
      MsgBox "没有找到文件版本信息!"
      Exit Function
   End If

   '**** Store info to udtVerBuffer struct ****
   ReDim sBuffer(lBufferLen)
   rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
   rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
   MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)

   '**** Determine File Version number ****
   GetFileVerInfo = Format$(udtVerBuffer.dwFileVersionMSh) & "." & Format$(udtVerBuffer.dwFileVersionMSl) & "." & Format$(udtVerBuffer.dwFileVersionLSh) & "." & Format$(udtVerBuffer.dwFileVersionLSl)
End Function

Public Function GetIPAddress() As String
    On Error GoTo IFERR
    GetIPAddress = Gm.SysID.iP
    Exit Function
IFERR:
    GetIPAddress = ""
End Function

