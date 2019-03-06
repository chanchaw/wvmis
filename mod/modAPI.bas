Attribute VB_Name = "modAPI"
Option Explicit

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal name As String, ByVal level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long


'硬件ID识别API
'===============
Public Declare Function GetHardWareID Lib "HardwareID.DLL" Alias "GetHardwareID" (ByVal HDD As Boolean, ByVal NIC As Boolean, ByVal CPU As Boolean, ByVal BIOS As Boolean, ByVal RegCode As String) As String
Public Declare Function GetHardwareIDWithAppID Lib "HardwareID.DLL" (ByVal AppID As String, ByVal HDD As Boolean, ByVal NIC As Boolean, ByVal CPU As Boolean, ByVal BIOS As Boolean, ByVal RegCode As String) As String
'===============


Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

'将窗口显示在最前
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Public Sub BringWindow2Top(ByVal hwnd As Long)
    BringWindowToTop hwnd
End Sub
