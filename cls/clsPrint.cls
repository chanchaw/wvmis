VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String

'打印
Public Function PrintRP(ByRef oRs As RecordSet, ByVal vPrinter As String, _
    ByVal vReportObjectID As String, Optional ByVal vPreview As Boolean = False) As Boolean

    On Error GoTo IFERR
    
    Dim rpt1 As New ActiveReport1
    Dim szFile As String
    szFile = DownloadReport(vReportObjectID)
    With rpt1
        .Refresh
        .DataControl1.RecordSet = oRs
        .LoadLayout szFile
        .Printer.DeviceName = vPrinter
        .PrintReport vPreview  'FALSE为不预览直接打印
    End With
    
    Set rpt1 = Nothing
    
    PrintRP = True
    Exit Function
IFERR:
    PrintRP = False
End Function


'返回创建的本地文件的绝对路径
Public Function DownloadReport(ByVal vObjectID As String) As String
    Dim szFile As String
    szFile = App.Path & "\" & vObjectID & ".xml"
    
    '先删除本地文件
    DelFile szFile
    
    Dim rs As New RecordSet
    strSQL = "exec dbo.P_GetReport '" & vObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount <= 0 Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    Dim byteFile() As Byte
    byteFile = rs!B_ReportImage
    
    Open szFile For Binary As #1
    Put #1, , byteFile()
    Close #1
    
    DownloadReport = szFile
End Function

'删除本地所有XML文件
Private Sub DelFile(ByVal vFile As String)
    On Error Resume Next
    Dim cls1 As New clsFile
    cls1.DeleteFile vFile
End Sub
