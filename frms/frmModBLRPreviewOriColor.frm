VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmModBLRPreviewOriColor 
   Caption         =   "打印预览"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   7785
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7785
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      _LayoutVersion  =   1
      _ExtentX        =   21828
      _ExtentY        =   13732
      _DataPath       =   ""
      Bands           =   "frmModBLRPreviewOriColor.frx":0000
      Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer21 
         Height          =   5595
         Left            =   300
         TabIndex        =   1
         Top             =   540
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   9869
         SectionData     =   "frmModBLRPreviewOriColor.frx":1834
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   5040
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin XtremeSuiteControls.CommonDialog CommonDialog1 
         Left            =   720
         Top             =   6240
         _Version        =   1048578
         _ExtentX        =   423
         _ExtentY        =   423
         _StockProps     =   4
      End
   End
End
Attribute VB_Name = "frmModBLRPreviewOriColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private A_rs As New RecordSet
Private A_ObjectID As String
Public obj As String
'Private rpt As New DDActiveReports2.ActiveReport
'Private rpt As New ActiveReport1
Private rpt As New ActiveReport5

Public Property Set RecordSet(ByRef vData As RecordSet)
    Set A_rs = vData
End Property

Public Property Get RecordSet() As RecordSet
    Set RecordSet = A_rs
End Property

Public Property Let ObjectID(ByVal vData As String)
    A_ObjectID = vData
End Property

Public Property Get ObjectID() As String
    ObjectID = A_ObjectID
End Property


Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = ARViewer21
        .RecalcLayout
    End With
    
    PreviewReport
End Sub

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "刷新"
            PreviewReport
        Case "到PDF"
            ExportPDF
        Case "到Excel"
            ExportExcel
        Case "到无格式Excel"
            Export2ExcelOri (obj)
        Case "关闭"
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    InitFrm
End Sub

Private Sub PreviewReport()
    

    Dim cls1 As New clsPrint
    Dim szReportFile As String
    
    szReportFile = cls1.DownloadReport(ObjectID)

'    Dim DataControl1 As Object
'    Set DataControl1 = rpt.Sections("Detail").Controls.Add("DDActiveReports2.DAODataControl")
'
'    With rpt
'        '.Refresh
'        Set .DataControl1.RecordSet = RecordSet.Clone
'        .LoadLayout szReportFile
'        '.Show
'    End With
    
    Set rpt = New ActiveReport5
    Set rpt.DataControl1.RecordSet = RecordSet.Clone
    rpt.LoadLayout szReportFile
    rpt.CreateReportName
    Set ARViewer21.ReportSource = rpt
    
    
    rpt.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload rpt
End Sub


Private Sub ExportPDF()
    On Error Resume Next
    Dim pdf As ARExportPDF
    Set pdf = New ARExportPDF
    
    Dim cls1 As New clsFile
    
    pdf.FileName = cls1.ShowSaveFileDialog("PDF文件(*.pdf)|*.pdf")
    
    pdf.Export ARViewer21.Pages
    
    
    MsgBox "已导出为PDF文件：" & vbNewLine & pdf.FileName
    
    Set pdf = Nothing
End Sub

Private Sub ExportExcel()
    On Error Resume Next
    Dim excel As ARExportExcel
    Set excel = New ARExportExcel
    
    'Dim cls1 As New clsFile
    
    'excel.FileName = cls1.ShowSaveFileDialog("Excel文件(*.xls)|*.xls")
    
    '使用本地控件弹出保存对话框
    With CommonDialog1
        .Filter = "Excel文件(*.xls)|*.xls"
        .ShowSave
        excel.FileName = .FileName
    End With
    
    excel.Export ARViewer21.Pages
    
    
    MsgBox "已导出为Excel文件：" & vbNewLine & excel.FileName
    
    Set excel = Nothing
End Sub

'使用记录集导出数据到EXCEL
Private Sub Export2ExcelOri(ByVal obj As String)
    
    On Error Resume Next
    If A_rs.State <> adStateOpen Then
        MsgBox "没有数据不可导出！", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
    If A_rs.RecordCount <= 0 Then
        MsgBox "没有数据不可导出！", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    
'    Dim xlApp As excel.Application
'    Dim xlBook As excel.Workbook
'    Dim xlSheet As excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.add
    Set xlSheet = xlBook.Worksheets(1)
    xlApp.Visible = True '显示表格
    
    Dim i As Long
    Dim j As Long
    Dim f As Long
    
    j = 1
    
    Dim rs As New RecordSet
    strSQL = "SELECT * FROM G_BLSField AS gb WHERE gb.B_ObjectID='" & obj & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    
    i = 1
    rs.MoveFirst
    Do While Not rs.EOF
        
        xlSheet.Cells(1, i).Value = rs!B_CnName
        i = i + 1
        rs.movenext
    Loop
    
    
    i = 1
    f = 0
    A_rs.MoveFirst
    Do While Not A_rs.EOF
        rs.MoveFirst
        Do While Not rs.EOF
            xlSheet.Cells(j + 1, f + 1).Value = A_rs(Trim(rs!B_FieldName))
            rs.movenext
            f = f + 1
        Loop
            
        f = 0
        A_rs.movenext
        j = j + 1
    Loop
    
    
    'xlBook.Save '保存"
    Set xlApp = Nothing '交还控制给Excel
    
    rs.Close
    Set rs = Nothing
End Sub

