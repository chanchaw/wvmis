VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ActiveReport3 
   Caption         =   "ActiveReport3"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11850
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   _ExtentX        =   20902
   _ExtentY        =   13309
   SectionData     =   "ActiveReport3.dsx":0000
End
Attribute VB_Name = "ActiveReport3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As RecordSet
Public itmeid As String
Public flowCardprint As Long
Private Sub ActiveReport_ReportStart()
    GetRs
End Sub

Private Sub GetRs()
   
    DataControl1.RecordSet = rs
End Sub
Private Sub ActiveReport_PrintProgress(ByVal pageNumber As Long)
        Dim sql As String
'        Dim a As String
'        a = IIf(IsNull(rs!B_print), 0, rs!B_print)
        sql = "update G_Billdetailwhite set B_print='" & flowCardprint & "'+1 where B_itemid='" & itmeid & "'"
        Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
End Sub
Private Sub CreateReportName()
    On Error Resume Next
    
    Dim oCtl As Object
    If Len(CompanyInfo_Name4Report) <= 0 Then
        Exit Sub
    End If
    
    For Each oCtl In Me.Sections("PageHeader").Controls
        If oCtl.name = "lblReportName" Then
            Debug.Print oCtl.Caption
            oCtl.Caption = CompanyInfo_Name4Report
            Debug.Print oCtl.Caption
            Exit Sub
        End If
    Next
    
    
    For Each oCtl In Me.Sections("ReportHeader").Controls
        If oCtl.name = "lblReportName" Then
            Debug.Print oCtl.Caption
            oCtl.Caption = CompanyInfo_Name4Report
            Debug.Print oCtl.Caption
            Exit Sub
        End If
    Next
End Sub


Private Sub ActiveReport_Activate()
    CreateReportName
End Sub

