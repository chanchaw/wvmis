VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ActiveReport2 
   Caption         =   "ActiveReport2"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11535
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   _ExtentX        =   20346
   _ExtentY        =   13441
   SectionData     =   "ActiveReport2.dsx":0000
End
Attribute VB_Name = "ActiveReport2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rs As RecordSet
Public itmeid As String
Public flowCardprint As Long

Private Sub ActiveReport_PrintProgress(ByVal pageNumber As Long)
        Dim sql As String
        sql = "update G_BilldetailColor set B_flowCardprint='" & flowCardprint & "'+1 where B_itemid='" & itmeid & "'"
        Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
End Sub

Private Sub ActiveReport_ReportStart()
    GetRS
End Sub

Private Sub GetRS()
   
    DataControl1.RecordSet = rs
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

