VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ActiveReport6 
   Caption         =   "ActiveReport6"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11385
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   _ExtentX        =   20082
   _ExtentY        =   12965
   SectionData     =   "ActiveReport6.dsx":0000
End
Attribute VB_Name = "ActiveReport6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
