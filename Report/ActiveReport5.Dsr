VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ActiveReport5 
   Caption         =   "ActiveReport5"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11505
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   _ExtentX        =   20294
   _ExtentY        =   13600
   SectionData     =   "ActiveReport5.dsx":0000
End
Attribute VB_Name = "ActiveReport5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private A_rsColor As RecordSet
Private sql As String
'Private Sub ActiveReport_ReportStart()
'    GetRs
'End Sub
'
'Private Sub GetRs()
'    With DataControl1
'        .ConnectionString = Gm.cnnTool.cnnStr
'        .Source = sql
''        Debug.Print sql
'        .Refresh
'    End With
'    Set A_rsColor = New RecordSet
'    A_rsColor.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
'End Sub
'


Public Sub CreateReportName()
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
    
    
'    For Each oCtl In Me.Sections("Detail").Controls
'        If oCtl.name = "lblReportName" Then
'            Debug.Print oCtl.Caption
'            oCtl.Caption = CompanyInfo_Name4Report
'            Debug.Print oCtl.Caption
'            Exit Sub
'        End If
'    Next
End Sub


Private Sub ActiveReport_Activate()
    CreateReportName
End Sub

