VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ActiveReport7 
   Caption         =   "ActiveReport7"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12645
   StartUpPosition =   3  '����ȱʡ
   _ExtentX        =   22304
   _ExtentY        =   12938
   SectionData     =   "ActiveReport7.dsx":0000
End
Attribute VB_Name = "ActiveReport7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sql As String
Public id As String
Public itemid As String
Public departdanprint As String
Private A_rsColor As RecordSet
Public color As String
Public depart As String

Private Sub ErgodicColor()
    On Error Resume Next
    Dim oCtl As Object

    For Each oCtl In Me.Sections("Detail").Controls
        If oCtl.name = "Field21" Then
            
            oCtl.BackColor = A_rsColor!B_hex
            oCtl.ForeColor = A_rsColor!B_hex
            Debug.Print oCtl.name & "=" & oCtl.Text
            A_rsColor.movenext
            
        End If
    Next
    
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)
    ErgodicColor
End Sub
Private Sub ActiveReport_ReportStart()
    GetRs
End Sub

Private Sub GetRs()
    With DataControl1
        .ConnectionString = Gm.cnnTool.cnnStr
        .Source = sql
        Debug.Print sql
        .Refresh
    End With
    Set A_rsColor = New RecordSet
    A_rsColor.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
End Sub

Private Sub ActiveReport_PrintProgress(ByVal pageNumber As Long)
        Dim sql As String
        sql = "update G_BilldetailColor set B_departdanprint='" & departdanprint & "'+1 where B_ID='" & id & "'  "
'        If Trim(itemid) > 0 Then
        sql = sql & " and B_depart='" & depart & "'"
'        End If
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


