VERSION 5.00
Object = "{0E8071F7-A7DD-47AC-95A9-365FFFE096DF}#1.0#0"; "UCButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectUser 
   BackColor       =   &H00CEDFDE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择用户"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   Icon            =   "frmSelectUser.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ListView List1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5106
      View            =   3
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16252927
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   2340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectUser.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TA_UCButton.UCButton Command2 
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   780
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      Caption         =   "取消  "
      Icon            =   "frmSelectUser.frx":0724
      IconMask        =   "frmSelectUser.frx":09BA
      CaptionAlignment=   1
   End
   Begin TA_UCButton.UCButton Command1 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   180
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      Caption         =   "确定  "
      Icon            =   "frmSelectUser.frx":0C50
      IconMask        =   "frmSelectUser.frx":0FEA
      CaptionAlignment=   1
   End
End
Attribute VB_Name = "frmSelectUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OK As Boolean
Public sUserList As String
Private Sub GetRecordSet()
    Dim rs As New Recordset
    Dim strSQL As String

    Set rs = New Recordset

    strSQL = "Select * From G_SystemUser"
    rs.Open strSQL, cnn.cnn, adOpenStatic, 3
    
    List1.ListItems.Clear
    List1.ColumnHeaders.Clear

    List1.ColumnHeaders.Add , , "用户名称", 3200
    Do While Not rs.EOF
        Set oItem = List1.ListItems.Add()
        With oItem
            .Text = rs("B_UserName")
            .SmallIcon = 1
        End With
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    sUserList = ""
    For i = 1 To List1.ListItems.Count
        If List1.ListItems(i).Checked = True Then
            sUserList = sUserList & List1.ListItems(i).Text & ","
        End If
    Next
    
    If Len(sUserList) < 1 Then
        MsgBox "您没有选择用户!", vbExclamation, "选择"
        Exit Sub
    End If
    sUserList = Mid(sUserList, 1, Len(sUserList) - 1)
    OK = True
    Me.Hide
End Sub

Private Sub Command2_Click()
    OK = False
    Me.Hide
End Sub

Private Sub Form_Load()
    GetRecordSet
    AnimateForm Me
End Sub

Private Sub List1_DblClick()
    On Error Resume Next
    If List1.SelectedItem.Checked = True Then
        List1.SelectedItem.Checked = False
    Else
        List1.SelectedItem.Checked = True
    End If
End Sub

