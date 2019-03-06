VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmOrderProduct_FZ_Woven 
   Caption         =   "织标位置图"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOrderProduct_FZ_Woven.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8355
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _LayoutVersion  =   1
      _ExtentX        =   17595
      _ExtentY        =   14843
      _DataPath       =   ""
      Bands           =   "frmOrderProduct_FZ_Woven.frx":058A
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   8175
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   9375
         _cx             =   16536
         _cy             =   14420
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   800
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   6
         GridCols        =   4
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmOrderProduct_FZ_Woven.frx":0752
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   660
            Left            =   90
            ScaleHeight     =   660
            ScaleWidth      =   9195
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   90
            Width           =   9195
            Begin XtremeSuiteControls.PushButton PushButton3 
               Height          =   615
               Left            =   5160
               TabIndex        =   3
               Top             =   0
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   1085
               _StockProps     =   79
               Caption         =   "删除图片"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton PushButton2 
               Height          =   615
               Left            =   2880
               TabIndex        =   4
               Top             =   0
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   1085
               _StockProps     =   79
               Caption         =   "上传图片"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   615
               Left            =   720
               TabIndex        =   5
               Top             =   0
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   1085
               _StockProps     =   79
               Caption         =   "预览图片"
               UseVisualStyle  =   -1  'True
            End
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   0
               Top             =   0
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin XtremeSuiteControls.PushButton PushButton4 
               Height          =   615
               Left            =   7320
               TabIndex        =   6
               Top             =   0
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   1085
               _StockProps     =   79
               Caption         =   "退  出"
               UseVisualStyle  =   -1  'True
            End
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            Height          =   1095
            Index           =   1
            Left            =   90
            TabIndex        =   7
            Top             =   6990
            Width           =   9195
            _Version        =   1048578
            _ExtentX        =   16219
            _ExtentY        =   1931
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MultiLine       =   -1  'True
            ScrollBars      =   2
         End
         Begin VB.Image Image1 
            Height          =   6120
            Left            =   90
            Stretch         =   -1  'True
            Top             =   810
            Width           =   9195
         End
      End
   End
End
Attribute VB_Name = "frmOrderProduct_FZ_Woven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cls1 As New clsPicture
Private szFile As String

Public m_KuanHao As String

Public m_OrderID As Long   '存放表G_BillOrder 的主键字段

Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
End Sub

Private Sub Form_Load()
InitFrm

OpenImage

End Sub
'预览图片
Private Sub PushButton1_Click()
On Error GoTo IFERR
    
    With CommonDialog1
        .ShowOpen
   
        szFile = .FileName
    End With
     
    If Len(szFile) <= 0 Then
        Exit Sub
    End If
'    cls1.InitCls szFile, Picture5
     Image1.Picture = LoadPicture(szFile)
    
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "图片格式不支持，请提供JPG或者BMP格式图片！" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub
'上传图片
Private Sub PushButton2_Click()
Dim sql As String
Dim rs As New RecordSet

    If szFile <> "" Then
'                 需要引用：Microsoft Scripting Runtime
        Dim fso As New FileSystemObject
    
        Dim lLength As Long
        Dim oFile As File
    
        Set oFile = fso.GetFile(szFile)
        
        '设置上传图片的大小
        sql = "select * from G_ImageSize"
        rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
        
        
        If oFile.Size / 1000000 > rs!B_size Then
            MsgBox "图片太大不能上传", vbInformation, "提示"
            Exit Sub
        End If
    
        '获取的长度的单位是：字节
        saveImage
       
            MsgBox "图片上传成功", vbInformation, "提示"
       
    End If
End Sub
Private Sub saveImage()

    Dim rs As New RecordSet
    Dim sql As String
    sql = "SELECT * FROM WVAccountImage.dbo.G_image_NEW_Woven where 1=0"
    rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
    
     If szFile = "" Then
     
     Else
        Dim rs1 As New RecordSet
        Dim sql1 As String
        sql1 = "SELECT * FROM WVAccountImage.dbo.G_image_NEW_Woven where B_OrderID='" & m_OrderID & "'"
        rs1.Open sql1, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
        
        If rs1.RecordCount > 0 Then
            
            PicSaveToDB rs1!B_picture, szFile
            rs1.Update
            rs1.Close
            Set rs1 = Nothing
        Else
            rs.AddNew
'            rs!B_id = theid
            PicSaveToDB rs!B_picture, szFile
            rs!B_OrderID = m_OrderID '合同计划主键
            rs.Update
            rs.Close
            Set rs = Nothing
        End If
        
    End If
End Sub
'删除图片
Private Sub PushButton3_Click()
Dim sql As String
    If m_ID > 0 Then
        sql = "delete from WVAccountImage.dbo.G_image_NEW_Woven where B_OrderID='" & m_OrderID & "'"
        Gm.cnnToolImage.cnn.Execute sql
    
    End If
    Image1.Picture = Nothing
End Sub
'上传图片到服务器
'fld：记录集中的字段
'vFilePath：图片文件的绝对路径，包含图片文件名和扩展名
Private Sub PicSaveToDB(ByRef fld As ADODB.Field, ByVal vFilePath As String)
    Const blocksize = 4096
    Dim bytedata() As Byte
    Dim numblocks As Long
    Dim filelength As Long
    Dim leftover As Long
    Dim sourcefile As Long
    Dim i As Long
    sourcefile = FreeFile
    
    Open Trim(vFilePath) For Binary Access Read As sourcefile
    filelength = LOF(sourcefile)
    
    If filelength = 0 Then
        Close sourcefile
        'MsgBox Trim(vFilePath) & "无内容或不存在！"
    Else
        numblocks = filelength \ blocksize
        leftover = filelength Mod blocksize
        fld.Value = Null
        
        ReDim bytedata(blocksize)
        
        For i = 1 To numblocks
            Get sourcefile, , bytedata()
            fld.AppendChunk bytedata()
        Next
        

        ReDim bytedata(leftover)
        Get sourcefile, , bytedata()
        fld.AppendChunk bytedata()
        Close sourcefile
    End If
End Sub
'从DB中下载图片并且显示到UI的图片控件上
'vRs：包含有图片资源的数据源
'vPicField：保存图片文件的字段名
'oCtl：用于显示的控件。PictureBox、Image
Private Sub PicShow2Ctl(ByRef vFld As ADODB.Field, ByRef oCtl As Object)
    'On Error GoTo IFERR
    Dim Stream As ADODB.Stream
    Set Stream = New ADODB.Stream
    

    oCtl.Picture = LoadPicture("")
    If Not IsNull(vFld) Then
        Stream.Type = adTypeBinary
        Stream.Open
        Stream.Write vFld
        Stream.SaveToFile "filename", adSaveCreateOverWrite
        'Stream.SaveToFile "c:\aaa.jpg", adSaveCreateOverWrite
        
        szFile = LoadPicture("filename")
'        Debug.Print FileName
        oCtl.Picture = LoadPicture("filename")
        Stream.Close
    End If
    
    Set Stream = Nothing
'    Exit Sub
'IFERR:
'    Dim szErr As String
'    szErr = "错误发生于下载图片中，" & Err.Description
'    MsgBox szErr
End Sub
Private Sub OpenImage()

                Dim rs1 As New RecordSet
                Dim sql As String
                Dim clsFile01 As New clsFile
                Dim szPic As String
                
            sql = "SELECT * FROM WVAccountImage.dbo.G_image_NEW_Woven where B_OrderID='" & m_OrderID & "'"
            rs1.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
                
                If rs1.RecordCount > 0 Then
                    szPic = App.Path & "\temp\" & rs1!B_id & rs1!B_KuanHao & ".JPG"
                    Debug.Print szPic
                    
'                    clsFile01.DownloadPic rs1!B_picture, szPic
'                    cls1.InitCls szPic, frm1.Picture5
                    PicShow2Ctl rs1!B_picture, Image1
'                    PicShow2Ctl rs1!B_picture, Picture5
                    'PicShow2Ctl rs!B_picture, Picture3
                Else
                    Image1.Picture = Nothing
                End If

End Sub
'退出
Private Sub PushButton4_Click()
Unload Me
End Sub


