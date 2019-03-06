VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmOrderProduct_FZ 
   Caption         =   "水洗标内容"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOrderProduct_FZ.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8085
   ScaleWidth      =   14715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14415
      _LayoutVersion  =   1
      _ExtentX        =   25426
      _ExtentY        =   13785
      _DataPath       =   ""
      Bands           =   "frmOrderProduct_FZ.frx":058A
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   7575
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   14055
         _cx             =   24791
         _cy             =   13361
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
         _GridInfo       =   $"frmOrderProduct_FZ.frx":0752
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture3 
            Height          =   4530
            Left            =   90
            ScaleHeight     =   7.885
            ScaleMode       =   7  'Centimeter
            ScaleWidth      =   5.927
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3420
            Begin VB.Image Image1 
               Height          =   4455
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   3480
            End
         End
         Begin VB.PictureBox Picture4 
            Height          =   4530
            Left            =   3570
            ScaleHeight     =   7.885
            ScaleMode       =   7  'Centimeter
            ScaleWidth      =   5.9
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3405
            Begin VB.Image Image2 
               Height          =   4455
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   3375
            End
         End
         Begin VB.PictureBox Picture6 
            Height          =   4530
            Left            =   10530
            ScaleHeight     =   7.885
            ScaleMode       =   7  'Centimeter
            ScaleWidth      =   5.953
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3435
            Begin VB.Image Image4 
               Height          =   4455
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   3495
            End
         End
         Begin XtremeSuiteControls.PushButton PushButton7 
            Height          =   450
            Left            =   10530
            TabIndex        =   12
            Top             =   1410
            Width           =   3435
            _Version        =   1048578
            _ExtentX        =   6059
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "预览图片"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton6 
            Height          =   450
            Left            =   7035
            TabIndex        =   11
            Top             =   1410
            Width           =   3435
            _Version        =   1048578
            _ExtentX        =   6059
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "预览图片"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton5 
            Height          =   450
            Left            =   3570
            TabIndex        =   10
            Top             =   1410
            Width           =   3405
            _Version        =   1048578
            _ExtentX        =   6006
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "预览图片"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   585
            Left            =   90
            ScaleHeight     =   585
            ScaleWidth      =   13875
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   765
            Width           =   13875
            Begin XtremeSuiteControls.FlatEdit FlatEdit1 
               Height          =   615
               Index           =   0
               Left            =   1200
               TabIndex        =   15
               Top             =   0
               Width           =   2295
               _Version        =   1048578
               _ExtentX        =   4048
               _ExtentY        =   1085
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "款号："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   120
               TabIndex        =   7
               Top             =   120
               Width           =   1080
            End
         End
         Begin VB.PictureBox Picture5 
            Height          =   4530
            Left            =   7035
            ScaleHeight     =   7.885
            ScaleMode       =   7  'Centimeter
            ScaleWidth      =   5.953
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3435
            Begin VB.Image Image3 
               Height          =   4455
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   3495
            End
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   90
            ScaleHeight     =   615
            ScaleWidth      =   13875
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   90
            Width           =   13875
            Begin XtremeSuiteControls.PushButton PushButton3 
               Height          =   615
               Left            =   3720
               TabIndex        =   4
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
               Left            =   1320
               TabIndex        =   3
               Top             =   0
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   1085
               _StockProps     =   79
               Caption         =   "上传图片"
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
               Left            =   5880
               TabIndex        =   8
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
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   450
            Left            =   90
            TabIndex        =   9
            Top             =   1410
            Width           =   3420
            _Version        =   1048578
            _ExtentX        =   6032
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "预览图片"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            Height          =   975
            Index           =   1
            Left            =   90
            TabIndex        =   16
            Top             =   6510
            Width           =   3420
            _Version        =   1048578
            _ExtentX        =   6032
            _ExtentY        =   1720
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
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            Height          =   975
            Index           =   2
            Left            =   3570
            TabIndex        =   17
            Top             =   6510
            Width           =   3405
            _Version        =   1048578
            _ExtentX        =   6006
            _ExtentY        =   1720
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
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            Height          =   975
            Index           =   3
            Left            =   7035
            TabIndex        =   18
            Top             =   6510
            Width           =   3435
            _Version        =   1048578
            _ExtentX        =   6059
            _ExtentY        =   1720
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
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            Height          =   975
            Index           =   4
            Left            =   10530
            TabIndex        =   19
            Top             =   6510
            Width           =   3435
            _Version        =   1048578
            _ExtentX        =   6059
            _ExtentY        =   1720
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
      End
   End
End
Attribute VB_Name = "frmOrderProduct_FZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cls1 As New clsPicture
Private clsPicture1 As New clsPicture
Private szFile As String
Private szFile1 As String
Private szFile2 As String
Private szFile3 As String
Private szFile4 As String

Public m_KuanHao As String
Public m_ID As Long    '存放表  G_BillSew 的 主键字段
Public m_OrderID As Long   '存放表G_BillOrder 的主键字段

Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
End Sub



Private Sub Form_Load()
InitFrm

OpenImage Image1, 1
OpenImage Image2, 2
OpenImage Image3, 3
OpenImage Image4, 4

End Sub
'上传图片
Private Sub PushButton2_Click()

   JudegSize szFile1, 1
   JudegSize szFile2, 2
   JudegSize szFile3, 3
   JudegSize szFile4, 4
   
   MsgBox "图片上传成功", vbInformation, "提示"
End Sub

Private Sub JudegSize(ByVal VszFile As String, ByVal vNumber As Long)
Dim sql As String
Dim rs As New RecordSet

    If VszFile <> "" Then
'                 需要引用：Microsoft Scripting Runtime
        Dim fso As New FileSystemObject
    
        Dim lLength As Long
        Dim oFile As File
    
        Set oFile = fso.GetFile(VszFile)
        
        '设置上传图片的大小
        sql = "select * from G_ImageSize"
        rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
        
        
        If oFile.Size / 1000000 > rs!B_size Then
            MsgBox "图片太大不能上传", vbInformation, "提示"
            Exit Sub
        End If
  End If
    saveImage VszFile, vNumber
  
 End Sub

Private Sub saveImage(ByVal VszFile As String, ByVal vNumber As Long)

    Dim rs As New RecordSet
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim sql1 As String
    sql = "SELECT * FROM WVAccountImage.dbo.G_image_NEW_FZ where 1=0"
    rs.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
    
     If VszFile = "" Then
     '在不修改图片的情况下只修改备注和款号
     sql1 = "SELECT * FROM WVAccountImage.dbo.G_image_NEW_FZ "
        sql1 = sql1 & " where B_OrderID='" & m_OrderID & "' and B_BDCItemID='" & m_ID & "' and B_Number='" & vNumber & "'"
        rs1.Open sql1, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
         If rs1.RecordCount > 0 Then
            rs1!B_KuanHao = FlatEdit1(0).Text
            rs1!B_memo = FlatEdit1(vNumber).Text
            rs1.Update
            rs1.Close
            Set rs1 = Nothing
        Else
            rs.AddNew
            rs!B_KuanHao = FlatEdit1(0).Text
            rs!B_memo = FlatEdit1(vNumber).Text
            rs!B_Number = vNumber
            rs.Update
            rs.Close
            Set rs = Nothing
        End If
        
        
     Else
        sql1 = "SELECT * FROM WVAccountImage.dbo.G_image_NEW_FZ "
        sql1 = sql1 & " where B_OrderID='" & m_OrderID & "' and B_BDCItemID='" & m_ID & "' and B_Number='" & vNumber & "'"
        rs1.Open sql1, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
        
        If rs1.RecordCount > 0 Then
            
            PicSaveToDB rs1!B_picture, VszFile
            rs1!B_KuanHao = FlatEdit1(0).Text
            rs1!B_Number = vNumber
            rs1!B_memo = FlatEdit1(vNumber).Text
            rs1.Update
            rs1.Close
            Set rs1 = Nothing
        Else
            rs.AddNew
'            rs!B_id = theid
            PicSaveToDB rs!B_picture, VszFile
            rs!B_BDCItemID = m_ID  '缝制计划表 一个主键对应一张图片
            rs!B_OrderID = m_OrderID '合同计划主键
            rs!B_KuanHao = FlatEdit1(0).Text
            rs!B_memo = FlatEdit1(vNumber).Text
            rs!B_Number = vNumber
            rs.Update
            rs.Close
            Set rs = Nothing
        End If
        
'         MsgBox "图片上传成功", vbInformation, "提示"
    End If
End Sub
'删除图片
Private Sub PushButton3_Click()
Dim sql As String
    If m_ID > 0 Then
        sql = "delete from WVAccountImage.dbo.G_image_NEW_FZ where B_OrderID='" & m_OrderID & "' and B_BDCItemID='" & m_ID & "'"
        Gm.cnnToolImage.cnn.Execute sql
    
    End If
    Image1.Picture = Nothing
    Image2.Picture = Nothing
    Image3.Picture = Nothing
    Image4.Picture = Nothing
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

Private Sub OpenImage(ByRef vPicture As Image, ByVal vNumber As Long)
'PictureBox
                Dim rs1 As New RecordSet
                Dim sql As String
                Dim clsFile01 As New clsFile
                Dim szPic As String
                
                sql = "SELECT * FROM WVAccountImage.dbo.G_image_NEW_FZ "
                sql = sql & " where B_OrderID='" & m_OrderID & "' and B_BDCItemID='" & m_ID & "' and B_Number='" & vNumber & "'"
                rs1.Open sql, Gm.cnnToolImage.cnn, adOpenKeyset, adLockPessimistic
                
                If rs1.RecordCount > 0 Then
                    szPic = App.Path & "\temp\" & rs1!B_id & rs1!B_KuanHao & ".JPG"
                    Debug.Print szPic
                    
'                    clsFile01.DownloadPic rs1!B_picture, szPic
'                    cls1.InitCls szPic, frm1.Picture5

                    FlatEdit1(0).Text = rs1!B_KuanHao
                    FlatEdit1(vNumber).Text = IIf(IsNull(rs1!B_memo), "", rs1!B_memo)
                    PicShow2Ctl rs1!B_picture, vPicture
                    'PicShow2Ctl rs!B_picture, Picture3
                Else
                    vPicture.Picture = Nothing
                End If

End Sub

'退出
Private Sub PushButton4_Click()
Unload Me
End Sub
'预览图片1
Private Sub PushButton1_Click()
'On Error GoTo IFERR
    
    With CommonDialog1
        .ShowOpen
   
        szFile1 = .FileName
    End With
     
    If Len(szFile1) <= 0 Then
        Exit Sub
    End If
    'cls1.InitCls szFile1, Image1
    'cls1.InitCls szFile1, Picture3
    
    Image1.Picture = LoadPicture(szFile1)
    
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "图片格式不支持，请提供JPG或者BMP格式图片！" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub
'预览图片2
Private Sub PushButton5_Click()
On Error GoTo IFERR
    With CommonDialog1
        .ShowOpen
        szFile2 = .FileName
    End With
    If Len(szFile2) <= 0 Then
        Exit Sub
    End If
'    cls1.InitCls szFile2, Image2
    Image2.Picture = LoadPicture(szFile2)
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "图片格式不支持，请提供JPG或者BMP格式图片！" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub
'预览图片3
Private Sub PushButton6_Click()
On Error GoTo IFERR
    With CommonDialog1
        .ShowOpen
        szFile3 = .FileName
    End With
    If Len(szFile3) <= 0 Then
        Exit Sub
    End If
'    cls1.InitCls szFile3, Image3
     Image3.Picture = LoadPicture(szFile3)
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "图片格式不支持，请提供JPG或者BMP格式图片！" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub
'预览图片4
Private Sub PushButton7_Click()
On Error GoTo IFERR
    With CommonDialog1
        .ShowOpen
        szFile4 = .FileName
    End With
    If Len(szFile4) <= 0 Then
        Exit Sub
    End If
'    cls1.InitCls szFile4, Image4
    Image4.Picture = LoadPicture(szFile4)
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "图片格式不支持，请提供JPG或者BMP格式图片！" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub
