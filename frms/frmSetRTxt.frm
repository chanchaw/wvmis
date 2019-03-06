VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSetRTxt 
   Caption         =   "编辑RTF"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   Icon            =   "frmSetRTxt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   8640
   StartUpPosition =   2  '屏幕中心
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   6540
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8640
      _cx             =   15240
      _cy             =   11536
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
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
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   0
      ChildSpacing    =   0
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
      GridRows        =   3
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmSetRTxt.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   8640
         _LayoutVersion  =   1
         _ExtentX        =   15240
         _ExtentY        =   661
         _DataPath       =   ""
         Bands           =   "frmSetRTxt.frx":03E2
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   7260
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin RichTextLib.RichTextBox rtf 
         Height          =   5520
         Left            =   0
         TabIndex        =   5
         Top             =   375
         Width           =   8640
         _ExtentX        =   15240
         _ExtentY        =   9737
         _Version        =   393217
         BackColor       =   10286079
         HideSelection   =   0   'False
         ScrollBars      =   2
         TextRTF         =   $"frmSetRTxt.frx":3B96
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   0
         ScaleHeight     =   645
         ScaleWidth      =   4605
         TabIndex        =   4
         Top             =   5895
         Width           =   4605
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   4605
         ScaleHeight     =   645
         ScaleWidth      =   4035
         TabIndex        =   1
         Top             =   5895
         Width           =   4035
         Begin VB.CommandButton Command1 
            Caption         =   "确定"
            Height          =   435
            Left            =   900
            TabIndex        =   3
            Top             =   120
            Width           =   1395
         End
         Begin VB.CommandButton Command2 
            Caption         =   "关闭"
            Height          =   435
            Left            =   2520
            TabIndex        =   2
            Top             =   120
            Width           =   1395
         End
      End
   End
End
Attribute VB_Name = "frmSetRTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OK As Boolean
Public mRTxt As String

Private Sub ActiveBar21_ComboSelChange(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "tFont"
            rtf.SelFontName = Tool.Text
        Case "tFontSize"
            rtf.SelFontSize = Tool.Text
    End Select
    
    'rtf.SetFocus
End Sub

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "tOpen"
            OpenRTFFile
        Case "tCut"
            Clipboard.SetText rtf.SelRTF
            rtf.SelText = vbNullString
        Case "tCopy"
            Clipboard.SetText rtf.SelRTF
        Case "tPaste"
            rtf.SelRTF = Clipboard.GetText
        Case "tBold"
            rtf.SelBold = True
        Case "tItalic"
            rtf.SelItalic = True
        Case "tUnderline"
            rtf.SelUnderline = True
        Case "tLeft"
            rtf.SelAlignment = rtfLeft
        Case "tCenter"
            rtf.SelAlignment = rtfCenter
        Case "tRight"
            rtf.SelAlignment = rtfRight
    End Select
    
    'Set focus back to rtf control
    rtf.SetFocus
End Sub

Private Sub OpenRTFFile()
    Dim sFilename As String
    With CommonDialog1
        .DialogTitle = "读取文件"
        .CancelError = False
        'ToDo: 设置 common dialog 控件的标志和属性
        .Filter = "(*.*)|*.*"
        .ShowOpen
         sFilename = .FileName
        .FileName = ""
    End With
    If Len(Trim(sFilename)) > 1 Then
        rtf.LoadFile sFilename, 1
    End If
End Sub

Private Sub InitActiveBar()
    Dim i As Long
    Dim oTool As ActiveBar2LibraryCtl.Tool
    
    Set oTool = ActiveBar21.Bands(0).Tools("tFont")
    For i = 0 To Screen.FontCount - 1
        oTool.CBAddItem Screen.Fonts(i)
    Next
    oTool.Text = rtf.Font.name
    
    Set oTool = ActiveBar21.Bands(0).Tools("tFontSize")

    For i = 4 To 40 Step 2
        oTool.CBAddItem i
    Next
    oTool.Text = CInt(rtf.Font.Size)
    
    ActiveBar21.RecalcLayout
End Sub

Private Sub Command1_Click()
    mRTxt = rtf.TextRTF
    OK = True
    Me.Hide
End Sub

Private Sub Command2_Click()
    OK = False
    Me.Hide
End Sub

Private Sub Form_Load()
    InitActiveBar
    rtf.TextRTF = mRTxt
    
    AnimateForm Me
    
End Sub

Private Sub rtf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        'If X and Y are not specified the mouse coordinates are used
        ActiveBar21.Bands("Band2").PopupMenu
    End If
End Sub
