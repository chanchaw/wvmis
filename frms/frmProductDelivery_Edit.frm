VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmProductDelivery_Edit 
   Caption         =   "成品发货查询"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13575
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
   ScaleHeight     =   7410
   ScaleWidth      =   13575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7410
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13575
      _LayoutVersion  =   1
      _ExtentX        =   23945
      _ExtentY        =   13070
      _DataPath       =   ""
      Bands           =   "frmProductDelivery_Edit.frx":0000
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5775
         Left            =   300
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   7455
         _cx             =   13150
         _cy             =   10186
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
         BorderWidth     =   2
         ChildSpacing    =   2
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
         GridRows        =   4
         GridCols        =   4
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmProductDelivery_Edit.frx":17B0
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   1185
            Left            =   30
            ScaleHeight     =   1185
            ScaleWidth      =   7395
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   30
            Width           =   7395
            Begin XtremeSuiteControls.FlatEdit FlatEdit1 
               Height          =   375
               Left            =   1800
               TabIndex        =   3
               Top             =   420
               Width           =   1695
               _Version        =   1048578
               _ExtentX        =   2990
               _ExtentY        =   661
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   255
               Left            =   720
               TabIndex        =   4
               Top             =   480
               Width           =   735
               _Version        =   1048578
               _ExtentX        =   1296
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "订单号:"
            End
         End
         Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
            Height          =   4500
            Left            =   30
            TabIndex        =   5
            Top             =   1245
            Width           =   7395
            _ExtentX        =   13044
            _ExtentY        =   7938
            _LayoutType     =   0
            _RowHeight      =   25
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   953
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColSelect=   0   'False
            Splits(0).DividerColor=   15790320
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=17"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            ColumnFooters   =   -1  'True
            DefColWidth     =   0
            HeadLines       =   1.5
            FootLines       =   1.5
            MultipleLines   =   0
            CellTipsWidth   =   0
            MultiSelect     =   2
            DeadAreaBackColor=   15790320
            RowDividerColor =   15790320
            RowSubDividerColor=   15790320
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(8)   =   ":id=1,.fontname=宋体"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgpicMode=2,.bgbmp=1,.bold=0"
            _StyleDefs(11)  =   ":id=2,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(12)  =   ":id=2,.fontname=宋体"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgpicMode=2,.bgbmp=2,.bold=0"
            _StyleDefs(14)  =   ":id=3,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(15)  =   ":id=3,.fontname=宋体"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.alignment=2,.valignment=2"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
            _StyleDefs(65)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIyanIyanIyanIyanIyanIya"
            _StyleDefs(66)  =   "bmp(1):id=1,nIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIya"
            _StyleDefs(67)  =   "bmp(2):id=1,nIyanIyanAAAAJSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSm"
            _StyleDefs(68)  =   "bmp(3):id=1,pZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpQAAAJyurZyurZyurZyurZyurZyurZyu"
            _StyleDefs(69)  =   "bmp(4):id=1,rZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyu"
            _StyleDefs(70)  =   "bmp(5):id=1,rZyurQAAAKW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2"
            _StyleDefs(71)  =   "bmp(6):id=1,taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2tQAAAK2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(72)  =   "bmp(7):id=1,va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(73)  =   "bmp(8):id=1,vQAAAK2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(74)  =   "bmp(9):id=1,va2+va2+va2+va2+va2+va2+va2+va2+va2+vQAAALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
            _StyleDefs(75)  =   "bmp(10):id=1,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAA"
            _StyleDefs(76)  =   "bmp(11):id=1,ALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
            _StyleDefs(77)  =   "bmp(12):id=1,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(78)  =   "bmp(13):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3P"
            _StyleDefs(79)  =   "bmp(14):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(80)  =   "bmp(15):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(81)  =   "bmp(16):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAM7X1s7X"
            _StyleDefs(82)  =   "bmp(17):id=1,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
            _StyleDefs(83)  =   "bmp(18):id=1,1s7X1s7X1s7X1s7X1s7X1s7X1gAAAM7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
            _StyleDefs(84)  =   "bmp(19):id=1,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1gAAANbj59bj59bj"
            _StyleDefs(85)  =   "bmp(20):id=1,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
            _StyleDefs(86)  =   "bmp(21):id=1,59bj59bj59bj59bj59bj5wAAANbj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
            _StyleDefs(87)  =   "bmp(22):id=1,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj5wAAAN7r797r797r797r"
            _StyleDefs(88)  =   "bmp(23):id=1,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(89)  =   "bmp(24):id=1,797r797r797r797r7wAAAN7r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(90)  =   "bmp(25):id=1,797r797r797r797r797r797r797r797r797r797r797r797r797r7wAAAN7r797r797r797r797r"
            _StyleDefs(91)  =   "bmp(26):id=1,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(92)  =   "bmp(27):id=1,797r797r797r7wAAAA=="
            _StyleDefs(93)  =   "bmp(0):id=2,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIyanIyanIyanIyanIyanIya"
            _StyleDefs(94)  =   "bmp(1):id=2,nIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIya"
            _StyleDefs(95)  =   "bmp(2):id=2,nIyanIyanAAAAJSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSm"
            _StyleDefs(96)  =   "bmp(3):id=2,pZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpQAAAJyurZyurZyurZyurZyurZyurZyu"
            _StyleDefs(97)  =   "bmp(4):id=2,rZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyu"
            _StyleDefs(98)  =   "bmp(5):id=2,rZyurQAAAKW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2"
            _StyleDefs(99)  =   "bmp(6):id=2,taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2tQAAAK2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(100) =   "bmp(7):id=2,va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(101) =   "bmp(8):id=2,vQAAAK2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(102) =   "bmp(9):id=2,va2+va2+va2+va2+va2+va2+va2+va2+va2+vQAAALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
            _StyleDefs(103) =   "bmp(10):id=2,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAA"
            _StyleDefs(104) =   "bmp(11):id=2,ALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
            _StyleDefs(105) =   "bmp(12):id=2,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(106) =   "bmp(13):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3P"
            _StyleDefs(107) =   "bmp(14):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(108) =   "bmp(15):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(109) =   "bmp(16):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAM7X1s7X"
            _StyleDefs(110) =   "bmp(17):id=2,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
            _StyleDefs(111) =   "bmp(18):id=2,1s7X1s7X1s7X1s7X1s7X1s7X1gAAAM7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
            _StyleDefs(112) =   "bmp(19):id=2,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1gAAANbj59bj59bj"
            _StyleDefs(113) =   "bmp(20):id=2,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
            _StyleDefs(114) =   "bmp(21):id=2,59bj59bj59bj59bj59bj5wAAANbj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
            _StyleDefs(115) =   "bmp(22):id=2,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj5wAAAN7r797r797r797r"
            _StyleDefs(116) =   "bmp(23):id=2,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(117) =   "bmp(24):id=2,797r797r797r797r7wAAAN7r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(118) =   "bmp(25):id=2,797r797r797r797r797r797r797r797r797r797r797r797r797r7wAAAN7r797r797r797r797r"
            _StyleDefs(119) =   "bmp(26):id=2,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(120) =   "bmp(27):id=2,797r797r797r7wAAAA=="
         End
      End
   End
End
Attribute VB_Name = "frmProductDelivery_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rss As RecordSet
Public itemid As String
Public itemidb As String

Public B_GoodsNameAlias As String
Public B_SeHao As String
Public B_Producer As String
Public B_Width As String
Public B_weight As String
Public B_PactCode As String
Public B_Client As String
Public B_Clientid As String
Public B_hex As String
Public B_name As String
Public B_sid As String

Public bsaved As Boolean
Public dingdan As String

Public rsdetail As RecordSet

Private Sub setRs()
    Set rsdetail = New RecordSet
    rsdetail.Fields.Append "B_hex", adVarChar, 100
    rsdetail.Fields.Append "B_name", adVarChar, 100

    rsdetail.Fields.Append "B_sid", adVarChar, 100
    rsdetail.Fields.Append "B_itemid", adVarChar, 100
    rsdetail.Fields.Append "B_itemidb", adVarChar, 100
    rsdetail.Fields.Append "B_GoodsNameAlias", adVarChar, 100

    rsdetail.Fields.Append "B_SeHao", adVarChar, 100
    rsdetail.Fields.Append "B_Producer", adVarChar, 100
    rsdetail.Fields.Append "B_width", adVarChar, 100
    rsdetail.Fields.Append "B_weight", adVarChar, 100

    rsdetail.Fields.Append "B_PactCode", adVarChar, 100
    rsdetail.Fields.Append "B_Client", adVarChar, 100
    rsdetail.Fields.Append "B_Clientid", adVarChar, 100
    rsdetail.Fields.Append "B_logo", adVarChar, 100
    rsdetail.Open
End Sub

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    
        Case "查询"
            grid
        Case "退出"
            Unload Me
        Case "选择"
            choose
        Case "同订单号全选"
            chooseall
    End Select
End Sub

Private Sub FlatEdit1_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
           grid
    End Select
End Sub

Private Sub Form_Load()
    InitFrm
    grid
    itemid = ""
    bsaved = False
    
End Sub

Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
End Sub
'查询
Public Sub grid()
    Dim sql As String
    Set rss = New RecordSet
    sql = "exec usp_Productdelivery '" & Trim(FlatEdit1.Text) & "'"
    rss.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid1.DataSource = rss
    TDBGrid1.Columns("B_Hex").FetchStyle = True
    setgrid
End Sub
Private Sub setgrid()
    TDBGrid1.Columns("B_ItemIDB").Caption = "订单号"
    TDBGrid1.Columns("B_GoodsNameAlias").Caption = "色布名称"
    TDBGrid1.Columns("B_SeHao").Caption = "色号"
    TDBGrid1.Columns("B_Producer").Caption = "花型"
    TDBGrid1.Columns("B_width").Caption = "门幅"
    TDBGrid1.Columns("B_weight").Caption = "克重"
    TDBGrid1.Columns("B_PactCode").Caption = "合同号"
    TDBGrid1.Columns("B_ClientName").Caption = "客户名"
    TDBGrid1.Columns("B_Name").Caption = "订单颜色"
    TDBGrid1.Columns("B_departColor").Caption = "染厂下单颜色"
    TDBGrid1.Columns("B_hex").Caption = "颜色示例"
    TDBGrid1.Columns("B_hex").width = 1300
    TDBGrid1.Columns("B_ItemIDB").width = 1100
'    TDBGrid1.Columns("B_GoodsNameAlias").width = 1800
'    TDBGrid1.Columns("gg").width = 1800
'    TDBGrid1.Columns("B_Name").width = 1100
'    TDBGrid1.Columns("B_SeHao").width = 800
'    TDBGrid1.Columns("khname").width = 2400
'    TDBGrid1.Columns("B_Hex").width = 1000
'
    TDBGrid1.Columns("B_SID").Visible = False
    TDBGrid1.Columns("B_SID").Locked = True
    TDBGrid1.Columns("B_SID").AllowSizing = False
    TDBGrid1.Columns("B_ClientID").Visible = False
    TDBGrid1.Columns("B_ClientID").Locked = True
    TDBGrid1.Columns("B_ClientID").AllowSizing = False
        TDBGrid1.Columns("B_ItemID").Visible = False
    TDBGrid1.Columns("B_ItemID").Locked = True
    TDBGrid1.Columns("B_ItemID").AllowSizing = False
     TDBGrid1.Columns("B_logo").Visible = False
    TDBGrid1.Columns("B_logo").Locked = True
    TDBGrid1.Columns("B_logo").AllowSizing = False

    TDBGrid1.HoldFields
    TDBGrid1.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub choose()
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    setRs
    itemid = rss!B_itemid
    itemidb = rss!B_ItemIDB
    B_GoodsNameAlias = rss!B_GoodsNameAlias
    B_SeHao = IIf(IsNull(rss!B_SeHao), "", rss!B_SeHao)
    B_Producer = IIf(IsNull(rss!B_Producer), "", rss!B_Producer)
    B_Width = IIf(IsNull(rss!B_Width), "", rss!B_Width)
    B_weight = IIf(IsNull(rss!B_weight), "", rss!B_weight)
    B_PactCode = IIf(IsNull(rss!B_PactCode), "", rss!B_PactCode)
    B_Client = IIf(IsNull(rss!B_ClientName), "", rss!B_ClientName)
    B_Clientid = IIf(IsNull(rss!B_Clientid), "", rss!B_Clientid)
    If IIf(IsNull(rss!B_hex), "", rss!B_hex) <> "" Then
     B_hex = rss!B_hex
    End If
    B_name = IIf(IsNull(rss!B_name), "", rss!B_name)
    B_sid = IIf(IsNull(rss!B_sid), "", rss!B_sid)
    dingdan = FlatEdit1.Text
    bsaved = True
    Me.Hide
End Sub

Private Sub TDBGrid1_DblClick()
    choose
End Sub
Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, _
    ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, _
    ByVal CellStyle As TrueOleDBGrid80.StyleDisp)

   
    On Error Resume Next
    Debug.Print TDBGrid1.Columns("B_Hex").CellValue(Bookmark)
    CellStyle.BackColor = TDBGrid1.Columns("B_Hex").CellValue(Bookmark)
     CellStyle.ForeColor = TDBGrid1.Columns("B_Hex").CellValue(Bookmark)
End Sub

Private Sub chooseall()
    Dim bool As Boolean
    bool = False
     If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    
'    Do While Not rss.EOF
'        If UCase(FlatEdit1.Text) = UCase(rss!B_ItemIDB) Then
'            bool = True
'        End If
'        rss.movenext
'    Loop
'    rss.MoveFirst
'    If bool = False Then
'        MsgBox "没有相同的单号", vbInformation, "提示"
'        Exit Sub
'    End If

    Dim rs As New RecordSet
    Set rs = rss.Clone
    rs.Filter = " B_ItemIDB='" & FlatEdit1.Text & "'"
    If rs.RecordCount <= 0 Then
        MsgBox "没有该订单号！", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
    
    setRs
  
    rss.MoveFirst
    Do While Not rss.EOF
        
        If UCase(FlatEdit1.Text) = UCase(rss!B_ItemIDB) Then
                rsdetail.addnew
                rsdetail!B_itemid = rss!B_itemid
                rsdetail!B_ItemIDB = rss!B_ItemIDB
                rsdetail!B_GoodsNameAlias = rss!B_GoodsNameAlias
                rsdetail!B_SeHao = rss!B_SeHao
                rsdetail!B_Producer = rss!B_Producer
                rsdetail!B_Width = rss!B_Width
                rsdetail!B_weight = rss!B_weight
                rsdetail!B_PactCode = IIf(IsNull(rss!B_PactCode), "", rss!B_PactCode)
                rsdetail!B_Client = IIf(IsNull(rss!B_ClientName), "", rss!B_ClientName)
                rsdetail!B_Clientid = IIf(IsNull(rss!B_Clientid), "", rss!B_Clientid)
                rsdetail!B_hex = IIf(IsNull(rss!B_hex), "", rss!B_hex)
                rsdetail!B_name = IIf(IsNull(rss!B_name), "", rss!B_name)
                rsdetail!B_sid = IIf(IsNull(rss!B_sid), "", rss!B_sid)
                rss.Update
        End If
        rss.movenext
    Loop
    rss.MoveFirst
    bsaved = True
    dingdan = FlatEdit1.Text
    
    Me.Hide
End Sub
Private Sub TDBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            choose
    End Select
End Sub



