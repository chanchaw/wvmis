VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmOfficeCategory 
   Caption         =   "办公明细录入"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13005
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   13005
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13005
      _LayoutVersion  =   1
      _ExtentX        =   22939
      _ExtentY        =   13785
      _DataPath       =   ""
      Bands           =   "frmOfficeCategory.frx":0000
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6735
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   14415
         _cx             =   25426
         _cy             =   11880
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
         GridRows        =   4
         GridCols        =   4
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmOfficeCategory.frx":276E
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
            Height          =   6555
            Left            =   90
            TabIndex        =   2
            Top             =   90
            Width           =   14235
            _ExtentX        =   25109
            _ExtentY        =   11562
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
            ColumnFooters   =   -1  'True
            DefColWidth     =   0
            HeadLines       =   1.5
            FootLines       =   1.5
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   16777215
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&"
            _StyleDefs(7)   =   ":id=1,.borderColor=&H80000017&,.bold=0,.fontsize=900,.italic=0,.underline=0"
            _StyleDefs(8)   =   ":id=1,.strikethrough=0,.charset=134"
            _StyleDefs(9)   =   ":id=1,.fontname=宋体"
            _StyleDefs(10)  =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(11)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H80000007&,.bgpicMode=2"
            _StyleDefs(12)  =   ":id=2,.bgbmp=1,.bold=0,.fontsize=900,.italic=0,.underline=0,.strikethrough=0"
            _StyleDefs(13)  =   ":id=2,.charset=134"
            _StyleDefs(14)  =   ":id=2,.fontname=宋体"
            _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H80000007&,.bgpicMode=2"
            _StyleDefs(16)  =   ":id=3,.bgbmp=2,.bold=0,.fontsize=900,.italic=0,.underline=0,.strikethrough=0"
            _StyleDefs(17)  =   ":id=3,.charset=134"
            _StyleDefs(18)  =   ":id=3,.fontname=宋体"
            _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80000002&"
            _StyleDefs(21)  =   ":id=6,.fgcolor=&H80000012&"
            _StyleDefs(22)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000005&"
            _StyleDefs(23)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80FFFF&"
            _StyleDefs(24)  =   ":id=8,.fgcolor=&H80000012&,.borderColor=&H80000017&"
            _StyleDefs(25)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(26)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(27)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(28)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(29)  =   "Splits(0).Style:id=13,.parent=1,.alignment=2,.valignment=2"
            _StyleDefs(30)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(31)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HFFF0E1&"
            _StyleDefs(32)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(33)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(34)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H80000002&"
            _StyleDefs(35)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(36)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(37)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(38)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(39)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(40)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(41)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(42)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(45)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(46)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(47)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(48)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(49)  =   "Named:id=33:Normal"
            _StyleDefs(50)  =   ":id=33,.parent=0"
            _StyleDefs(51)  =   "Named:id=34:Heading"
            _StyleDefs(52)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(53)  =   ":id=34,.wraptext=-1"
            _StyleDefs(54)  =   "Named:id=35:Footing"
            _StyleDefs(55)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   "Named:id=36:Selected"
            _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(58)  =   "Named:id=37:Caption"
            _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(60)  =   "Named:id=38:HighlightRow"
            _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(62)  =   "Named:id=39:EvenRow"
            _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(64)  =   "Named:id=40:OddRow"
            _StyleDefs(65)  =   ":id=40,.parent=33"
            _StyleDefs(66)  =   "Named:id=41:RecordSelector"
            _StyleDefs(67)  =   ":id=41,.parent=34"
            _StyleDefs(68)  =   "Named:id=42:FilterBar"
            _StyleDefs(69)  =   ":id=42,.parent=33"
            _StyleDefs(70)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIyanIyanIyanIyanIyanIya"
            _StyleDefs(71)  =   "bmp(1):id=1,nIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIya"
            _StyleDefs(72)  =   "bmp(2):id=1,nIyanIyanAAAAJSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSm"
            _StyleDefs(73)  =   "bmp(3):id=1,pZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpQAAAJyurZyurZyurZyurZyurZyurZyu"
            _StyleDefs(74)  =   "bmp(4):id=1,rZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyu"
            _StyleDefs(75)  =   "bmp(5):id=1,rZyurQAAAKW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2"
            _StyleDefs(76)  =   "bmp(6):id=1,taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2tQAAAK2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(77)  =   "bmp(7):id=1,va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(78)  =   "bmp(8):id=1,vQAAAK2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(79)  =   "bmp(9):id=1,va2+va2+va2+va2+va2+va2+va2+va2+va2+vQAAALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
            _StyleDefs(80)  =   "bmp(10):id=1,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAA"
            _StyleDefs(81)  =   "bmp(11):id=1,ALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
            _StyleDefs(82)  =   "bmp(12):id=1,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(83)  =   "bmp(13):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3P"
            _StyleDefs(84)  =   "bmp(14):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(85)  =   "bmp(15):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(86)  =   "bmp(16):id=1,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAM7X1s7X"
            _StyleDefs(87)  =   "bmp(17):id=1,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
            _StyleDefs(88)  =   "bmp(18):id=1,1s7X1s7X1s7X1s7X1s7X1s7X1gAAAM7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
            _StyleDefs(89)  =   "bmp(19):id=1,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1gAAANbj59bj59bj"
            _StyleDefs(90)  =   "bmp(20):id=1,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
            _StyleDefs(91)  =   "bmp(21):id=1,59bj59bj59bj59bj59bj5wAAANbj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
            _StyleDefs(92)  =   "bmp(22):id=1,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj5wAAAN7r797r797r797r"
            _StyleDefs(93)  =   "bmp(23):id=1,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(94)  =   "bmp(24):id=1,797r797r797r797r7wAAAN7r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(95)  =   "bmp(25):id=1,797r797r797r797r797r797r797r797r797r797r797r797r797r7wAAAN7r797r797r797r797r"
            _StyleDefs(96)  =   "bmp(26):id=1,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(97)  =   "bmp(27):id=1,797r797r797r7wAAAA=="
            _StyleDefs(98)  =   "bmp(0):id=2,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIyanIyanIyanIyanIyanIya"
            _StyleDefs(99)  =   "bmp(1):id=2,nIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIyanIya"
            _StyleDefs(100) =   "bmp(2):id=2,nIyanIyanAAAAJSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSm"
            _StyleDefs(101) =   "bmp(3):id=2,pZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpZSmpQAAAJyurZyurZyurZyurZyurZyurZyu"
            _StyleDefs(102) =   "bmp(4):id=2,rZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyu"
            _StyleDefs(103) =   "bmp(5):id=2,rZyurQAAAKW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2"
            _StyleDefs(104) =   "bmp(6):id=2,taW2taW2taW2taW2taW2taW2taW2taW2taW2taW2tQAAAK2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(105) =   "bmp(7):id=2,va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(106) =   "bmp(8):id=2,vQAAAK2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+"
            _StyleDefs(107) =   "bmp(9):id=2,va2+va2+va2+va2+va2+va2+va2+va2+va2+vQAAALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
            _StyleDefs(108) =   "bmp(10):id=2,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAA"
            _StyleDefs(109) =   "bmp(11):id=2,ALXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXH"
            _StyleDefs(110) =   "bmp(12):id=2,xrXHxrXHxrXHxrXHxrXHxrXHxrXHxrXHxgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(111) =   "bmp(13):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3P"
            _StyleDefs(112) =   "bmp(14):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(113) =   "bmp(15):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            _StyleDefs(114) =   "bmp(16):id=2,zr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3PzgAAAM7X1s7X"
            _StyleDefs(115) =   "bmp(17):id=2,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
            _StyleDefs(116) =   "bmp(18):id=2,1s7X1s7X1s7X1s7X1s7X1s7X1gAAAM7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
            _StyleDefs(117) =   "bmp(19):id=2,1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1gAAANbj59bj59bj"
            _StyleDefs(118) =   "bmp(20):id=2,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
            _StyleDefs(119) =   "bmp(21):id=2,59bj59bj59bj59bj59bj5wAAANbj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj"
            _StyleDefs(120) =   "bmp(22):id=2,59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj59bj5wAAAN7r797r797r797r"
            _StyleDefs(121) =   "bmp(23):id=2,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(122) =   "bmp(24):id=2,797r797r797r797r7wAAAN7r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(123) =   "bmp(25):id=2,797r797r797r797r797r797r797r797r797r797r797r797r797r7wAAAN7r797r797r797r797r"
            _StyleDefs(124) =   "bmp(26):id=2,797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r797r"
            _StyleDefs(125) =   "bmp(27):id=2,797r797r797r7wAAAA=="
         End
      End
   End
End
Attribute VB_Name = "frmOfficeCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsdetail As New RecordSet
Public m_ItemID As Long

Public m_UserName As String
Public m_SDate As String
Public m_EDate As String
Public m_ItemCategory As String
Public m_YuanGongName As String

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    
        Case "保存"
            save
        Case "新增行"
            rsdetail.AddNew
        
        Case "删除"
                ' dete
        Case "删除行"
            deteHang
        Case "退出"
            Unload Me
            
    End Select
End Sub
'删除行
Private Sub deteHang()

If TDBGrid1.ApproxCount > 0 Then
      rsdetail.delete
      rsdetail.MoveFirst
Else
           Exit Sub
End If

End Sub
Private Sub item()
     With ActiveBar21
     .ClientAreaControl = C1Elastic1
     .RecalcLayout
    End With

setRs

End Sub

Private Sub Form_Load()
item

m_UserName = Gm.SysID.SystemUserName



End Sub

Private Sub setRs()
    Set rsdetail = New RecordSet

    
    rsdetail.Fields.Append "B_ItemID", adVarChar, 100
    rsdetail.Fields.Append "B_ItemCategory", adVarChar, 100

    rsdetail.Fields.Append "B_ItemName", adVarChar, 100
    rsdetail.Fields.Append "B_YuanGongName", adVarChar, 100
    rsdetail.Fields.Append "B_Money", adVarChar, 100
    rsdetail.Fields.Append "B_ItDate", adVarChar, 100
    rsdetail.Fields.Append "B_Mome", adVarChar, 100
    rsdetail.Open
    TDBGrid1.DataSource = rsdetail
    
    setrsDetail
End Sub

Private Sub setrsDetail()
   
    TDBGrid1.Columns("B_ItemCategory").Caption = "项目类别"
    TDBGrid1.Columns("B_ItemName").Caption = "项目名称"
    TDBGrid1.Columns("B_YuanGongName").Caption = "员工姓名"
    TDBGrid1.Columns("B_Money").Caption = "金额"
    TDBGrid1.Columns("B_ItDate").Caption = "项目时间"
    TDBGrid1.Columns("B_Mome").Caption = "备注"
    
    TDBGrid1.Columns("B_ItemCategory").Locked = False
    TDBGrid1.Columns("B_ItemCategory").Button = True
    TDBGrid1.Columns("B_YuanGongName").Locked = False
    TDBGrid1.Columns("B_YuanGongName").Button = True
    TDBGrid1.Columns("B_Money").NumberFormat = "0.00"

    
    TDBGrid1.Columns("B_ItemCategory").width = 2000
    TDBGrid1.Columns("B_ItemName").width = 2000
    TDBGrid1.Columns("B_YuanGongName").width = 1500
    TDBGrid1.Columns("B_Money").width = 1300
    TDBGrid1.Columns("B_ItDate").width = 2000
    TDBGrid1.Columns("B_Mome").width = 2500

    
'    TDBGrid1.Columns("B_CharacterUnit").Style.Alignment = dbgCenter
'    TDBGrid1.Columns("B_CharacterUnit").Style.VerticalAlignment = dbgVertCenter
'    TDBGrid1.Columns("rowindex").Style.Alignment = dbgCenter
    
    TDBGrid1.Columns("B_itemid").AllowSizing = False
    TDBGrid1.Columns("B_itemid").Visible = False
    TDBGrid1.Columns("B_itemid").Locked = True
     
    TDBGrid1.HoldFields
    TDBGrid1.MarqueeStyle = dbgHighlightRow
    
End Sub


Private Sub TDBGrid1_ButtonClick(ByVal colIndex As Integer)
    If TDBGrid1.Columns("B_ItemCategory").colIndex = colIndex Then
        Dim frm2 As New frmOfficeShow
        frm2.Show vbModal
        If Len(frm2.DyeName) > 0 Then
             rsdetail!B_ItemCategory = frm2.DyeName
        End If
       
        Unload frm2
    End If
    
       If TDBGrid1.Columns("B_YuanGongName").colIndex = colIndex Then
        Dim frm1 As New frmOfficeYuanGong
        frm1.Show vbModal
        If Len(frm1.DyeName) > 0 Then
            rsdetail!B_YuanGongName = frm1.DyeName
        End If
      
        Unload frm1
    End If
End Sub
'保存并修改
Private Sub save()
Dim sql As String
Dim sql1 As String
Dim sql2 As String
Dim rs As New RecordSet

Dim rss As New RecordSet
Dim j As Long

If TDBGrid1.ApproxCount <= 0 Then
 Exit Sub
End If
'sql = "SELECT * FROM G_OfficeDetails"
'rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

    Set rss = rsdetail.Clone
    rss.MoveFirst
    j = 1
    Do While Not rss.EOF
        If Trim(rss!B_ItemCategory) = "" Then
           MsgBox "第" & j & "行项目类别不能为空", vbInformation, "提示"
            Exit Sub
        End If
        If Trim(rss!B_ItemName) = "" Then
           MsgBox "第" & j & "行项目名称不能为空", vbInformation, "提示"
            Exit Sub
        End If
        If Trim(rss!B_YuanGongName) = "" Then
           MsgBox "第" & j & "行员工姓名不能为空", vbInformation, "提示"
            Exit Sub
        End If
        If Trim(rss!B_Money) = "" Then
           MsgBox "第" & j & "行金额不能为空", vbInformation, "提示"
            Exit Sub
        End If
        rss.movenext
        j = j + 1
    Loop


Dim m_Date As String
m_Date = Now


    If m_ItemID <= 0 Then
        rsdetail.MoveFirst
            Do While Not rsdetail.EOF
            
            sql2 = "INSERT INTO G_OfficeDetails"
            sql2 = sql2 & " (B_ItemCategory,B_ItemName,B_YuanGongName,B_Money,B_PDate,B_ItDate,B_Mome,B_USerName)"
            sql2 = sql2 & " VALUES"
            sql2 = sql2 & " ('" & rsdetail!B_ItemCategory & "','" & rsdetail!B_ItemName & "','" & rsdetail!B_YuanGongName & "',"
            sql2 = sql2 & " '" & rsdetail!B_Money & "','" & m_Date & "','" & rsdetail!B_ItDate & "',"
            sql2 = sql2 & " '" & rsdetail!B_Mome & "','" & m_UserName & "')"
            Gm.cnnTool.cnn.Execute sql2
'            rs.AddNew
'              rs!B_ItemCategory = rsdetail!B_ItemCategory
'              rs!B_ItemName = rsdetail!B_ItemName
'              rs!B_YuanGongName = rsdetail!B_YuanGongName
'              rs!B_Money = rsdetail!B_Money
'              rs!B_PDate = m_Date
'              rs!B_ItDate = rsdetail!B_ItDate
'              rs!B_Mome = rsdetail!B_Mome
'              rs!B_USerName = m_UserName   '登录名
'            rs.Update
            rsdetail.movenext
            Loop
    Else
        sql1 = "update G_OfficeDetails set B_ItemCategory='" & rsdetail!B_ItemCategory & "',B_ItemName='" & rsdetail!B_ItemName & "',"
        sql1 = sql1 & " B_YuanGongName='" & rsdetail!B_YuanGongName & "',B_Money='" & rsdetail!B_Money & "',"
        sql1 = sql1 & " B_ItDate='" & rsdetail!B_ItDate & "',B_Mome='" & rsdetail!B_Mome & "'"
        sql1 = sql1 & " where B_ItemID='" & m_ItemID & "'"
       Gm.cnnTool.cnn.Execute sql1
    End If
        setRs
End Sub
'删除数据
Private Sub dete()
Dim sql As String

If Len(m_ItemID) > 0 Then
 sql = "delete from G_OfficeDetails where B_ItemID='" & m_ItemID & "'"
Gm.cnnTool.cnn.Execute sql
End If

End Sub


Public Sub OpenGrid()
item

Dim sql As String
Dim rs As New RecordSet

sql = "select B_ItemID,B_ItemCategory,B_ItemName,B_YuanGongName,B_Money,B_ItDate,B_Mome"
sql = sql & " FROM G_OfficeDetails"
sql = sql & "  WHERE B_USerName='" & m_UserName & "' and  CONVERT(varchar(100),B_PDate, 23) BETWEEN '" & m_SDate & "' and '" & m_EDate & "'"
sql = sql & "  AND B_ItemCategory='" & m_ItemCategory & "' AND B_YuanGongName='" & m_YuanGongName & "'"
Debug.Print sql
 rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
If rs.RecordCount <= 0 Then
    Exit Sub
End If

rs.MoveFirst
Do While Not rs.EOF
rsdetail.AddNew

    rsdetail!B_ItemID = rs!B_ItemID
    rsdetail!B_ItemCategory = rs!B_ItemCategory
    rsdetail!B_ItemName = rs!B_ItemName
    rsdetail!B_YuanGongName = rs!B_YuanGongName
    rsdetail!B_Money = rs!B_Money
    rsdetail!B_ItDate = rs!B_ItDate
    rsdetail!B_Mome = rs!B_Mome

rs.movenext
Loop

sumall
End Sub

Public Sub OpenDetails()

item

Dim sql As String
Dim rs As New RecordSet

sql = "select * from G_OfficeDetails where B_ItemID='" & m_ItemID & "'"
rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic

rsdetail.AddNew
    rsdetail!B_ItemID = rs!B_ItemID
    rsdetail!B_ItemCategory = rs!B_ItemCategory
    rsdetail!B_ItemName = rs!B_ItemName
    rsdetail!B_YuanGongName = rs!B_YuanGongName
    rsdetail!B_Money = rs!B_Money
    rsdetail!B_ItDate = rs!B_ItDate
    rsdetail!B_Mome = rs!B_Mome
sumall
End Sub

Private Sub TDBGrid1_BeforeColUpdate(ByVal colIndex As Integer, OldValue As Variant, Cancel As Integer)
     If colIndex = TDBGrid1.Columns("B_Money").colIndex Then
        TDBGrid1.Columns("B_Money").Value = Val(TDBGrid1.Columns("B_Money").Value)
        sumall
    End If
    
End Sub
'合计
Private Sub sumall()
    Dim a As Double
    Dim rs As New RecordSet
    Set rs = rsdetail.Clone
    a = 0
    If rsdetail.RecordCount <= 0 Then
        a = 0
    Else
        rs.MoveFirst
        Do While Not rs.EOF
            a = a + IIf(IsNull(rs!B_Money), 0, rs!B_Money)
            rs.movenext
        Loop
        rs.MoveFirst
    End If
    TDBGrid1.Columns("B_ItemCategory").FooterText = "合计"
    TDBGrid1.Columns("B_Money").FooterText = "" & a & ""
End Sub


