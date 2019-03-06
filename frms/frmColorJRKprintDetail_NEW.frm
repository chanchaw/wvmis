VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmColorJRKprintDetail_NEW 
   Caption         =   "发货打印细码单"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColorJRKprintDetail_NEW.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7650
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   8085
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12285
      _LayoutVersion  =   1
      _ExtentX        =   21669
      _ExtentY        =   14261
      _DataPath       =   ""
      Bands           =   "frmColorJRKprintDetail_NEW.frx":058A
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5775
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   11175
         _cx             =   19711
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
         _GridInfo       =   $"frmColorJRKprintDetail_NEW.frx":31A2
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
            Height          =   5280
            Left            =   30
            TabIndex        =   3
            Top             =   465
            Width           =   11115
            _ExtentX        =   19606
            _ExtentY        =   9313
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
            AllowDelete     =   -1  'True
            ColumnFooters   =   -1  'True
            DefColWidth     =   0
            HeadLines       =   1.5
            FootLines       =   1.5
            MultipleLines   =   0
            CellTipsWidth   =   0
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "“公斤”为包装后的毛重，不含“手动空加值”"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   30
            TabIndex        =   2
            Top             =   30
            Width           =   11115
         End
      End
   End
End
Attribute VB_Name = "frmColorJRKprintDetail_NEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_JRKID As Long   '传入G_JRKbill的B_ID
Public m_ZCDate As String  '传入传入G_JRKbill的B_ID的入库时间
Public m_DingDanHao As String
Public m_Color As String
Public m_ItemID As String  '用来接收G_BillDetailColor的主键

Private GridRS As New RecordSet
Private GridRS1 As New RecordSet
Private rsdetail As RecordSet  '离线记录集
Private rsdetail1 As RecordSet
Private m_Judeg1 As String     '用来判断执行的是新增行   还是新增复制行
Private m_Judeg2 As String
Private m As Long    '用来记录空加行的次数
Private n As Long     '用来记录复制行的次数

Private X As Long  '用来判断是否是重新打印



Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "复制行"
            CopyAdd
        Case "保存"
                  save
        Case "新增空加"
            AddNew
        Case "退出"
            DelNumber  '先删除草稿数据
'            ClearAll
            Unload Me
        Case "删除"
            Dele
        Case "全部空加"
            AllKJZ
        Case "打印选择细码单"
            Pri
'       Case "重新打印细码单"
'            AgainPrint
    End Select
End Sub

Private Sub Form_Load()
m_Judeg = 0
X = 0
DelNumber
InitFrm
setRs
openJRKBill
ClearAll
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
End Sub
'填充表中数据
Private Sub openJRKBill()
Set GridRS = New RecordSet
Dim sql As String

sql = "SELECT ROW_NUMBER() OVER (ORDER BY a.B_itemid) AS XUHAO,a.B_itemid,a.B_DataPrint,a.B_GJ,a.B_MS, "
sql = sql & " isnull(a.B_KJZ_SD,0)AS B_KJZ_SD,ISNULL(a.B_KJZ_SD_MS,0)AS B_KJZ_SD_MS, "
sql = sql & " ISNULL(a.B_KJZ_SD_MaS,0)AS B_KJZ_SD_MaS,isnull(a.B_KJZ_Judeg,0)as B_KJZ_Judeg  "
sql = sql & " FROM G_JRKBill a left outer join G_BillDetailColor b on a.B_ID=b.B_ItemID left outer join G_Color f on b.B_Color=f.B_SID"
sql = sql & " WHERE a.B_ZCDate='" & m_ZCDate & "' and isnull(a.B_JudegNumaber,0)=0 and f.B_Name='" & m_Color & "'"
Debug.Print sql
GridRS.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
If GridRS.RecordCount <= 0 Then
    Exit Sub
End If

GridRS.MoveFirst
Do While Not GridRS.EOF
        rsdetail.AddNew

        rsdetail!XuHao = GridRS!XuHao
        rsdetail!B_ItemID = GridRS!B_ItemID
        rsdetail!B_GJ = Format(IIf(IsNull(GridRS!B_GJ), 0, GridRS!B_GJ), "0.0")
        rsdetail!B_MS = Format(IIf(IsNull(GridRS!B_MS), 0, GridRS!B_MS), "0.0")
        rsdetail!B_KJZ_SD = Format(IIf(IsNull(GridRS!B_KJZ_SD), 0, GridRS!B_KJZ_SD), "0.0")
        rsdetail!B_KJZ_SD_MS = Format(IIf(IsNull(GridRS!B_KJZ_SD_MS), 0, GridRS!B_KJZ_SD_MS), "0.0")
        rsdetail!B_KJZ_SD_MaS = Format(IIf(IsNull(GridRS!B_KJZ_SD_MaS), 0, GridRS!B_KJZ_SD_MaS), "0.0")
        rsdetail!B_KJZ_Judeg = IIf(IsNull(GridRS!B_KJZ_Judeg), 0, GridRS!B_KJZ_Judeg)
        rsdetail!B_DataPrint = Format(GridRS!B_DataPrint, "YYYY-MM-DD HH:MM:SS")
        GridRS.movenext

    Loop
    If GridRS.RecordCount > 0 Then
        rsdetail.MoveFirst
    End If
sumall
End Sub

Private Sub setRs()
    Set rsdetail = New RecordSet
    rsdetail.Fields.Append "XUHAO", adVarChar, 100
    rsdetail.Fields.Append "B_itemid", adVarChar, 100
    rsdetail.Fields.Append "B_GJ", adVarChar, 100
    rsdetail.Fields.Append "B_MS", adVarChar, 100
    rsdetail.Fields.Append "B_KJZ_SD", adVarChar, 100
    rsdetail.Fields.Append "B_KJZ_SD_MS", adVarChar, 100
    rsdetail.Fields.Append "B_KJZ_SD_MaS", adVarChar, 100
    rsdetail.Fields.Append "B_KJZ_Judeg", adVarChar, 100
    rsdetail.Fields.Append "B_JudegPrint", adVarChar, 100  '选中打印的标记
    rsdetail.Fields.Append "B_DataPrint", adVarChar, 100   '打印时间
    rsdetail.Open
    
    TDBGrid1.DataSource = rsdetail
    Grid
End Sub
Private Sub Grid()
    TDBGrid1.Columns("XUHAO").Caption = "序号"
    TDBGrid1.Columns("B_GJ").Caption = "公斤"
    TDBGrid1.Columns("B_MS").Caption = "米数"
    TDBGrid1.Columns("B_KJZ_SD").Caption = "手动空加公斤"
    TDBGrid1.Columns("B_KJZ_SD_MS").Caption = "手动空加米数"
    TDBGrid1.Columns("B_KJZ_SD_MaS").Caption = "手动空加码数"
    TDBGrid1.Columns("B_KJZ_Judeg").Caption = "是否为空加行"
'    TDBGrid1.Columns("B_JudegPrint").Caption = "是否打印细码单"
    
    TDBGrid1.Columns("B_KJZ_Judeg").ValueItems.Presentation = dbgCheckBox
    TDBGrid1.Columns("B_JudegPrint").ValueItems.Presentation = dbgCheckBox
    
    TDBGrid1.Columns("B_KJZ_Judeg").Locked = True   '改列不能被修改
    
    TDBGrid1.Columns("XUHAO").width = 1500
    TDBGrid1.Columns("B_GJ").width = 2000
    TDBGrid1.Columns("B_MS").width = 2000
    TDBGrid1.Columns("B_KJZ_SD").width = 1500
    TDBGrid1.Columns("B_KJZ_SD_MS").width = 1500
    TDBGrid1.Columns("B_KJZ_SD_MaS").width = 1500
    TDBGrid1.Columns("B_KJZ_Judeg").width = 1500
    
   
    TDBGrid1.Columns("B_itemid").Visible = False
    TDBGrid1.Columns("B_itemid").AllowSizing = False
    TDBGrid1.Columns("B_itemid").Locked = True
    TDBGrid1.Columns("B_DataPrint").Visible = False
    TDBGrid1.Columns("B_DataPrint").AllowSizing = False
    TDBGrid1.Columns("B_DataPrint").Locked = True
      TDBGrid1.Columns("B_JudegPrint").Visible = False
    TDBGrid1.Columns("B_JudegPrint").AllowSizing = False
    TDBGrid1.Columns("B_JudegPrint").Locked = True


    '设置网格的列头高度
    TDBGrid1.HeadLines = 2#
     '设置选中行背景颜色
    TDBGrid1.HighlightRowStyle.BackColor = &HC0C0C0
    
    TDBGrid1.Style.Font.Size = 14    '内容
    
    TDBGrid1.MarqueeStyle = dbgHighlightRow
    TDBGrid1.HoldFields
    
    sumall
End Sub


'新增空加行
Private Sub AddNew()
Dim sql As String
Dim sql1 As String
Dim rs As New RecordSet

sql = "exec usp_ColorPrintDetail'" & rsdetail!B_ItemID & "','" & 1 & "'"
Debug.Print sql
Gm.cnnTool.cnn.Execute sql
sql1 = "SELECT * FROM G_JRKBill WHERE B_ID='" & m_JRKID & "'"
rs.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
rs.movelast

m_Judeg1 = 1
Dim m_XuHao As String
Dim m_ItemID As Long
m_XuHao = rsdetail.RecordCount + 1
m_ItemID = rs!B_ItemID
rsdetail.AddNew
rsdetail!XuHao = m_XuHao
rsdetail!B_ItemID = str(m_ItemID)
rsdetail!B_KJZ_SD = 0
rsdetail!B_KJZ_SD_MS = 0
rsdetail!B_KJZ_SD_MaS = 0
rsdetail!B_GJ = 0
rsdetail!B_MS = 0
rsdetail!B_KJZ_Judeg = 1
'openJRKBill
End Sub
'新增复制行
Private Sub CopyAdd()
m_Judeg2 = 2

Dim sql As String
Dim sql1 As String
Dim rs As New RecordSet
sql = "exec usp_ColorPrintDetail'" & rsdetail!B_ItemID & "','" & 0 & "'"
Gm.cnnTool.cnn.Execute sql
sql1 = "SELECT * FROM G_JRKBill WHERE B_ID='" & m_JRKID & "'"
rs.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
rs.movelast
Dim m_XuHao As String
Dim m_ItemID As Long
m_XuHao = rsdetail.RecordCount + 1
m_ItemID = rs!B_ItemID
rsdetail.AddNew
rsdetail!XuHao = m_XuHao
rsdetail!B_ItemID = str(m_ItemID)
rsdetail!B_KJZ_SD = 0
rsdetail!B_KJZ_SD_MS = 0
rsdetail!B_KJZ_SD_MaS = 0
rsdetail!B_GJ = 0
rsdetail!B_MS = 0

End Sub
'删除行
Private Sub Dele()
Dim sql As String
Dim rs As New RecordSet
 If MsgBox("确定要删除选中行吗？", vbInformation + vbYesNo + vbDefaultButton2, "提示") = vbNo Then
                Exit Sub
End If
'sql = "DELETE FROM G_JRKBill WHERE B_ItemID='" & GridRS!B_itemid & "'"
'Gm.cnnTool.cnn.Execute sql
If rsdetail.RecordCount > 0 Then
    If IIf(IsNull(rsdetail!B_ItemID), 0, rsdetail!B_ItemID) > 0 Then
    sql = "DELETE FROM G_JRKBill WHERE B_ItemID='" & rsdetail!B_ItemID & "'"
    Gm.cnnTool.cnn.Execute sql
    rsdetail.delete
    Else
     rsdetail.delete
    End If
Else
    Exit Sub
End If

Dim m_XuHao As Long
m_XuHao = 1

If rsdetail.RecordCount > 0 Then
rsdetail.MoveFirst
    Do While Not rsdetail.EOF
     rsdetail!XuHao = m_XuHao
    
       m_XuHao = m_XuHao + 1
      rsdetail.movenext
    Loop
End If
    
    
'openJRKBill
End Sub

Private Sub save()

If rsdetail.RecordCount <= 0 Then
  Exit Sub
End If

Dim sql As String
Dim sql1 As String

rsdetail.MoveFirst
Do While Not rsdetail.EOF
 sql = "UPDATE G_JRKbill SET B_GJ ='" & rsdetail!B_GJ & "',B_MS ='" & rsdetail!B_MS & "',B_KJZ_SD='" & IIf(IsNull(rsdetail!B_KJZ_SD), 0, rsdetail!B_KJZ_SD) & "',B_KJZ_SD_MS='" & IIf(IsNull(rsdetail!B_KJZ_SD_MS), 0, rsdetail!B_KJZ_SD_MS) & "',B_KJZ_SD_MaS='" & IIf(IsNull(rsdetail!B_KJZ_SD_MaS), 0, rsdetail!B_KJZ_SD_MaS) & "',B_Judeg_XS=0 WHERE B_ItemID='" & rsdetail!B_ItemID & "'"
Gm.cnnTool.cnn.Execute sql
        
rsdetail.movenext
Loop
 MsgBox "保存完毕！", vbOKOnly + vbInformation, "提示"
 
Unload Me
End Sub
'关闭窗体或者打开窗体时先删除未保存的草稿数据     即B_Judeg_XS=1  的数据
Private Sub DelNumber()
Dim sql As String
sql = "delete from G_jrkbill where B_Judeg_XS=1"
Gm.cnnTool.cnn.Execute sql
End Sub

Private Sub AllKJZ()
Dim a As String   '用来存放全部增加的空加重量
Dim c As String
Dim d As String
Dim b As Long  '用来就收判断字符

Dim frm1 As New frmColorJRKprintDetail_Write
frm1.Show vbModal
a = frm1.m_TXT1
c = frm1.m_TXT2
d = frm1.m_TXT3
b = frm1.m_Judeg
Unload frm1
If b = 0 Then
Exit Sub
End If

If rsdetail.RecordCount > 0 Then
rsdetail.MoveFirst
    Do While Not rsdetail.EOF
     rsdetail!B_KJZ_SD = a
     rsdetail!B_KJZ_SD_MS = c
     rsdetail!B_KJZ_SD_MaS = d
    
    rsdetail.movenext
    Loop
End If
  MsgBox "全部空加完成！", vbOKOnly + vbInformation, "提示"
End Sub
Private Sub sumall()
     Dim rs As New RecordSet
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim e As String
    Dim f As String
    Dim k As String
    If rsdetail.RecordCount <= 0 Then
        Exit Sub
    End If
    a = 0
    b = 0
    c = 0
    d = 0
    Set rs = rsdetail.Clone
    rs.MoveFirst
    Do While Not rs.EOF
        a = a + IIf(IsNull(rs!B_GJ), 0, rs!B_GJ)
        b = b + IIf(IsNull(rs!B_MS), 0, rs!B_MS)
        c = c + IIf(IsNull(rs!B_KJZ_SD), 0, rs!B_KJZ_SD)
        'd = d + IIf(IsNull(rs!sum4), 0, rs!sum4)
        rs.movenext
    Loop
    rs.MoveFirst
    e = Format(b, "0.0")
    f = Format(c, "0.0")
    a = Format(a, "0.0")
    TDBGrid1.Columns("XUHAO").FooterText = "合计"
    TDBGrid1.Columns("B_GJ").FooterText = "" & a & ""
    TDBGrid1.Columns("B_MS").FooterText = "" & e & ""
    TDBGrid1.Columns("B_KJZ_SD").FooterText = "" & f & ""
    'TDBGrid1.Columns("sum4").FooterText = "" & k & ""
End Sub

      
'打印
Private Sub Pri()
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    
Dim sql1 As String
Dim m_SJ As String
m_SJ = Format(Now, "YYYY-MM-DD HH:MM:SS")


rsdetail.MoveFirst
Do While Not rsdetail.EOF
 
sql1 = "UPDATE G_jrkbill SET B_BDCItemID='" & m_ItemID & "',B_DataPrint='" & m_SJ & "' WHERE B_ItemID='" & rsdetail!B_ItemID & "'"
Gm.cnnTool.cnn.Execute sql1

rsdetail.movenext
Loop
   
    
    Dim sql As String
    Dim rs As New RecordSet
    sql = "exec P_Report_GetDetailFormal_print_NEW '" & m_SJ & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Debug.Print sql
    Dim frm1 As New frmModBLRPreviewOriColor
    Set frm1.RecordSet = rs.Clone
            
    frm1.ObjectID = "22B072"
    frm1.Show vbModal
    Unload frm1
 
 '打印完直接退出
     DelNumber  '先删除草稿数据
    Unload Me
End Sub
'清空已打印的选择
Private Sub ClearAll()
    If TDBGrid1.ApproxCount <= 0 Then
           Exit Sub
     End If
     
Dim sql As String
rsdetail.MoveFirst
Do While Not rsdetail.EOF
sql = "UPDATE G_jrkbill SET B_JudegPrint=0 WHERE B_ItemID='" & rsdetail!B_ItemID & "'"
Gm.cnnTool.cnn.Execute sql
rsdetail.movenext
Loop
End Sub


'重新打印
Private Sub AgainPrint()
If C1Tab1.CurrTab <> 1 Then
MsgBox "请在汇总表中选择要重新打印的记录！", vbOKOnly + vbInformation, "提示"
        Exit Sub
End If
X = 1

Dim sql As String
sql = "UPDATE G_jrkbill SET B_JudegNumaber=0 WHERE B_DataPrint='" & rsdetail1!B_DataPrint & "'"
Gm.cnnTool.cnn.Execute sql

setRs
openJRKBill

rsdetail.MoveFirst
Do While Not rsdetail.EOF
    If rsdetail!B_DataPrint = rsdetail1!B_DataPrint Then
           rsdetail!B_JudegPrint = 1
     Else
          rsdetail!B_JudegPrint = 0
          
    End If
rsdetail.movenext
Loop

C1Tab1.CurrTab = 0

End Sub


