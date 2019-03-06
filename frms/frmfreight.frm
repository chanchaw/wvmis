VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmfreight 
   Caption         =   "生成运费单"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmfreight.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   10950
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20160
      _LayoutVersion  =   1
      _ExtentX        =   35560
      _ExtentY        =   19315
      _DataPath       =   ""
      Bands           =   "frmfreight.frx":058A
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6495
         Left            =   300
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   9915
         _cx             =   17489
         _cy             =   11456
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
         _GridInfo       =   $"frmfreight.frx":2ED4
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   6435
            Left            =   30
            TabIndex        =   2
            Top             =   30
            Width           =   9855
            _cx             =   17383
            _cy             =   11351
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   0
            MousePointer    =   0
            Version         =   800
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FrontTabColor   =   3263743
            BackTabColor    =   8948341
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "待处理运费|已生成的运费"
            Align           =   0
            CurrTab         =   1
            FirstTab        =   0
            Style           =   6
            Position        =   5
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   -1  'True
            ShowFocusRect   =   -1  'True
            TabsPerPage     =   0
            BorderWidth     =   0
            BoldCurrent     =   -1  'True
            DogEars         =   -1  'True
            MultiRow        =   0   'False
            MultiRowOffset  =   200
            CaptionStyle    =   0
            TabHeight       =   1500
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   37
            Picture(0)      =   "frmfreight.frx":2F58
            Picture(1)      =   "frmfreight.frx":34F2
            Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
               Height          =   6405
               Left            =   -8940
               TabIndex        =   3
               Top             =   15
               Width           =   8325
               _ExtentX        =   14684
               _ExtentY        =   11298
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
               Splits(0).FilterBar=   -1  'True
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
            Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
               Height          =   6405
               Left            =   1515
               TabIndex        =   4
               Top             =   15
               Width           =   8325
               _ExtentX        =   14684
               _ExtentY        =   11298
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
               Splits(0).FilterBar=   -1  'True
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
End
Attribute VB_Name = "frmfreight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rss As RecordSet
Private rss1 As RecordSet
Private szRS As New RecordSet
Private prtRS As New RecordSet  '打印筛选过的记录集
Private szFilterString01 As String
Private szFilterString02 As String
Private str As String

Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
End Sub

Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "生成运费单"
            setFreight
        Case "退出"
            Unload Me
        Case "删除当行"
            del
        Case "删除运费单"
            delall
        Case "刷新"
            ref
        Case "运费付款"
            costpay
        Case "打印"
            prt
        Case "打印筛选"
            prt1
    End Select
End Sub
Private Sub ref()
    rss.requery
    rss1.requery
End Sub



Private Sub Form_Load()
    InitFrm
    Grid

End Sub

Private Sub Grid()
    Grid1
    grid2
End Sub
'显示第一个网格待处理
Private Sub Grid1()
    Dim sql As String
    Set rss = New RecordSet
    sql = "exec usp_Freight '0'"
    
    rss.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid1.DataSource = rss
    SetGrid1
End Sub

'显示第一个网格待处理
Private Sub grid2()
    Dim sql As String
    Set rss1 = New RecordSet
    sql = "exec usp_Freight '1'"
    rss1.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    TDBGrid2.DataSource = rss1
    setgrid2
End Sub

Private Sub SetGrid1()

    TDBGrid1.Columns("B_BillName").Caption = "单据类型"
    TDBGrid1.Columns("B_CodeID").Caption = "单据编号"
    TDBGrid1.Columns("B_PIShu").Caption = "总匹数"
    TDBGrid1.Columns("B_KG").Caption = "总公斤"
    TDBGrid1.Columns("B_qty").Caption = "总米数"
    TDBGrid1.Columns("B_BoxQty").Caption = "总码数"
    TDBGrid1.Columns("B_CC").Caption = "往来单位"
    TDBGrid1.Columns("B_Date").Caption = "单据日期"
    TDBGrid1.Columns("B_drivename").Caption = "驾驶员"
     TDBGrid1.Columns("B_Deliveryaddress").Caption = "送货地址"
      TDBGrid1.Columns("B_cope").Caption = "付费方式"
    TDBGrid1.Columns("B_PIShu").width = 1000
    TDBGrid1.Columns("B_CC").width = 3000
     TDBGrid1.Columns("B_KG").width = 1200
      TDBGrid1.Columns("B_qty").width = 1200
       TDBGrid1.Columns("B_BoxQty").width = 1200
       TDBGrid1.Columns("B_cope").width = 1200
       
    
    TDBGrid1.Columns("B_ID").Visible = False
   TDBGrid1.Columns("B_ID").Locked = True
   TDBGrid1.Columns("B_ID").AllowSizing = False
   TDBGrid1.FetchRowStyle = True
    TDBGrid1.HoldFields
   TDBGrid1.MarqueeStyle = dbgHighlightRow
   sumall
End Sub

Private Sub setgrid2()
    TDBGrid2.Columns("B_ClientName").Caption = "运方"
    TDBGrid2.Columns("B_date").Caption = "制单时间"
    TDBGrid2.Columns("B_freight").Caption = "运费"
    TDBGrid2.Columns("B_pnumber").Caption = "车牌号"
    TDBGrid2.Columns("B_drivename").Caption = "驾驶员"
    TDBGrid2.Columns("B_drivetelephone").Caption = "驾驶员电话"
    TDBGrid2.Columns("B_Freightlogo").Caption = "运费已付"
    TDBGrid2.Columns("B_singleLogo").Caption = "已收回单"
    TDBGrid2.Columns("B_Freightlogo").ValueItems.Presentation = dbgCheckBox
    TDBGrid2.Columns("B_singleLogo").ValueItems.Presentation = dbgCheckBox
    
    TDBGrid2.Columns("B_BillName").Caption = "单据类型"
    TDBGrid2.Columns("B_CodeID").Caption = "单据编号"
    TDBGrid2.Columns("B_PIShu").Caption = "总匹数"
    TDBGrid2.Columns("B_KG").Caption = "总公斤"
    TDBGrid2.Columns("B_qty").Caption = "总米数"
    TDBGrid2.Columns("B_BoxQty").Caption = "总码数"
    TDBGrid2.Columns("B_CC").Caption = "往来单位"
    TDBGrid2.Columns("B_createDate").Caption = "单据日期"
    TDBGrid2.Columns("B_Deliveryaddress").Caption = "送货地址"
    TDBGrid2.Columns("B_cope").Caption = "付费方式"
    
    TDBGrid2.Columns("B_freight").NumberFormat = "0.0"
  
    TDBGrid2.Columns("B_ClientName").width = 1300
    TDBGrid2.Columns("B_date").width = 1000
    TDBGrid2.Columns("B_cope").width = 1200
    TDBGrid2.Columns("B_freight").width = 1000
    TDBGrid2.Columns("B_pnumber").width = 1200
    TDBGrid2.Columns("B_drivename").width = 1000
    TDBGrid2.Columns("B_drivetelephone").width = 1500
    TDBGrid2.Columns("B_BillName").width = 1400
    TDBGrid2.Columns("B_PIShu").width = 1000
    TDBGrid2.Columns("B_KG").width = 1000
    TDBGrid2.Columns("B_qty").width = 1000
    TDBGrid2.Columns("B_BoxQty").width = 1000
    TDBGrid2.Columns("B_CC").width = 2000
    
        TDBGrid2.Columns("B_FreightCard").Visible = False
   TDBGrid2.Columns("B_FreightCard").Locked = True
   TDBGrid2.Columns("B_FreightCard").AllowSizing = False
         TDBGrid2.Columns("B_Card").Visible = False
   TDBGrid2.Columns("B_Card").Locked = True
   TDBGrid2.Columns("B_Card").AllowSizing = False
         TDBGrid2.Columns("B_CardYh").Visible = False
   TDBGrid2.Columns("B_CardYh").Locked = True
   TDBGrid2.Columns("B_CardYh").AllowSizing = False
    
    TDBGrid2.Columns("B_ID").Visible = False
   TDBGrid2.Columns("B_ID").Locked = True
   TDBGrid2.Columns("B_ID").AllowSizing = False
     TDBGrid2.Columns("whiteid").Visible = False
   TDBGrid2.Columns("whiteid").Locked = True
   TDBGrid2.Columns("whiteid").AllowSizing = False
       TDBGrid2.Columns("B_Freightid").Visible = False
   TDBGrid2.Columns("B_Freightid").Locked = True
   TDBGrid2.Columns("B_Freightid").AllowSizing = False
   TDBGrid2.Columns("B_Freightitemid").Visible = False
   TDBGrid2.Columns("B_Freightitemid").Locked = True
   TDBGrid2.Columns("B_Freightitemid").AllowSizing = False
      TDBGrid2.Columns("B_ClientID").Visible = False
   TDBGrid2.Columns("B_ClientID").Locked = True
   TDBGrid2.Columns("B_ClientID").AllowSizing = False
   
'    TDBGridMergeCell TDBGrid2, "B_ClientName"
'    TDBGridMergeCell TDBGrid2, "B_date"
    TDBGridMergeCell TDBGrid2, "B_ID"
        TDBGridMergeCell TDBGrid2, "B_freight"
'    TDBGridMergeCell TDBGrid2, "B_pnumber"
'        TDBGridMergeCell TDBGrid2, "B_drivename"
'    TDBGridMergeCell TDBGrid2, "B_drivetelephone"
   TDBGrid2.FetchRowStyle = True
    TDBGrid2.HoldFields
   TDBGrid2.MarqueeStyle = dbgHighlightRow
   sumall1
End Sub

Private Sub setRs()
    Set prtRS = New RecordSet
    prtRS.Fields.Append "B_ID", adVarChar, 100
    prtRS.Fields.Append "B_ClientName", adVarChar, 100
    prtRS.Fields.Append "B_date", adVarChar, 100
    prtRS.Fields.Append "B_freight", adVarChar, 100
    prtRS.Fields.Append "B_pnumber", adVarChar, 100
    prtRS.Fields.Append "B_drivename", adVarChar, 100
    prtRS.Fields.Append "B_drivetelephone", adVarChar, 100
    prtRS.Fields.Append "B_Freightlogo", adVarChar, 100
    prtRS.Fields.Append "B_singleLogo", adVarChar, 100
    prtRS.Fields.Append "B_BillName", adVarChar, 100
   prtRS.Fields.Append "B_CodeID", adVarChar, 100
    prtRS.Fields.Append "B_PIShu", adVarChar, 100
    prtRS.Fields.Append "B_KG", adVarChar, 100
    prtRS.Fields.Append "B_qty", adVarChar, 100
    prtRS.Fields.Append "B_BoxQty", adVarChar, 100
    prtRS.Fields.Append "B_CC", adVarChar, 100
    prtRS.Fields.Append "B_createDate", adVarChar, 100
    prtRS.Fields.Append "B_Deliveryaddress", adVarChar, 100
    
'    prtRS.Fields.Append "B_DeliveryGoods", adVarChar, 100
'    prtRS.Fields.Append "B_Deliveryaddress", adVarChar, 100
'    prtRS.Fields.Append "B_PactCode", adVarChar, 100
'    prtRS.Fields.Append "B_Client", adVarChar, 100
'    prtRS.Fields.Append "B_Clientid", adVarChar, 100
'    prtRS.Fields.Append "B_Waitfreight", adVarChar, 100
'    prtRS.Fields.Append "B_freight", adVarChar, 100
'    prtRS.Fields.Append "B_Prepaidfreight", adVarChar, 100
'
'    prtRS.Fields.Append "B_memo", adVarChar, 100
'    prtRS.Fields.Append "B_orderitemid", adVarChar, 100
    prtRS.Open
    
End Sub
Private Sub setFreight()
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    
    '运费增删改查的权限
    Dim sqll As String
    Dim SysRS As New RecordSet
    Dim m_UserName As String
    m_UserName = Gm.SysID.SystemUser
    sqll = "SELECT * FROM G_UserPro WHERE B_UserName='" & m_UserName & "'  AND B_Update=1"
   SysRS.Open sqll, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    '没有权限的用户不能进行操作
    If SysRS.RecordCount <= 0 Then
     MsgBox "没有修改权限，不能进行操作", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
    If C1Tab1.CurrTab <> 0 Then
        Exit Sub
    End If
    If TDBGrid1.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim frm1 As New frmfreight_Edit
    frm1.Show vbModal
    If frm1.bool = False Then
        Exit Sub
    End If
    sql = "select * from G_FreightMain where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    sql1 = "select * from G_Freightdetail where 1=1"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.AddNew
    rs!B_Shipment = frm1.Originalsuppliers
    rs!B_Freight = frm1.FlatEdit2.Text
    rs!B_PNumber = frm1.FlatEdit4.Text
    rs!B_drivename = frm1.FlatEdit8.Text
    rs!B_drivetelephone = frm1.FlatEdit7.Text
    rs!B_Freightlogo = frm1.Check1.Value
    rs!B_singleLogo = frm1.Check2.Value
    rs!B_FreightCard = frm1.FlatEdit3.Text
    rs!B_Card = frm1.FlatEdit5.Text
    rs!B_CardYh = frm1.FlatEdit6.Text

    rs.Update
    
    Dim tdbgRow As Variant
    For Each tdbgRow In TDBGrid1.SelBookmarks
        rss.bookmark = tdbgRow
        rs1.AddNew
            rs1!B_id = rs!B_id
            rs1!B_Codeid = rss!B_Codeid
            rs1!B_OrderID = rss!B_id
        rs1.Update
    Next
    
    Unload frm1
    rss.requery
    rss1.requery
    sumall
    sumall1
     Grid1
    grid2
End Sub
        
Private Sub TDBGrid2_DblClick()
    Dim sql As String
    Dim rs As New RecordSet
    Dim a As Long
    If C1Tab1.CurrTab <> 1 Then
        Exit Sub
    End If
    If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If
        
    a = TDBGrid1.bookmark
    Dim frm1 As New frmfreight_Edit
    frm1.Originalsuppliers = IIf(IsNull(rss1!B_Clientid), "", rss1!B_Clientid)
  
    frm1.FlatEdit1.Text = IIf(IsNull(rss1!B_ClientName), "", rss1!B_ClientName)
    frm1.FlatEdit2.Text = rss1!B_Freight
    frm1.FlatEdit4.Text = rss1!B_PNumber
    frm1.FlatEdit8.Text = rss1!B_drivename
    frm1.FlatEdit7.Text = rss1!B_drivetelephone
    frm1.Check1.Value = rss1!B_Freightlogo
     frm1.Check2.Value = IIf(IsNull(rss1!B_singleLogo), 0, rss1!B_singleLogo)
     frm1.FlatEdit3.Text = IIf(IsNull(rss1!B_FreightCard), "", rss1!B_FreightCard)
      frm1.FlatEdit5.Text = IIf(IsNull(rss1!B_Card), "", rss1!B_Card)
      frm1.FlatEdit6.Text = IIf(IsNull(rss1!B_CardYh), "", rss1!B_CardYh)
    frm1.Show vbModal
    If frm1.bool = False Then
        Exit Sub
    End If
    
    sql = "update G_FreightMain set B_Shipment ='" & frm1.Originalsuppliers & "', B_Freight ='" & frm1.FlatEdit2.Text & "',"
    sql = sql & " B_PNumber ='" & frm1.FlatEdit4.Text & "',B_drivename ='" & frm1.FlatEdit8.Text & "',B_drivetelephone ='" & frm1.FlatEdit7.Text & "',"
    sql = sql & " B_FreightLogo ='" & frm1.Check1.Value & "',B_singleLogo ='" & frm1.Check2.Value & "',B_FreightCard ='" & frm1.FlatEdit3.Text & "',B_Card ='" & frm1.FlatEdit5.Text & "',B_CardYh ='" & frm1.FlatEdit6.Text & "' where B_id='" & rss1!B_Freightid & "'"
    Debug.Print sql
    Gm.cnnTool.cnn.Execute sql
    'MsgBox "修改成功", vbInformation, "提示"
    Unload frm1
    rss1.requery
    sumall1
     TDBGrid1.bookmark = a
End Sub
'进行删除行
Private Sub del()
    '没有权限的用户不能进行操作
    Dim sqll As String
    Dim SysRS As New RecordSet
    Dim m_UserName As String
    m_UserName = Gm.SysID.SystemUser
    sqll = "SELECT * FROM G_UserPro WHERE B_UserName='" & m_UserName & "'  AND B_Delete=1"
   SysRS.Open sqll, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If SysRS.RecordCount <= 0 Then
     MsgBox "没有删除权限，不能进行操作", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
    If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet
    sql1 = "select * from G_Freightdetail where B_ID='" & rss1!B_Freightid & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If rs1.RecordCount = 1 Then
        sql2 = "delete from G_Freightmain where B_ID='" & rss1!B_Freightid & "'"
        rs2.Open sql2, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    End If
    sql = "delete from G_Freightdetail where B_itemid='" & rss1!B_Freightitemid & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    rss1.requery
    rss.requery
    sumall1
    sumall
End Sub
'删除整个运费单
Private Sub delall()
    '没有权限的用户不能进行操作
   Dim sqll As String
    Dim SysRS As New RecordSet
    Dim m_UserName As String
    m_UserName = Gm.SysID.SystemUser
    sqll = "SELECT * FROM G_UserPro WHERE B_UserName='" & m_UserName & "'  AND B_Delete=1"
   SysRS.Open sqll, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If SysRS.RecordCount <= 0 Then
     MsgBox "没有删除权限，不能进行操作", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    
    If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    sql = "delete from G_Freightdetail where B_ID='" & rss1!B_Freightid & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    sql1 = "delete from G_Freightmain where B_ID='" & rss1!B_Freightid & "'"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rss1.requery
    rss.requery
    sumall1
    sumall
End Sub

'小计
Private Sub sumall()
    
     Dim rs As New RecordSet
    Dim a As Long
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim e As String
    Dim f As String
    Dim h As String
 
    If rss.RecordCount <= 0 Then
        Exit Sub
    End If
    a = 0
    b = 0
    c = 0
    d = 0
    
    Set rs = rss.Clone
    rs.MoveFirst
    Do While Not rs.EOF
        a = a + IIf(IsNull(rs!B_PIShu), 0, rs!B_PIShu)
        b = b + IIf(IsNull(rs!B_kg), 0, rs!B_kg)
        c = c + IIf(IsNull(rs!B_qty), 0, rs!B_qty)
        d = d + IIf(IsNull(rs!B_BoxQty), 0, rs!B_BoxQty)
        rs.movenext
    Loop
    e = Format(b, "0.0")
    f = Format(c, "0.0")
    h = Format(d, "0.0")
    TDBGrid1.Columns("B_BillName").FooterText = "合计"
    TDBGrid1.Columns("B_PIShu").FooterText = "" & a & ""
    TDBGrid1.Columns("B_kg").FooterText = "" & e & ""
    TDBGrid1.Columns("B_qty").FooterText = "" & f & ""
    TDBGrid1.Columns("B_BoxQty").FooterText = "" & h & ""
   
End Sub
Private Sub sumall1()
    
     Dim rs As New RecordSet
    Dim a As Long
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim e As String
    Dim f As String
    Dim h As String
    
    Dim j As Double
    Dim k As String
    If rss1.RecordCount <= 0 Then
        Exit Sub
    End If
    a = 0
    b = 0
    c = 0
    d = 0
    j = 0
    Set rs = rss1.Clone
    rs.MoveFirst
    Do While Not rs.EOF
        a = a + IIf(IsNull(rs!B_PIShu), 0, rs!B_PIShu)
        b = b + IIf(IsNull(rs!B_kg), 0, rs!B_kg)
        c = c + IIf(IsNull(rs!B_qty), 0, rs!B_qty)
        d = d + IIf(IsNull(rs!B_BoxQty), 0, rs!B_BoxQty)
        j = j + IIf(IsNull(rs!B_Freight), 0, rs!B_Freight)
        rs.movenext
    Loop
    e = Format(b, "0.0")
    f = Format(c, "0.0")
    h = Format(d, "0.0")
    k = Format(j, "0.0")
    TDBGrid2.Columns("B_ClientName").FooterText = "合计"
    TDBGrid2.Columns("B_PIShu").FooterText = "" & a & ""
    TDBGrid2.Columns("B_kg").FooterText = "" & e & ""
    TDBGrid2.Columns("B_qty").FooterText = "" & f & ""
    TDBGrid2.Columns("B_BoxQty").FooterText = "" & h & ""
   TDBGrid2.Columns("B_freight").FooterText = "" & k & ""
End Sub

Private Sub TDBGrid2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ActiveBar21.Bands("Band2").PopupMenu
    End If
End Sub

Private Sub TDBGridMergeCell(ByRef vTDBGrid As TDBGrid, ByVal vFieldName As String)

'    vTDBGrid.Columns(vFieldName).Merge = dbgMergeFree
    vTDBGrid.Columns(vFieldName).Merge = dbgMergeRestricted
    

End Sub
'字段过滤
Private Sub TDBGrid1_FilterChange()
    ExeTDBGridFilterChange TDBGrid1, rss
    
  sumGrid
End Sub
Private Sub TDBGrid2_FilterChange()
    ExeTDBGridFilterChange TDBGrid2, rss1
    sumGrid1
End Sub

Private Sub ExeTDBGridFilterChange(ByRef vTDBGrid As TDBGrid, ByRef vRs As RecordSet)
    On Error GoTo IFERR
    Dim Col As Integer
    'Set szRS = New RecordSet
    Col = vTDBGrid.Col
    
    vTDBGrid.HoldFields
    str = vTDBGrid.Columns("B_drivename").Text
    
    vRs.Filter = GetTDBGridFilterString(vTDBGrid)
    vTDBGrid.Col = Col
    vTDBGrid.EditActive = True
    Debug.Print vRs.Filter
    Debug.Print vRs.RecordCount
    Set prtRS = vRs.Clone
    Debug.Print prtRS.RecordCount
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "错误发生于对网格控件进行过滤中" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "提示"
End Sub
  
Private Function GetTDBGridFilterString(ByRef vTDBGrid As TDBGrid) As String
    On Error Resume Next
    Dim tmp As String
    Dim n As Integer
    Dim Col As TrueOleDBGrid80.Column
    Dim cols As TrueOleDBGrid80.Columns
       
    Set cols = vTDBGrid.Columns
    
'If C1Tab1.CurrTab = 0 Then
'    For Each Col In cols
'        If Trim(Col.FilterText) <> "" Then
'            n = n + 1
'            If n > 1 Then
'                szFilterString01 = szFilterString01 & " AND "
'            End If
'            Select Case Col.DataWidth
'                Case 23, 6, 11
'                    szFilterString01 = szFilterString01 & Col.DataField & " =" & Col.FilterText
'                Case Else
'                    szFilterString01 = szFilterString01 & Col.DataField & " LIKE '%" & Col.FilterText & "%'"
'            End Select
'        End If
'
'    Next Col
'    GetTDBGridFilterString = szFilterString01
'
'ElseIf C1Tab1.CurrTab = 1 Then
'    For Each Col In cols
'        If Trim(Col.FilterText) <> "" Then
'            n = n + 1
'            If n > 1 Then
'                szFilterString02 = szFilterString02 & " AND "
'            End If
'            Select Case Col.DataWidth
'                Case 23, 6, 11
'                    szFilterString02 = szFilterString02 & Col.DataField & " =" & Col.FilterText
'                Case Else
'                    szFilterString02 = szFilterString02 & Col.DataField & " LIKE '%" & Col.FilterText & "%'"
'            End Select
'        End If
'
'    Next Col
'    GetTDBGridFilterString = szFilterString02
'End If
    For Each Col In cols
        If Trim(Col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
            Select Case Col.DataWidth
                Case 23, 6, 11
                    tmp = tmp & Col.DataField & " =" & Col.FilterText
                Case Else
                    tmp = tmp & Col.DataField & " LIKE '%" & Col.FilterText & "%'"
            End Select
        End If

    Next Col
   GetTDBGridFilterString = tmp
End Function

Private Sub costpay()
    '没有权限的用户不能进行操作
    Dim sqll As String
    Dim SysRS As New RecordSet
    Dim m_UserName As String
    m_UserName = Gm.SysID.SystemUser
    sqll = "SELECT * FROM G_UserPro WHERE B_UserName='" & m_UserName & "'  AND B_Update=1"
   SysRS.Open sqll, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    If SysRS.RecordCount <= 0 Then
     MsgBox "没有修改权限，不能进行操作", vbOKOnly + vbInformation, "提示"
        Exit Sub
    End If
    If C1Tab1.CurrTab <> 1 Then
        Exit Sub
    End If
    If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim frm1 As New frmfreight_copy
    frm1.id = rss1!B_id
    frm1.Show vbModal
    Unload frm1
End Sub
Private Sub prt()
On Error Resume Next
    Dim sql As String
    Dim rs As New RecordSet
    If C1Tab1.CurrTab <> 1 Then
        Exit Sub
    End If
    If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If
    If rss1!B_id = "" Then
        Exit Sub
    End If
    sql = "EXEC usp_FreightPrint '" & rss1!B_id & "','" & Gm.SysID.SystemUserName & "'"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    Dim frm1 As New frmModBLRPreviewOri
    Set frm1.RecordSet = rs.Clone
    frm1.obj = "11S067"
    frm1.ObjectID = "22B090"
    frm1.Show
End Sub

'打印筛选
Private Sub prt1()
    Dim sql As String
    Dim rs As New RecordSet
    If C1Tab1.CurrTab <> 1 Then
        Exit Sub
    End If
    If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If
    
'    setRs
    

'rss1.Filter = " B_drivename = '" & str & "'"
'rss1.Filter = " B_drivename like '%" & str & "%'"


   
'   rss1.Filter = str
'    Do While Not rss1.EOF
'
'      prtRS.AddNew
'        prtRS!B_id = IIf(IsNull(rss1!B_id), "", rss1!B_id)
'        prtRS!B_ClientName = IIf(IsNull(rss1!B_ClientName), "", rss1!B_ClientName)
'        prtRS!B_Date = IIf(IsNull(rss1!B_Date), "", rss1!B_Date)
'        prtRS!B_Freight = IIf(IsNull(rss1!B_Freight), "", rss1!B_Freight)
'        prtRS!B_PNumber = IIf(IsNull(rss1!B_PNumber), "", rss1!B_PNumber)
'        prtRS!B_drivename = IIf(IsNull(rss1!B_drivename), "", rss1!B_drivename)
'        prtRS!B_drivetelephone = IIf(IsNull(rss1!B_drivetelephone), "", rss1!B_drivetelephone)
'        prtRS!B_Freightlogo = IIf(IsNull(rss1!B_Freightlogo), "", rss1!B_Freightlogo)
'        prtRS!B_singleLogo = IIf(IsNull(rss1!B_singleLogo), "", rss1!B_singleLogo)
'        prtRS!B_BillName = IIf(IsNull(rss1!B_BillName), "", rss1!B_BillName)
'        prtRS!B_Codeid = IIf(IsNull(rss1!B_Codeid), "", rss1!B_Codeid)
'        prtRS!B_PIShu = IIf(IsNull(rss1!B_PIShu), "", rss1!B_PIShu)
'        prtRS!B_kg = IIf(IsNull(rss1!B_kg), "", rss1!B_kg)
'        prtRS!B_qty = IIf(IsNull(rss1!B_qty), "", rss1!B_qty)
'        prtRS!B_BoxQty = IIf(IsNull(rss1!B_BoxQty), "", rss1!B_BoxQty)
'        prtRS!B_CC = IIf(IsNull(rss1!B_CC), "", rss1!B_CC)
'        prtRS!B_createDate = IIf(IsNull(rss1!B_createDate), "", rss1!B_createDate)
'        prtRS!B_Deliveryaddress = IIf(IsNull(rss1!B_Deliveryaddress), "", rss1!B_Deliveryaddress)
'
'        prtRS.Update
'        rss1.movenext
'    Loop
''    prtRS.Open
'   Debug.Print prtRS.RecordCount
'   str = buildFilterString
'   rss1.Filter = str
   str = buildFilterString
   Dim cls1 As New clsRecordset
    str = buildFilterString
   Set prtRS = cls1.buildRsWithData(rss1, str).Clone

   
    ' 打开打印预览
    Dim frm1 As New frmModBLRPreviewOri
    Set frm1.RecordSet = prtRS.Clone
    frm1.obj = "11S067"
    frm1.ObjectID = "22B090"
    frm1.m_Judeg = 1
    frm1.Show
End Sub

'用于字段筛选时合计
Private Sub sumGrid()
    
    Dim a As Long
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim e As String
    Dim f As String
    Dim h As String
 
    If rss.RecordCount <= 0 Then
        Exit Sub
    End If
    a = 0
    b = 0
    c = 0
    d = 0
    
    TDBGrid1.MoveFirst
    Do While Not TDBGrid1.EOF
        a = a + IIf(IsNull(Val(TDBGrid1.Columns("B_PIShu").Text)), 0, Val(TDBGrid1.Columns("B_PIShu").Text))
        b = b + IIf(IsNull(Val(TDBGrid1.Columns("B_kg").Text)), 0, Val(TDBGrid1.Columns("B_kg").Text))
        c = c + IIf(IsNull(Val(TDBGrid1.Columns("B_qty").Text)), 0, Val(TDBGrid1.Columns("B_qty").Text))
        d = d + IIf(IsNull(Val(TDBGrid1.Columns("B_BoxQty").Text)), 0, Val(TDBGrid1.Columns("B_BoxQty").Text))
        TDBGrid1.movenext
    Loop
    e = Format(b, "0.0")
    f = Format(c, "0.0")
    h = Format(d, "0.0")
    TDBGrid1.Columns("B_BillName").FooterText = "合计"
    TDBGrid1.Columns("B_PIShu").FooterText = "" & a & ""
    TDBGrid1.Columns("B_kg").FooterText = "" & e & ""
    TDBGrid1.Columns("B_qty").FooterText = "" & f & ""
    TDBGrid1.Columns("B_BoxQty").FooterText = "" & h & ""
End Sub
Private Sub sumGrid1()
     Dim a As Long
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim e As String
    Dim f As String
    Dim h As String
    
    Dim j As Double
    Dim k As String
    If rss1.RecordCount <= 0 Then
        Exit Sub
    End If
    a = 0
    b = 0
    c = 0
    d = 0
    j = 0
    
    TDBGrid2.MoveFirst
    Do While Not TDBGrid2.EOF
        a = a + IIf(IsNull(Val(TDBGrid2.Columns("B_PIShu").Text)), 0, Val(TDBGrid2.Columns("B_PIShu").Text))
        b = b + IIf(IsNull(Val(TDBGrid2.Columns("B_kg").Text)), 0, Val(TDBGrid2.Columns("B_kg").Text))
        c = c + IIf(IsNull(Val(TDBGrid2.Columns("B_qty").Text)), 0, Val(TDBGrid2.Columns("B_qty").Text))
        d = d + IIf(IsNull(Val(TDBGrid2.Columns("B_BoxQty").Text)), 0, Val(TDBGrid2.Columns("B_BoxQty").Text))
        j = j + IIf(IsNull(Val(TDBGrid2.Columns("B_freight").Text)), 0, Val(TDBGrid2.Columns("B_freight").Text))
        TDBGrid2.movenext
    Loop
    e = Format(b, "0.0")
    f = Format(c, "0.0")
    h = Format(d, "0.0")
    k = Format(j, "0.0")
    TDBGrid2.Columns("B_ClientName").FooterText = "合计"
    TDBGrid2.Columns("B_PIShu").FooterText = "" & a & ""
    TDBGrid2.Columns("B_kg").FooterText = "" & e & ""
    TDBGrid2.Columns("B_qty").FooterText = "" & f & ""
    TDBGrid2.Columns("B_BoxQty").FooterText = "" & h & ""
   TDBGrid2.Columns("B_freight").FooterText = "" & k & ""
End Sub

Private Sub sumGrid2()
     Dim a As Long
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim e As String
    Dim f As String
    Dim h As String
    
    Dim j As Double
    Dim k As String
    If szRS.RecordCount <= 0 Then
        Exit Sub
    End If
    a = 0
    b = 0
    c = 0
    d = 0
    j = 0
    
    szRS.MoveFirst
    Do While Not szRS.EOF
        a = a + IIf(IsNull(szRS!B_PIShu), 0, szRS!B_PIShu)
        b = b + IIf(IsNull(szRS!B_kg), 0, szRS!B_kg)
        c = c + IIf(IsNull(szRS!B_qty), 0, szRS!B_qty)
        d = d + IIf(IsNull(szRS!B_BoxQty), 0, szRS!B_BoxQty)
        j = j + IIf(IsNull(szRS!B_Freight), 0, szRS!B_Freight)
        szRS.movenext
    Loop
    e = Format(b, "0.0")
    f = Format(c, "0.0")
    h = Format(d, "0.0")
    k = Format(j, "0.0")
    TDBGrid2.Columns("B_ClientName").FooterText = "合计"
    TDBGrid2.Columns("B_PIShu").FooterText = "" & a & ""
    TDBGrid2.Columns("B_kg").FooterText = "" & e & ""
    TDBGrid2.Columns("B_qty").FooterText = "" & f & ""
    TDBGrid2.Columns("B_BoxQty").FooterText = "" & h & ""
   TDBGrid2.Columns("B_freight").FooterText = "" & k & ""
End Sub


'构建过滤字符串
Private Function buildFilterString() As String
    Dim i As Long
    Dim j As Long
    Dim szReturn As String
    
    szReturn = ""
    j = 0
    For i = 0 To TDBGrid2.Columns.Count - 1
        If Len(Trim(TDBGrid2.Columns(i).FilterText)) > 0 Then
            If j > 0 Then
                szReturn = szReturn & " And "
            End If
        
            szReturn = szReturn & TDBGrid2.Columns(i).DataField & " like '%" & TDBGrid2.Columns(i).FilterText & "%'"
            j = j + 1
        End If
    
    Next
    
    Debug.Print szReturn
    buildFilterString = szReturn
End Function


