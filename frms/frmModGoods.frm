VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmModGoods 
   Caption         =   "商品情况"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11145
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmModGoods.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7305
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10425
      _LayoutVersion  =   1
      _ExtentX        =   18389
      _ExtentY        =   12885
      _DataPath       =   ""
      Bands           =   "frmModGoods.frx":038A
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   1620
         Top             =   5760
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5595
         Left            =   660
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
         Width           =   8355
         _cx             =   14737
         _cy             =   9869
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
         BackColor       =   13557726
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   3
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
         GridRows        =   2
         GridCols        =   2
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmModGoods.frx":7600
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox Picture1 
            Height          =   1740
            Left            =   2325
            ScaleHeight     =   1680
            ScaleWidth      =   5925
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   3810
            Visible         =   0   'False
            Width           =   5985
         End
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   1620
            Top             =   1320
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmModGoods.frx":7651
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmModGoods.frx":7BEB
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmModGoods.frx":8185
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   5505
            Left            =   45
            TabIndex        =   3
            Top             =   45
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   9710
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   529
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            ImageList       =   "ImageList2"
            Appearance      =   0
         End
         Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
            Bindings        =   "frmModGoods.frx":871F
            Height          =   5505
            Left            =   2325
            TabIndex        =   4
            Top             =   45
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   9710
            _LayoutType     =   0
            _RowHeight      =   23
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
            Splits(0).ShowCollapseExpandIcons=   0   'False
            Splits(0).MarqueeStyle=   2
            Splits(0).Size  =   220
            Splits(0).Size.vt=   2
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   3
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=65809"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2752"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=65809"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   0
            BorderStyle     =   0
            DefColWidth     =   0
            HeadLines       =   1.75
            FootLines       =   1.1
            PictureCurrentRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            PictureCurrentRow(0)=   "bHQAAA4DAABCTQ4DAAAAAAAANgAAACgAAAARAAAADgAAAAEAGAAAAAAA2AIAAAAAAAAAAAAAAAAA"
            PictureCurrentRow(1)=   "AAAAAACltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrUA"
            PictureCurrentRow(2)=   "lKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKalAKW2taW2"
            PictureCurrentRow(3)=   "taW2taW2taW2taW2taW2taW2taW2taW2tZSmpaW2taW2taW2taW2taW2taW2tQCtvr2tvr2tvr2t"
            PictureCurrentRow(4)=   "vr2tvr2tvr2tvr2tvr2tvr2tvr0YGBicrq2tvr2tvr2tvr2tvr2tvr0ArcfGrcfGrcfGrcfGrcfG"
            PictureCurrentRow(5)=   "rcfGrcfGrcfGrcfGrcfGAAAAGBgYnK6trcfGrcfGrcfGrcfGALXHxrXHxrXHxrXHxrXHxrXHxrXH"
            PictureCurrentRow(6)=   "xrXHxrXHxrXHxgAAAAAAABgYGKW2tbXHxrXHxrXHxgC9z869z869z869z869z869z869z869z869"
            PictureCurrentRow(7)=   "z869z84AAAAAAAAAAAAYGBitvrW9z869z84Avc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/O"
            PictureCurrentRow(8)=   "AAAAAAAAAAAAAAAAKTAxvc/Ovc/OAMbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1sbX1gAAAAAA"
            PictureCurrentRow(9)=   "AAAAABAQEKW2tcbX1sbX1gDO19bO19bO19bO19bO19bO19bO19bO19bO19bO19YAAAAAAAAQEBC1"
            PictureCurrentRow(10)=   "trXO19bO19bO19YA1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufnAAAAEBAQtb7G1ufn1ufn"
            PictureCurrentRow(11)=   "1ufn1ufnAN7n797n797n797n797n797n797n797n797n797n7xAQEL3Hxt7n797n797n797n797n"
            PictureCurrentRow(12)=   "7wDe7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7++9z87e7+/e7+/e7+/e7+/e7+/e7+8A5+/3"
            PictureCurrentRow(13)=   "5+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/35+/3AA=="
            PictureCurrentRow.vt=   9
            PictureStandardRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            PictureStandardRow(0)=   "bHQAAN4DAABCTd4DAAAAAAAANgAAACgAAAARAAAAEgAAAAEAGAAAAAAAqAMAAAAAAAAAAAAAAAAA"
            PictureStandardRow(1)=   "AAAAAACMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpyMnpwA"
            PictureStandardRow(2)=   "lKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKallKalAJyurZyu"
            PictureStandardRow(3)=   "rZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurZyurQCltrWltrWltrWl"
            PictureStandardRow(4)=   "trWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrWltrUArb69rb69rb69rb69rb69"
            PictureStandardRow(5)=   "rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69rb69AK2+va2+va2+va2+va2+va2+va2+"
            PictureStandardRow(6)=   "va2+va2+va2+va2+va2+va2+va2+va2+va2+va2+vQC1x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8a1"
            PictureStandardRow(7)=   "x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8a1x8YAtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfGtcfG"
            PictureStandardRow(8)=   "tcfGtcfGtcfGtcfGtcfGtcfGtcfGAL3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3Pzr3P"
            PictureStandardRow(9)=   "zr3Pzr3Pzr3Pzr3Pzr3PzgC9z869z869z869z869z869z869z869z869z869z869z869z869z869"
            PictureStandardRow(10)=   "z869z869z869z84Avc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/Ovc/O"
            PictureStandardRow(11)=   "vc/Ovc/OAM7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X1s7X"
            PictureStandardRow(12)=   "1gDO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19bO19YA1ufn"
            PictureStandardRow(13)=   "1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufn1ufnANbn59bn59bn"
            PictureStandardRow(14)=   "59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wDe7+/e7+/e7+/e7+/e"
            PictureStandardRow(15)=   "7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+/e7+8A3u/v3u/v3u/v3u/v3u/v3u/v"
            PictureStandardRow(16)=   "3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/v3u/vAN7v797v797v797v797v797v797v797v"
            PictureStandardRow(17)=   "797v797v797v797v797v797v797v797v797v7wA="
            PictureStandardRow.vt=   9
            MultipleLines   =   0
            CellTipsWidth   =   0
            MultiSelect     =   2
            DeadAreaBackColor=   16252927
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&HF7FFFF&,.bold=0,.fontsize=900"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(8)   =   ":id=1,.fontname=宋体"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bgcolor=&H8000000F&"
            _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000012&,.bgpicMode=2,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(12)  =   ":id=2,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(13)  =   ":id=2,.fontname=宋体"
            _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=31,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=134"
            _StyleDefs(16)  =   ":id=3,.fontname=宋体"
            _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32,.bgcolor=&H31CFFF&"
            _StyleDefs(19)  =   ":id=6,.fgcolor=&H80000008&"
            _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HF7FFFF&"
            _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
            _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
            _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
            _StyleDefs(24)  =   "RecordSelectorStyle:id=163,.parent=2,.namedParent=167"
            _StyleDefs(25)  =   "FilterBarStyle:id=168,.parent=1,.namedParent=172"
            _StyleDefs(26)  =   "Splits(0).Style:id=57,.parent=1,.alignment=2,.valignment=2"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=66,.parent=4"
            _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=58,.parent=2"
            _StyleDefs(29)  =   "Splits(0).FooterStyle:id=59,.parent=3"
            _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=60,.parent=5"
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=62,.parent=6"
            _StyleDefs(32)  =   "Splits(0).EditorStyle:id=61,.parent=7"
            _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=63,.parent=8"
            _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=64,.parent=9"
            _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=65,.parent=10"
            _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=166,.parent=163"
            _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=171,.parent=168"
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=22,.parent=57"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=19,.parent=58"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=20,.parent=59"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=21,.parent=61"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=72,.parent=57"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=69,.parent=58"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=70,.parent=59"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=71,.parent=61"
            _StyleDefs(46)  =   "Named:id=29:Normal"
            _StyleDefs(47)  =   ":id=29,.parent=0,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0"
            _StyleDefs(48)  =   ":id=29,.charset=0"
            _StyleDefs(49)  =   ":id=29,.fontname=Tahoma"
            _StyleDefs(50)  =   "Named:id=30:Heading"
            _StyleDefs(51)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(52)  =   ":id=30,.wraptext=-1,.bgpicMode=1,.bgbmp=1"
            _StyleDefs(53)  =   "Named:id=31:Footing"
            _StyleDefs(54)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(55)  =   ":id=31,.bgpicMode=1,.bgbmp=2"
            _StyleDefs(56)  =   "Named:id=32:Selected"
            _StyleDefs(57)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(58)  =   "Named:id=33:Caption"
            _StyleDefs(59)  =   ":id=33,.parent=30,.alignment=2"
            _StyleDefs(60)  =   "Named:id=34:HighlightRow"
            _StyleDefs(61)  =   ":id=34,.parent=29,.bgcolor=&H31CFFF&,.fgcolor=&H0&"
            _StyleDefs(62)  =   "Named:id=35:EvenRow"
            _StyleDefs(63)  =   ":id=35,.parent=29,.bgcolor=&HFFFF80&"
            _StyleDefs(64)  =   "Named:id=36:OddRow"
            _StyleDefs(65)  =   ":id=36,.parent=29"
            _StyleDefs(66)  =   "Named:id=167:RecordSelector"
            _StyleDefs(67)  =   ":id=167,.parent=30,.bold=0,.fontsize=900,.italic=0,.underline=0,.strikethrough=0"
            _StyleDefs(68)  =   ":id=167,.charset=0"
            _StyleDefs(69)  =   ":id=167,.fontname=宋体"
            _StyleDefs(70)  =   "Named:id=172:FilterBar"
            _StyleDefs(71)  =   ":id=172,.parent=29"
            _StyleDefs(72)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
            _StyleDefs(73)  =   "bmp(1):id=1,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
            _StyleDefs(74)  =   "bmp(2):id=1,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
            _StyleDefs(75)  =   "bmp(3):id=1,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
            _StyleDefs(76)  =   "bmp(4):id=1,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
            _StyleDefs(77)  =   "bmp(5):id=1,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
            _StyleDefs(78)  =   "bmp(6):id=1,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
            _StyleDefs(79)  =   "bmp(7):id=1,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
            _StyleDefs(80)  =   "bmp(8):id=1,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
            _StyleDefs(81)  =   "bmp(9):id=1,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
            _StyleDefs(82)  =   "bmp(10):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
            _StyleDefs(83)  =   "bmp(11):id=1,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
            _StyleDefs(84)  =   "bmp(12):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
            _StyleDefs(85)  =   "bmp(13):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
            _StyleDefs(86)  =   "bmp(14):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
            _StyleDefs(87)  =   "bmp(15):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
            _StyleDefs(88)  =   "bmp(16):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
            _StyleDefs(89)  =   "bmp(17):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
            _StyleDefs(90)  =   "bmp(18):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
            _StyleDefs(91)  =   "bmp(19):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
            _StyleDefs(92)  =   "bmp(20):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
            _StyleDefs(93)  =   "bmp(21):id=1,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
            _StyleDefs(94)  =   "bmp(22):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
            _StyleDefs(95)  =   "bmp(23):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
            _StyleDefs(96)  =   "bmp(24):id=1,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
            _StyleDefs(97)  =   "bmp(25):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
            _StyleDefs(98)  =   "bmp(26):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
            _StyleDefs(99)  =   "bmp(27):id=1,797v797v797v7wAAAA=="
            _StyleDefs(100) =   "bmp(0):id=2,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
            _StyleDefs(101) =   "bmp(1):id=2,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
            _StyleDefs(102) =   "bmp(2):id=2,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
            _StyleDefs(103) =   "bmp(3):id=2,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
            _StyleDefs(104) =   "bmp(4):id=2,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
            _StyleDefs(105) =   "bmp(5):id=2,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
            _StyleDefs(106) =   "bmp(6):id=2,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
            _StyleDefs(107) =   "bmp(7):id=2,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
            _StyleDefs(108) =   "bmp(8):id=2,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
            _StyleDefs(109) =   "bmp(9):id=2,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
            _StyleDefs(110) =   "bmp(10):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
            _StyleDefs(111) =   "bmp(11):id=2,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
            _StyleDefs(112) =   "bmp(12):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
            _StyleDefs(113) =   "bmp(13):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
            _StyleDefs(114) =   "bmp(14):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
            _StyleDefs(115) =   "bmp(15):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
            _StyleDefs(116) =   "bmp(16):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
            _StyleDefs(117) =   "bmp(17):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
            _StyleDefs(118) =   "bmp(18):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
            _StyleDefs(119) =   "bmp(19):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
            _StyleDefs(120) =   "bmp(20):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
            _StyleDefs(121) =   "bmp(21):id=2,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
            _StyleDefs(122) =   "bmp(22):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
            _StyleDefs(123) =   "bmp(23):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
            _StyleDefs(124) =   "bmp(24):id=2,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
            _StyleDefs(125) =   "bmp(25):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
            _StyleDefs(126) =   "bmp(26):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
            _StyleDefs(127) =   "bmp(27):id=2,797v797v797v7wAAAA=="
         End
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   4440
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmModGoods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_SQL As String
Private m_EditFormName As String
Private clsGridShow1 As clsGridShow

Private boolIsShow As Boolean
Private mvarObjectID As String '局部复制
Private sFilter As String
Private iKeyIndex As Integer
Public m_KeyID As Variant

Private m_FieldID As String
Private m_TableName As String

'单据窗体获取客户资料中用的参数
'==============================
Public frmName As String
Public frm1 As Object
'==============================

Dim Col As TrueOleDBGrid80.Column

'=================================
'自动化弹出窗体用到的参数
Public fatherFrm As Object   '单据窗体(传址)
Private mvarfObjectID As String
Private mvarfFieldName As String  '单据窗体上主表的字段名
Private mvarSendIndex As Integer  '被弹出窗体记录集中的即将被传输出去的数据的Index
Private mvarBillOrDetail As Integer '0 为主表  1为明细表
'=================================


Private A_IsTree As Long
Private A_TreeTableName As String
Private A_TreeParentField As String
Private A_TreeChildField As String
Private A_SBillParentField As String

Private A_KeyField As String
Private A_PrimaryTable As String


Public Property Let BillOrDetail(ByVal vData As Integer)
    mvarBillOrDetail = vData
End Property

Public Property Get BillOrDetail() As Integer
    BillOrDetail = mvarBillOrDetail
End Property


Public Property Let fObjectID(ByVal vData As String)
    mvarfObjectID = vData
End Property

Public Property Get fObjectID() As String
    fObjectID = mvarfObjectID
End Property

Public Property Let fFieldName(ByVal vData As String)
    mvarfFieldName = vData
End Property

Public Property Get fFieldName() As String
    fFieldName = mvarfFieldName
End Property


Public Property Let SendIndex(ByVal vData As String)
    mvarSendIndex = vData
End Property

Public Property Get SendIndex() As String
    SendIndex = mvarSendIndex
End Property
'=================================

Public Property Let ObjectID(ByVal vData As String)

'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ObjectID = 5
    mvarObjectID = vData
End Property

Public Property Get ObjectID() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ObjectID
    ObjectID = mvarObjectID
End Property


'Private Sub SelectTo()
'    If Len(frmName) <= 0 Then
'        Exit Sub
'    End If
'
'
'    Select Case frmName
'        Case "商品情况"
'            frm1.A_rs("B_GoodsName") = TDBGrid1.Columns("B_GoodsName").Value
'            frm1.A_rs("B_Specification") = TDBGrid1.Columns("B_Specification").Value
'        Case "库存查询"
'            frm1.Text1(0).Text = TDBGrid1.Columns("B_GoodsName").Value
'        Case "生产工艺单"
'            frm1.Text1(12).Text = TDBGrid1.Columns("B_GoodsName").Value
'
'        Case "130012"
'            frm1.Controls("B_PinMing").Text = TDBGrid1.Columns("B_GoodsName").Value
'        Case Else
'    End Select
'
'
'    Unload Me
'End Sub

Private Sub SelectTo()
    On Error Resume Next
    Dim rs As RecordSet
    Dim strSQL As String
    Dim m_ToString As String
    Dim oSGRow  As SGRow
    
    Dim m_IndexObject As Long
    Dim i As Long
    Dim j As Long
    
    Dim szSelected As String  '传输到主表的时候可以多选，多个元素间使用英文逗号做间隔
    
    
    m_ToString = ""
    If Len(Trim(mvarfFieldName)) <= 0 Then
        Exit Sub
    End If
    
    
    '传输数据到主表
    If mvarBillOrDetail = 0 Then
        i = 0
        i = InStr(1, mvarfFieldName, "(")
        j = InStr(1, mvarfFieldName, ")")
        
        '获取多选的行的中的某列的VALUE
        'szSelected = GetSGGridMulRowsSingleColValue(SGGrid1, Trim(SGGrid1.Columns(mvarSendIndex).Key))
        szSelected = Adodc1.RecordSet.Fields(mvarSendIndex)
        
        If i > 0 Then
            m_IndexObject = Val(Trim(Mid(mvarfFieldName, i + 1, j - i - 1)))
            'fatherFrm.Controls(left(mvarfFieldName, i - 1))(m_IndexObject).Text = Trim(SGGrid1.Rows.Current.Cells(mvarSendIndex).Text)
            fatherFrm.Controls(Left(mvarfFieldName, i - 1))(m_IndexObject).Text = szSelected
        Else
            'fatherFrm.Controls(mvarfFieldName).Text = Trim(SGGrid1.Rows.Current.Cells(mvarSendIndex).Text)
            fatherFrm.Controls(mvarfFieldName).Text = szSelected
            fatherFrm.Adodc1.RecordSet(mvarfFieldName).Value = szSelected
            fatherFrm.Adodc1.RecordSet.Update
        End If
        
    Else
    '传输数据到明细表
        Set rs = New RecordSet
        strSQL = "Select * From G_PopUpDataSendBLDetail Where B_ObjectID='" & mvarfObjectID & "'"
        Debug.Print strSQL
        strSQL = strSQL & " And B_FieldName='" & mvarfFieldName & "'"
        
        Debug.Print strSQL
        rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        
        Do While Not rs.EOF
            m_ToString = ""
'            For Each oSGRow In SGGrid1.Selection.Grid.Rows
'                If SGGrid1.Selection.IsRowSelected(oSGRow.Position) = True Then
'                    m_ToString = m_ToString & oSGRow.Cells(rs!B_fFieldName).Value & ","
'                End If
'            Next
            
            Dim tdbgRow As Variant
            For Each tdbgRow In TDBGrid1.SelBookmarks
                i = rs!B_fFieldName
                m_ToString = m_ToString & TDBGrid1.Columns(i).Value & ","
            Next
            
            m_ToString = Left(m_ToString, Len(m_ToString) - 1)
            
            
            fatherFrm.TDBGrid1.Columns(rs("B_tFieldName")).Value = m_ToString
'            fatherFrm.TDBGrid1.Columns(rs("B_tFieldName")).Text = m_ToString
            fatherFrm.TDBGrid1.Update
            
            
'            fatherFrm.Adodc2.Recordset(rs("B_tFieldName")).Value = m_ToString
'            fatherFrm.Adodc2.Recordset.Update
            
            rs.movenext
        Loop
        
        
        rs.Close
        Set rs = Nothing
        
        fatherFrm.TDBGrid1.PostMsg 81
    End If
    
    
    Unload Me
End Sub


Private Sub SelectTo_2017年1月9日弃用()
    Dim strSQL  As String
    Dim rs As RecordSet
    Dim m_Index As Integer
    
    
    '数据传输到明细表中
    If mvarBillOrDetail = 1 Then
    
        Set rs = New RecordSet
        strSQL = "Select * From G_PopUpDataSendBLDetail Where B_ObjectID='" & mvarfObjectID & "'"
        strSQL = strSQL & " And B_FieldName='" & mvarfFieldName & "'"
        
        Debug.Print strSQL
        rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        Do While Not rs.EOF
            'fatherFrm.TDBGrid1.Columns(mvarfFieldName).Value = TDBGrid1.Columns(mvarSendIndex).Value
            m_Index = Val(Trim(rs("B_fFieldName")))
            fatherFrm.TDBGrid1.Columns(rs("B_tFieldName")).Value = TDBGrid1.Columns(m_Index).Value
            rs.movenext
        Loop
        
        rs.Close
        Set rs = Nothing
        
        
        Unload Me
    
    Else
    '数据传输到主表中
        Set rs = New RecordSet
        strSQL = "Select * From G_PopUpWindowSet Where B_fObjectID='" & mvarfObjectID & "'"
        strSQL = strSQL & " And B_fControlName='" & mvarfFieldName & "'"
        Debug.Print strSQL
        rs.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        
        m_Index = Val(Trim(rs("B_SendIndex")))
        fatherFrm.Controls(mvarfFieldName).Text = TDBGrid1.Columns(m_Index).Value
        
        
        rs.Close
        Set rs = Nothing
        
        Unload Me
        
    End If
End Sub



Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "新增"
            AddNewObject
        Case "编辑"
            EditObject m_KeyID
            
            GetGoods
        Case "删除"
            If MsgBox("是否要删除?", vbExclamation + vbOKCancel + vbDefaultButton2, "删除") = vbOK Then
                DeleteObject
            End If
        Case "导入"
            
        Case "导出"
            ExportToExcelA
        Case "刷新"
            RefreshGrid
            'LoadObject

        Case "过滤"
            FilterForm
        Case "关闭"
            Unload Me
        
        Case "自制"
            SetPruductType Tool.name
        Case "外购"
            SetPruductType Tool.name
        
        Case "成品"
            SetPruductClass Tool.name
        Case "零件"
            SetPruductClass Tool.name
        Case "原料"
            SetPruductClass Tool.name
        

        Case "选择"
            SelectTo
    End Select
    
End Sub



Private Sub SetPruductType(ByVal m_Type As String)
    On Error Resume Next
    Dim strSQL As String
    Dim Row As Variant
    Dim m_Mark
    
    m_Mark = TDBGrid1.bookmark
    For Each Row In TDBGrid1.SelBookmarks
        Adodc1.RecordSet.bookmark = Row
        
        strSQL = "Update G_Goods Set B_PruductType='" & m_Type & "' Where A_KeyField='" & Adodc1.RecordSet(A_KeyField) & "'"
        Gm.cnnTool.cnn.Execute strSQL
    Next Row
    
    Adodc1.RecordSet.requery
    TDBGrid1.bookmark = m_Mark
End Sub

Private Sub SetPruductClass(ByVal m_Type As String)
    On Error Resume Next
    Dim strSQL As String
    Dim m_Mark
    Dim Row As Variant
    
    m_Mark = TDBGrid1.bookmark
    For Each Row In TDBGrid1.SelBookmarks
        Adodc1.RecordSet.bookmark = Row
        
        strSQL = "Update G_Goods Set B_PruductClass='" & m_Type & "' Where A_KeyField='" & Adodc1.RecordSet(A_KeyField) & "'"
        Gm.cnnTool.cnn.Execute strSQL
    Next Row
    Adodc1.RecordSet.requery
    
    TDBGrid1.bookmark = m_Mark
    
End Sub


Private Sub Form_Load()

    ActiveBar21.ClientAreaControl = C1Elastic1
    ActiveBar21.RecalcLayout
    GetObjectParameter
    
    Me.Left = 0
    Me.Top = 0
    
    AnimateForm Me
    
    FillTreeView
    GetGoods
    
End Sub

'新增对象
Private Sub AddNewObject()
    
    '判断新增权限
    If Gm.PI.JudgeNew(Me.ObjectID) = False Then
        Exit Sub
    End If


    On Error Resume Next
    '刷新网格
    Dim o As Object
    
    Set o = GetFormNew(m_EditFormName)
    With o
        Set .AutoFillRs = GetAutoFillRs.Clone
        .AddNewObject ObjectID
        .Show vbModal
    End With
    LoadObject
    Adodc1.RecordSet.movelast
    
    GetGoods
End Sub

'编辑对象
Private Sub EditObject(ByVal m_KeyID As Variant)


'判断是否有修改的权限
    If Gm.PI.JudgeUpdate(Me.ObjectID) = False Then
        Exit Sub
    End If

    On Error Resume Next
    Dim o As Object
    Dim m_Mark
    
    m_Mark = TDBGrid1.bookmark
    Set o = GetFormNew(m_EditFormName)
    With o
        .m_KeyID = m_KeyID
        .EditObject ObjectID
        .Show vbModal
    End With
    LoadObject
    TDBGrid1.bookmark = m_Mark
    
    
    Dim lBM
    lBM = TDBGrid1.bookmark
    GetGoods
    TDBGrid1.bookmark = Val(lBM)

End Sub

Private Sub GetKeyIndex()
    'iKeyIndex
    Dim i As Integer
    For i = 0 To Adodc1.RecordSet.Fields.Count - 1
        If Adodc1.RecordSet.Fields(i).Properties.Item(4).Value = True Then
            iKeyIndex = i
            Exit Sub
        End If
    Next
End Sub

Private Sub DeleteObject()

'判断是否有删除的权限
If Gm.PI.JudgeDelete(Me.ObjectID) = False Then
    Exit Sub
End If


    On Error GoTo IFERR
    Dim sKey As String
    Dim strSQL As String
    sKey = m_KeyID

    GetField
    
    strSQL = "Delete From " & m_TableName & " Where " & m_FieldID & "='" & sKey & "'"
    Gm.cnnTool.cnn.Execute strSQL
    Adodc1.RecordSet.requery
    
    GetGoods
    
    Exit Sub
IFERR:
    Dim szTip As String
    szTip = "存在对应业务数据，不可删除！"
    MsgBox szTip, vbOKOnly + vbInformation, "提示"
End Sub

'取得记录
Public Sub LoadObject()
    '取得分类

End Sub
                             
Public Function BuillTreeView(ByVal nAreaID As String)
    Dim strSQL As String
    Dim rs As New RecordSet
    Dim Nodx As Node
    
    Set rs = New RecordSet
    
    'strSQL = "Select * From G_GoodsType Where B_Parent='" & nAreaID & "'"
    strSQL = "Select * From " & A_TreeTableName & " Where " & A_TreeParentField & "='" & nAreaID & "'"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF Then
        Exit Function
    End If

    Do While Not rs.EOF
        Set Nodx = TreeView1.Nodes.add("F" & Trim(nAreaID), tvwChild, "F" & Trim(rs(A_TreeChildField)), rs(A_TreeChildField), 1, 2)
        Nodx.Expanded = True
    
        Call BuillTreeView(rs.Fields(A_TreeChildField).Value)
        rs.movenext
    Loop
    
    rs.Close
    Set rs = Nothing
End Function

Private Sub FillTreeView()
    '取得记录集
    Dim Nodx As Node
    TreeView1.Nodes.Clear
    Set Nodx = TreeView1.Nodes.add(, tvwFirst, "F0", "全部分类", 3, 3)
    Nodx.Expanded = True
    Nodx.Selected = True
    BuillTreeView "0"
End Sub

Private Sub GetGoods()
    Dim strSQL As String
    Dim Node As Node
    
    Set Node = TreeView1.SelectedItem
    strSQL = m_SQL
    Debug.Print strSQL
    If Node.Text <> "全部分类" Then
        'strSQL = m_SQL & " And " & A_TreeChildField & "='" & Node.Text & "'"
        strSQL = m_SQL & " And " & A_SBillParentField & "='" & Node.Text & "'"
    End If
    Debug.Print strSQL
    
    With Adodc1
        .ConnectionString = Gm.cnnTool.cnnStr
        .CommandType = adCmdText
        .RecordSource = strSQL
        Debug.Print strSQL
        .Refresh
        
        GetKeyIndex
    End With
    SGridShow
    
    
    A_KeyField = Adodc1.RecordSet(0).name    '网格数据的主键字段
    A_PrimaryTable = Adodc1.RecordSet.Fields(0).Properties(1).Value   '网格数据的表名称
End Sub

Private Sub SGridShow()
    Set clsGridShow1 = New clsGridShow
    Adodc1.RecordSet.Filter = sFilter
    
    'If boolIsShow = False Then
        With clsGridShow1
            .ObjectID = mvarObjectID
            .InitClass TDBGrid1, 3
            .ShowGridFormat

        End With
        'boolIsShow = True
    'End If
End Sub

Private Sub FilterForm()
    Dim frm1 As New frmUCTSearch
    With frm1
        Set .rs = Adodc1.RecordSet.Clone
        .ObjectID = ObjectID
        .FieldType = 3
        .Show vbModal
    End With
    If frm1.OK = True Then
        Adodc1.RecordSet.Filter = ""
        sFilter = frm1.strResult
        SGridShow
    End If
    Unload frm1
    Set frm1 = Nothing
End Sub


Private Sub GetField()
    Dim rs As New RecordSet
    Dim strSQL As String
    
    Set rs = New RecordSet
    strSQL = m_SQL & " And 1=0"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    
    rs.AddNew
    
    m_FieldID = rs.Fields(0).name
    m_TableName = rs.Fields(0).Properties(1).Value
    
    rs.CancelUpdate
    rs.Close
    Set rs = Nothing
End Sub


'取得参数
Private Sub GetObjectParameter()
    Dim rs As New RecordSet
    Dim strSQL As String
    
    Set rs = New RecordSet
    strSQL = "Select * From G_BLS Where B_ObjectID='" & ObjectID & "'"
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    m_SQL = rs("B_SQL")
    m_EditFormName = rs("B_EditFormName")
    
    A_TreeTableName = IIf(IsNull(rs("B_TreeTableName")), "", rs("B_TreeTableName"))
    A_TreeParentField = IIf(IsNull(rs("B_TreeParentField")), "", rs("B_TreeParentField"))
    A_TreeChildField = IIf(IsNull(rs("B_TreeChildField")), "", rs("B_TreeChildField"))
    A_SBillParentField = IIf(IsNull(rs("B_SParentField")), "", rs("B_SParentField"))
    
    
    Me.width = rs("B_Width")
    Me.height = rs("B_Height")
    Me.Caption = rs("B_BillName")
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'CacheFrms.DelFrm mvarObjectID
End Sub

Private Sub TDBGrid1_DblClick()
    If Len(fObjectID) > 0 Then
        clsGridShow1.m_PopUp = True
        Call SelectTo
    Else
        EditObject m_KeyID
    End If
End Sub

Private Sub TDBGrid1_HeadClick(ByVal colIndex As Integer)
    On Error Resume Next
    Dim sSort As String
    Dim sField As String

    sField = TDBGrid1.Columns(colIndex).DataField
    sSort = Adodc1.RecordSet.Sort
    If Len(Trim(sSort)) = 0 Or Len(Trim(sSort)) < Len(Trim(sField)) Then
        Adodc1.RecordSet.Sort = TDBGrid1.Columns(colIndex).DataField & " ASC"
    Else
        If Mid(sSort, 1, Len(sField)) = sField Then
            If Mid(sSort, Len(sSort) - 2, 3) = "ASC" Then
                Adodc1.RecordSet.Sort = TDBGrid1.Columns(colIndex).DataField & " DESC"
            Else
                Adodc1.RecordSet.Sort = TDBGrid1.Columns(colIndex).DataField & " ASC"
            End If
        Else
            Adodc1.RecordSet.Sort = TDBGrid1.Columns(colIndex).DataField & " ASC"
        End If
    End If
End Sub

Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ActiveBar21.Bands("Band3").PopupMenu
    End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Adodc1.RecordSet.EOF Then
        m_KeyID = Adodc1.RecordSet(A_KeyField)
    Else
        m_KeyID = ""
    End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    GetGoods
End Sub

'将网格控件中显示的内容导出到EXCEL中
Private Sub ExportToExcelA()
    On Error Resume Next
    Dim Irow, Icol As Integer
    Dim Irowcount, Icolcount As Integer
    Dim Fieldlen() '存字段长度值
    Dim xlApp As excel.Application
    Dim xlBook As excel.Workbook
    Dim xlSheet As excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.add
    Set xlSheet = xlBook.Worksheets(1)

    '记录条数
    Irowcount = TDBGrid1.ApproxCount
    
    Icolcount = 0
    For Icolcount = 0 To TDBGrid1.Columns.Count - 1
        If TDBGrid1.Columns(Icolcount).width > 0 And TDBGrid1.Columns(Icolcount).Visible = True Then
            Icolcount = Icolcount + 1
        End If
    Next

    ReDim Fieldlen(Icolcount)
    xlApp.Visible = True '显示表格
    
    
    '逐行逐列，双重循环导出数据
    'TDBGrid
    TDBGrid1.MoveFirst
    For Irow = 0 To TDBGrid1.ApproxCount - 1
        TDBGrid1.bookmark = Irow
        For Icol = 0 To TDBGrid1.Columns.Count - 1
            If TDBGrid1.Columns(Icol).Visible = True And TDBGrid1.Columns(Icol).width > 0 Then
                xlSheet.Cells(Irow + 1, Icol + 1).NumberFormatLocal = "@"
                xlSheet.Cells(Irow + 1, Icol + 1).Value = TDBGrid1.Columns(Icol).Text
            End If
        Next
    Next
    
    
    xlApp.Visible = True '显示表格
    'xlBook.Save '保存"
    Set xlApp = Nothing '交还控制给Excel
    Exit Sub
IFERR:
    MsgBox "Excel导出时不正确!", vbExclamation, "Excel"
    Exit Sub

End Sub


'将当前选中的树形控件的节点传递给编辑页面并且自动填充
Private Function GetAutoFillRs() As RecordSet
    Dim Node As Node
    Dim szNodeText As String
    Set Node = TreeView1.SelectedItem
    szNodeText = Node.Text
    
    Dim rs As New RecordSet
    rs.Fields.Append "B_Parent", adVarChar, 100
    rs.Open
    
    rs.AddNew
    rs(0).Value = szNodeText
    
    Set GetAutoFillRs = rs.Clone
End Function

Private Sub RefreshGrid()
    Adodc1.RecordSet.Filter = ""
    LoadObject
End Sub

