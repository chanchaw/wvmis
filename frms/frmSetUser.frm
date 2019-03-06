VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{0E8071F7-A7DD-47AC-95A9-365FFFE096DF}#1.0#0"; "UCButton.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmSetUser 
   BackColor       =   &H00CEDFDE&
   Caption         =   "用户设置"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12945
   Icon            =   "frmSetUser.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7215
   ScaleWidth      =   12945
   StartUpPosition =   2  '屏幕中心
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12945
      _LayoutVersion  =   1
      _ExtentX        =   22834
      _ExtentY        =   12726
      _DataPath       =   ""
      Bands           =   "frmSetUser.frx":038A
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   330
         Left            =   8700
         Top             =   60
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
         Caption         =   "Adodc4"
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   7500
         Top             =   60
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   780
         Top             =   2340
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
               Picture         =   "frmSetUser.frx":286C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSetUser.frx":2E06
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSetUser.frx":31A0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6195
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   420
         Width           =   12915
         _cx             =   22781
         _cy             =   10927
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
         BorderWidth     =   4
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
         GridRows        =   1
         GridCols        =   2
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmSetUser.frx":353A
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   6075
            Left            =   2250
            TabIndex        =   3
            Top             =   60
            Width           =   10605
            _cx             =   18706
            _cy             =   10716
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
            Appearance      =   0
            MousePointer    =   0
            Version         =   800
            BackColor       =   15465210
            ForeColor       =   -2147483630
            FrontTabColor   =   13557726
            BackTabColor    =   15465210
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "子系统设置|权限设置"
            Align           =   0
            CurrTab         =   1
            FirstTab        =   0
            Style           =   3
            Position        =   1
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   -1  'True
            ShowFocusRect   =   0   'False
            TabsPerPage     =   0
            BorderWidth     =   0
            BoldCurrent     =   -1  'True
            DogEars         =   -1  'True
            MultiRow        =   0   'False
            MultiRowOffset  =   200
            CaptionStyle    =   0
            TabHeight       =   0
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   -1  'True
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   37
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   5775
               Index           =   0
               Left            =   -11190
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   15
               Width           =   10575
               _cx             =   18653
               _cy             =   10186
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
               BorderWidth     =   0
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
               GridRows        =   2
               GridCols        =   3
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmSetUser.frx":357F
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.PictureBox Picture1 
                  BackColor       =   &H00CEDFDE&
                  BorderStyle     =   0  'None
                  Height          =   5775
                  Index           =   0
                  Left            =   4680
                  ScaleHeight     =   5775
                  ScaleWidth      =   1200
                  TabIndex        =   5
                  Top             =   0
                  Width           =   1200
                  Begin MSAdodcLib.Adodc Adodc3 
                     Height          =   330
                     Left            =   120
                     Top             =   2460
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
                     Caption         =   "Adodc3"
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
                  Begin MSAdodcLib.Adodc Adodc2 
                     Height          =   330
                     Left            =   120
                     Top             =   1980
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
                  Begin TA_UCButton.UCButton Command1 
                     Height          =   315
                     Index           =   0
                     Left            =   60
                     TabIndex        =   8
                     Top             =   240
                     Width           =   1035
                     _ExtentX        =   1826
                     _ExtentY        =   556
                     Caption         =   ">"
                  End
                  Begin TA_UCButton.UCButton Command1 
                     Height          =   315
                     Index           =   1
                     Left            =   60
                     TabIndex        =   9
                     Top             =   660
                     Width           =   1035
                     _ExtentX        =   1826
                     _ExtentY        =   556
                     Caption         =   "<"
                  End
                  Begin TA_UCButton.UCButton Command1 
                     Height          =   315
                     Index           =   2
                     Left            =   60
                     TabIndex        =   10
                     Top             =   1080
                     Width           =   1035
                     _ExtentX        =   1826
                     _ExtentY        =   556
                     Caption         =   ">>"
                  End
                  Begin TA_UCButton.UCButton Command1 
                     Height          =   315
                     Index           =   3
                     Left            =   60
                     TabIndex        =   11
                     Top             =   1500
                     Width           =   1035
                     _ExtentX        =   1826
                     _ExtentY        =   556
                     Caption         =   "<<"
                  End
               End
               Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
                  Bindings        =   "frmSetUser.frx":35DC
                  Height          =   5775
                  Left            =   0
                  TabIndex        =   6
                  Top             =   0
                  Width           =   4620
                  _ExtentX        =   8149
                  _ExtentY        =   10186
                  _LayoutType     =   4
                  _RowHeight      =   23
                  _WasPersistedAsPixels=   0
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).Caption=   "子系统"
                  Columns(0).DataField=   "B_SubSystem"
                  Columns(0).DataWidth=   20
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   1
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).ShowCollapseExpandIcons=   0   'False
                  Splits(0).MarqueeStyle=   2
                  Splits(0).Size  =   220
                  Splits(0).Size.vt=   2
                  Splits(0).RecordSelectorWidth=   503
                  Splits(0)._SavedRecordSelectors=   0   'False
                  Splits(0).ScrollBars=   2
                  Splits(0).DividerColor=   12632256
                  Splits(0).SpringMode=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=1"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=65809"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
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
                  Caption         =   "   所有子系统"
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
                  _StyleDefs(42)  =   "Named:id=29:Normal"
                  _StyleDefs(43)  =   ":id=29,.parent=0,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0"
                  _StyleDefs(44)  =   ":id=29,.charset=0"
                  _StyleDefs(45)  =   ":id=29,.fontname=Tahoma"
                  _StyleDefs(46)  =   "Named:id=30:Heading"
                  _StyleDefs(47)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(48)  =   ":id=30,.wraptext=-1,.bgpicMode=1,.bgbmp=1"
                  _StyleDefs(49)  =   "Named:id=31:Footing"
                  _StyleDefs(50)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(51)  =   ":id=31,.bgpicMode=1,.bgbmp=2"
                  _StyleDefs(52)  =   "Named:id=32:Selected"
                  _StyleDefs(53)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(54)  =   "Named:id=33:Caption"
                  _StyleDefs(55)  =   ":id=33,.parent=30,.alignment=0,.bgcolor=&H80000015&,.fgcolor=&H80000008&,.bold=0"
                  _StyleDefs(56)  =   ":id=33,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(57)  =   ":id=33,.fontname=宋体"
                  _StyleDefs(58)  =   "Named:id=34:HighlightRow"
                  _StyleDefs(59)  =   ":id=34,.parent=29,.bgcolor=&H31CFFF&,.fgcolor=&H0&"
                  _StyleDefs(60)  =   "Named:id=35:EvenRow"
                  _StyleDefs(61)  =   ":id=35,.parent=29,.bgcolor=&HFFFF80&"
                  _StyleDefs(62)  =   "Named:id=36:OddRow"
                  _StyleDefs(63)  =   ":id=36,.parent=29"
                  _StyleDefs(64)  =   "Named:id=167:RecordSelector"
                  _StyleDefs(65)  =   ":id=167,.parent=30,.bold=0,.fontsize=900,.italic=0,.underline=0,.strikethrough=0"
                  _StyleDefs(66)  =   ":id=167,.charset=0"
                  _StyleDefs(67)  =   ":id=167,.fontname=宋体"
                  _StyleDefs(68)  =   "Named:id=172:FilterBar"
                  _StyleDefs(69)  =   ":id=172,.parent=29"
                  _StyleDefs(70)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(71)  =   "bmp(1):id=1,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(72)  =   "bmp(2):id=1,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
                  _StyleDefs(73)  =   "bmp(3):id=1,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(74)  =   "bmp(4):id=1,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(75)  =   "bmp(5):id=1,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
                  _StyleDefs(76)  =   "bmp(6):id=1,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
                  _StyleDefs(77)  =   "bmp(7):id=1,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(78)  =   "bmp(8):id=1,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(79)  =   "bmp(9):id=1,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(80)  =   "bmp(10):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
                  _StyleDefs(81)  =   "bmp(11):id=1,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(82)  =   "bmp(12):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(83)  =   "bmp(13):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
                  _StyleDefs(84)  =   "bmp(14):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(85)  =   "bmp(15):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(86)  =   "bmp(16):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
                  _StyleDefs(87)  =   "bmp(17):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(88)  =   "bmp(18):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(89)  =   "bmp(19):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
                  _StyleDefs(90)  =   "bmp(20):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(91)  =   "bmp(21):id=1,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(92)  =   "bmp(22):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
                  _StyleDefs(93)  =   "bmp(23):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(94)  =   "bmp(24):id=1,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(95)  =   "bmp(25):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
                  _StyleDefs(96)  =   "bmp(26):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(97)  =   "bmp(27):id=1,797v797v797v7wAAAA=="
                  _StyleDefs(98)  =   "bmp(0):id=2,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(99)  =   "bmp(1):id=2,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(100) =   "bmp(2):id=2,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
                  _StyleDefs(101) =   "bmp(3):id=2,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(102) =   "bmp(4):id=2,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(103) =   "bmp(5):id=2,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
                  _StyleDefs(104) =   "bmp(6):id=2,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
                  _StyleDefs(105) =   "bmp(7):id=2,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(106) =   "bmp(8):id=2,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(107) =   "bmp(9):id=2,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(108) =   "bmp(10):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
                  _StyleDefs(109) =   "bmp(11):id=2,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(110) =   "bmp(12):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(111) =   "bmp(13):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
                  _StyleDefs(112) =   "bmp(14):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(113) =   "bmp(15):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(114) =   "bmp(16):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
                  _StyleDefs(115) =   "bmp(17):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(116) =   "bmp(18):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(117) =   "bmp(19):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
                  _StyleDefs(118) =   "bmp(20):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(119) =   "bmp(21):id=2,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(120) =   "bmp(22):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
                  _StyleDefs(121) =   "bmp(23):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(122) =   "bmp(24):id=2,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(123) =   "bmp(25):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
                  _StyleDefs(124) =   "bmp(26):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(125) =   "bmp(27):id=2,797v797v797v7wAAAA=="
               End
               Begin TrueOleDBGrid80.TDBGrid TDBGrid3 
                  Bindings        =   "frmSetUser.frx":35F1
                  Height          =   5775
                  Left            =   5940
                  TabIndex        =   7
                  Top             =   0
                  Width           =   4635
                  _ExtentX        =   8176
                  _ExtentY        =   10186
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).Caption=   "子系统"
                  Columns(0).DataField=   "B_SubSystem"
                  Columns(0).DataWidth=   20
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   1
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).ShowCollapseExpandIcons=   0   'False
                  Splits(0).MarqueeStyle=   2
                  Splits(0).Size  =   220
                  Splits(0).Size.vt=   2
                  Splits(0).RecordSelectorWidth=   503
                  Splits(0)._SavedRecordSelectors=   0   'False
                  Splits(0).ScrollBars=   2
                  Splits(0).DividerColor=   12632256
                  Splits(0).SpringMode=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=1"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=65792"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
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
                  HeadLines       =   1.2
                  FootLines       =   1.1
                  Caption         =   "   可使用的子系统"
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
                  _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=宋体"
                  _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bgcolor=&H8000000F&"
                  _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000012&,.bold=0,.fontsize=900,.italic=0,.underline=0"
                  _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
                  _StyleDefs(13)  =   ":id=2,.fontname=宋体"
                  _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=31,.bold=0,.fontsize=900,.italic=0"
                  _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
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
                  _StyleDefs(26)  =   "Splits(0).Style:id=57,.parent=1"
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
                  _StyleDefs(42)  =   "Named:id=29:Normal"
                  _StyleDefs(43)  =   ":id=29,.parent=0,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0"
                  _StyleDefs(44)  =   ":id=29,.charset=0"
                  _StyleDefs(45)  =   ":id=29,.fontname=Tahoma"
                  _StyleDefs(46)  =   "Named:id=30:Heading"
                  _StyleDefs(47)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(48)  =   ":id=30,.wraptext=-1,.bgpicMode=1,.bgbmp=1"
                  _StyleDefs(49)  =   "Named:id=31:Footing"
                  _StyleDefs(50)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(51)  =   ":id=31,.bgpicMode=1,.bgbmp=2"
                  _StyleDefs(52)  =   "Named:id=32:Selected"
                  _StyleDefs(53)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(54)  =   "Named:id=33:Caption"
                  _StyleDefs(55)  =   ":id=33,.parent=30,.alignment=0,.bgcolor=&H80000015&,.fgcolor=&H80000006&,.bold=0"
                  _StyleDefs(56)  =   ":id=33,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(57)  =   ":id=33,.fontname=宋体"
                  _StyleDefs(58)  =   "Named:id=34:HighlightRow"
                  _StyleDefs(59)  =   ":id=34,.parent=29,.bgcolor=&H31CFFF&,.fgcolor=&H0&"
                  _StyleDefs(60)  =   "Named:id=35:EvenRow"
                  _StyleDefs(61)  =   ":id=35,.parent=29,.bgcolor=&HFFFF80&"
                  _StyleDefs(62)  =   "Named:id=36:OddRow"
                  _StyleDefs(63)  =   ":id=36,.parent=29"
                  _StyleDefs(64)  =   "Named:id=167:RecordSelector"
                  _StyleDefs(65)  =   ":id=167,.parent=30,.bold=0,.fontsize=900,.italic=0,.underline=0,.strikethrough=0"
                  _StyleDefs(66)  =   ":id=167,.charset=0"
                  _StyleDefs(67)  =   ":id=167,.fontname=宋体"
                  _StyleDefs(68)  =   "Named:id=172:FilterBar"
                  _StyleDefs(69)  =   ":id=172,.parent=29"
                  _StyleDefs(70)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(71)  =   "bmp(1):id=1,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(72)  =   "bmp(2):id=1,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
                  _StyleDefs(73)  =   "bmp(3):id=1,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(74)  =   "bmp(4):id=1,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(75)  =   "bmp(5):id=1,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
                  _StyleDefs(76)  =   "bmp(6):id=1,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
                  _StyleDefs(77)  =   "bmp(7):id=1,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(78)  =   "bmp(8):id=1,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(79)  =   "bmp(9):id=1,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(80)  =   "bmp(10):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
                  _StyleDefs(81)  =   "bmp(11):id=1,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(82)  =   "bmp(12):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(83)  =   "bmp(13):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
                  _StyleDefs(84)  =   "bmp(14):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(85)  =   "bmp(15):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(86)  =   "bmp(16):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
                  _StyleDefs(87)  =   "bmp(17):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(88)  =   "bmp(18):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(89)  =   "bmp(19):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
                  _StyleDefs(90)  =   "bmp(20):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(91)  =   "bmp(21):id=1,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(92)  =   "bmp(22):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
                  _StyleDefs(93)  =   "bmp(23):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(94)  =   "bmp(24):id=1,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(95)  =   "bmp(25):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
                  _StyleDefs(96)  =   "bmp(26):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(97)  =   "bmp(27):id=1,797v797v797v7wAAAA=="
                  _StyleDefs(98)  =   "bmp(0):id=2,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(99)  =   "bmp(1):id=2,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(100) =   "bmp(2):id=2,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
                  _StyleDefs(101) =   "bmp(3):id=2,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(102) =   "bmp(4):id=2,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(103) =   "bmp(5):id=2,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
                  _StyleDefs(104) =   "bmp(6):id=2,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
                  _StyleDefs(105) =   "bmp(7):id=2,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(106) =   "bmp(8):id=2,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(107) =   "bmp(9):id=2,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(108) =   "bmp(10):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
                  _StyleDefs(109) =   "bmp(11):id=2,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(110) =   "bmp(12):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(111) =   "bmp(13):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
                  _StyleDefs(112) =   "bmp(14):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(113) =   "bmp(15):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(114) =   "bmp(16):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
                  _StyleDefs(115) =   "bmp(17):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(116) =   "bmp(18):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(117) =   "bmp(19):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
                  _StyleDefs(118) =   "bmp(20):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(119) =   "bmp(21):id=2,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(120) =   "bmp(22):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
                  _StyleDefs(121) =   "bmp(23):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(122) =   "bmp(24):id=2,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(123) =   "bmp(25):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
                  _StyleDefs(124) =   "bmp(26):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(125) =   "bmp(27):id=2,797v797v797v7wAAAA=="
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   5775
               Index           =   1
               Left            =   15
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   15
               Width           =   10575
               _cx             =   18653
               _cy             =   10186
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
               BorderWidth     =   0
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
               GridRows        =   2
               GridCols        =   2
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmSetUser.frx":3606
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar22 
                  Height          =   390
                  Left            =   0
                  TabIndex        =   14
                  Top             =   0
                  Width           =   10575
                  _LayoutVersion  =   1
                  _ExtentX        =   18653
                  _ExtentY        =   688
                  _DataPath       =   ""
                  Bands           =   "frmSetUser.frx":3657
               End
               Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
                  Bindings        =   "frmSetUser.frx":5C1D
                  Height          =   5325
                  Left            =   0
                  TabIndex        =   13
                  Top             =   450
                  Width           =   3255
                  _ExtentX        =   5741
                  _ExtentY        =   9393
                  _LayoutType     =   4
                  _RowHeight      =   23
                  _WasPersistedAsPixels=   0
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).Caption=   "权限分类"
                  Columns(0).DataField=   "B_MenuClass"
                  Columns(0).DataWidth=   20
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   1
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).ShowCollapseExpandIcons=   0   'False
                  Splits(0).MarqueeStyle=   2
                  Splits(0).Size  =   220
                  Splits(0).Size.vt=   2
                  Splits(0).RecordSelectorWidth=   503
                  Splits(0)._SavedRecordSelectors=   0   'False
                  Splits(0).ScrollBars=   2
                  Splits(0).DividerColor=   12632256
                  Splits(0).SpringMode=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=1"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2461"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2355"
                  Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
                  Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=66065"
                  Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
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
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.alignment=2,.bgcolor=&H8000000F&"
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
                  _StyleDefs(42)  =   "Named:id=29:Normal"
                  _StyleDefs(43)  =   ":id=29,.parent=0,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0"
                  _StyleDefs(44)  =   ":id=29,.charset=0"
                  _StyleDefs(45)  =   ":id=29,.fontname=Tahoma"
                  _StyleDefs(46)  =   "Named:id=30:Heading"
                  _StyleDefs(47)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(48)  =   ":id=30,.wraptext=-1,.bgpicMode=1,.bgbmp=1"
                  _StyleDefs(49)  =   "Named:id=31:Footing"
                  _StyleDefs(50)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(51)  =   ":id=31,.bgpicMode=1,.bgbmp=2"
                  _StyleDefs(52)  =   "Named:id=32:Selected"
                  _StyleDefs(53)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(54)  =   "Named:id=33:Caption"
                  _StyleDefs(55)  =   ":id=33,.parent=30,.alignment=0,.bgcolor=&H80000015&,.fgcolor=&H80000006&,.bold=0"
                  _StyleDefs(56)  =   ":id=33,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(57)  =   ":id=33,.fontname=宋体"
                  _StyleDefs(58)  =   "Named:id=34:HighlightRow"
                  _StyleDefs(59)  =   ":id=34,.parent=29,.bgcolor=&H31CFFF&,.fgcolor=&H0&"
                  _StyleDefs(60)  =   "Named:id=35:EvenRow"
                  _StyleDefs(61)  =   ":id=35,.parent=29,.bgcolor=&HFFFF80&"
                  _StyleDefs(62)  =   "Named:id=36:OddRow"
                  _StyleDefs(63)  =   ":id=36,.parent=29"
                  _StyleDefs(64)  =   "Named:id=167:RecordSelector"
                  _StyleDefs(65)  =   ":id=167,.parent=30,.bold=0,.fontsize=900,.italic=0,.underline=0,.strikethrough=0"
                  _StyleDefs(66)  =   ":id=167,.charset=0"
                  _StyleDefs(67)  =   ":id=167,.fontname=宋体"
                  _StyleDefs(68)  =   "Named:id=172:FilterBar"
                  _StyleDefs(69)  =   ":id=172,.parent=29"
                  _StyleDefs(70)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(71)  =   "bmp(1):id=1,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(72)  =   "bmp(2):id=1,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
                  _StyleDefs(73)  =   "bmp(3):id=1,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(74)  =   "bmp(4):id=1,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(75)  =   "bmp(5):id=1,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
                  _StyleDefs(76)  =   "bmp(6):id=1,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
                  _StyleDefs(77)  =   "bmp(7):id=1,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(78)  =   "bmp(8):id=1,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(79)  =   "bmp(9):id=1,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(80)  =   "bmp(10):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
                  _StyleDefs(81)  =   "bmp(11):id=1,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(82)  =   "bmp(12):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(83)  =   "bmp(13):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
                  _StyleDefs(84)  =   "bmp(14):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(85)  =   "bmp(15):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(86)  =   "bmp(16):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
                  _StyleDefs(87)  =   "bmp(17):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(88)  =   "bmp(18):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(89)  =   "bmp(19):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
                  _StyleDefs(90)  =   "bmp(20):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(91)  =   "bmp(21):id=1,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(92)  =   "bmp(22):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
                  _StyleDefs(93)  =   "bmp(23):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(94)  =   "bmp(24):id=1,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(95)  =   "bmp(25):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
                  _StyleDefs(96)  =   "bmp(26):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(97)  =   "bmp(27):id=1,797v797v797v7wAAAA=="
                  _StyleDefs(98)  =   "bmp(0):id=2,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(99)  =   "bmp(1):id=2,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(100) =   "bmp(2):id=2,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
                  _StyleDefs(101) =   "bmp(3):id=2,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(102) =   "bmp(4):id=2,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(103) =   "bmp(5):id=2,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
                  _StyleDefs(104) =   "bmp(6):id=2,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
                  _StyleDefs(105) =   "bmp(7):id=2,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(106) =   "bmp(8):id=2,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(107) =   "bmp(9):id=2,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(108) =   "bmp(10):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
                  _StyleDefs(109) =   "bmp(11):id=2,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(110) =   "bmp(12):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(111) =   "bmp(13):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
                  _StyleDefs(112) =   "bmp(14):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(113) =   "bmp(15):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(114) =   "bmp(16):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
                  _StyleDefs(115) =   "bmp(17):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(116) =   "bmp(18):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(117) =   "bmp(19):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
                  _StyleDefs(118) =   "bmp(20):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(119) =   "bmp(21):id=2,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(120) =   "bmp(22):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
                  _StyleDefs(121) =   "bmp(23):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(122) =   "bmp(24):id=2,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(123) =   "bmp(25):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
                  _StyleDefs(124) =   "bmp(26):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(125) =   "bmp(27):id=2,797v797v797v7wAAAA=="
               End
               Begin TrueOleDBGrid80.TDBGrid TDBGrid4 
                  Height          =   5325
                  Left            =   3315
                  TabIndex        =   15
                  Top             =   450
                  Width           =   7260
                  _ExtentX        =   12806
                  _ExtentY        =   9393
                  _LayoutType     =   4
                  _RowHeight      =   23
                  _WasPersistedAsPixels=   0
                  Columns(0)._VlistStyle=   4
                  Columns(0)._MaxComboItems=   5
                  Columns(0).Caption=   "选择"
                  Columns(0).DataField=   "B_Check"
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).Caption=   "编号"
                  Columns(1).DataField=   "B_ObjectID"
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(2)._VlistStyle=   0
                  Columns(2)._MaxComboItems=   5
                  Columns(2).Caption=   "名称"
                  Columns(2).DataField=   "B_MenuItem"
                  Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(3)._VlistStyle=   4
                  Columns(3)._MaxComboItems=   5
                  Columns(3).Caption=   "新增"
                  Columns(3).DataField=   "B_New"
                  Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(4)._VlistStyle=   4
                  Columns(4)._MaxComboItems=   5
                  Columns(4).Caption=   "修改"
                  Columns(4).DataField=   "B_Update"
                  Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(5)._VlistStyle=   4
                  Columns(5)._MaxComboItems=   5
                  Columns(5).Caption=   "删除"
                  Columns(5).DataField=   "B_Delete"
                  Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   6
                  Splits(0)._UserFlags=   0
                  Splits(0).ExtendRightColumn=   -1  'True
                  Splits(0).RecordSelectorWidth=   503
                  Splits(0)._SavedRecordSelectors=   0   'False
                  Splits(0).ScrollBars=   2
                  Splits(0).DividerColor=   13160660
                  Splits(0).FilterBar=   -1  'True
                  Splits(0).SpringMode=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=6"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1138"
                  Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
                  Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(6)=   "Column(1).Width=1931"
                  Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1826"
                  Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=17"
                  Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
                  Splits(0)._ColumnProps(11)=   "Column(2).Width=3175"
                  Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
                  Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3069"
                  Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=17"
                  Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
                  Splits(0)._ColumnProps(16)=   "Column(3).Width=1244"
                  Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
                  Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1138"
                  Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=17"
                  Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
                  Splits(0)._ColumnProps(21)=   "Column(4).Width=1244"
                  Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
                  Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1138"
                  Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=17"
                  Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
                  Splits(0)._ColumnProps(26)=   "Column(5).Width=1244"
                  Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
                  Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1138"
                  Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=17"
                  Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
                  Splits.Count    =   1
                  PrintInfos(0)._StateFlags=   3
                  PrintInfos(0).Name=   "piInternal 0"
                  PrintInfos(0).PageHeaderFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
                  PrintInfos(0).PageFooterFont=   "Size=9,Charset=134,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=宋体"
                  PrintInfos(0).PageHeaderHeight=   0
                  PrintInfos(0).PageFooterHeight=   0
                  PrintInfos.Count=   1
                  Appearance      =   0
                  ColumnFooters   =   -1  'True
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
                  EmptyRows       =   -1  'True
                  CellTipsWidth   =   0
                  MultiSelect     =   2
                  DeadAreaBackColor=   16252927
                  RowDividerColor =   13160660
                  RowSubDividerColor=   13160660
                  DirectionAfterEnter=   1
                  MaxRows         =   250000
                  ViewColumnCaptionWidth=   0
                  ViewColumnWidth =   0
                  _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=134"
                  _StyleDefs(5)   =   ":id=0,.fontname=宋体"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HF7FFFF&,.bold=0,.fontsize=900"
                  _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=134"
                  _StyleDefs(8)   =   ":id=1,.fontname=宋体"
                  _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgpicMode=2,.bold=0,.fontsize=900"
                  _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=134"
                  _StyleDefs(12)  =   ":id=2,.fontname=宋体"
                  _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
                  _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=134"
                  _StyleDefs(15)  =   ":id=3,.fontname=宋体"
                  _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H31CFFF&"
                  _StyleDefs(18)  =   ":id=6,.fgcolor=&H80000008&"
                  _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
                  _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
                  _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
                  _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
                  _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgpicMode=2,.appearance=1"
                  _StyleDefs(24)  =   ":id=11,.bgbmp=3"
                  _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
                  _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1,.alignment=2,.valignment=2"
                  _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                  _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
                  _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
                  _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
                  _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
                  _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
                  _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
                  _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFFFF&"
                  _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
                  _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
                  _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
                  _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
                  _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
                  _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
                  _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
                  _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
                  _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
                  _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
                  _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
                  _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
                  _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
                  _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
                  _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
                  _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
                  _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
                  _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
                  _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
                  _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
                  _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
                  _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
                  _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
                  _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
                  _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
                  _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
                  _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
                  _StyleDefs(62)  =   "Named:id=33:Normal"
                  _StyleDefs(63)  =   ":id=33,.parent=0"
                  _StyleDefs(64)  =   "Named:id=34:Heading"
                  _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(66)  =   ":id=34,.wraptext=-1,.bgpicMode=1,.bgbmp=1"
                  _StyleDefs(67)  =   "Named:id=35:Footing"
                  _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(69)  =   ":id=35,.bgpicMode=1,.bgbmp=2"
                  _StyleDefs(70)  =   "Named:id=36:Selected"
                  _StyleDefs(71)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(72)  =   "Named:id=37:Caption"
                  _StyleDefs(73)  =   ":id=37,.parent=34,.alignment=2"
                  _StyleDefs(74)  =   "Named:id=38:HighlightRow"
                  _StyleDefs(75)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(76)  =   "Named:id=39:EvenRow"
                  _StyleDefs(77)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                  _StyleDefs(78)  =   "Named:id=40:OddRow"
                  _StyleDefs(79)  =   ":id=40,.parent=33"
                  _StyleDefs(80)  =   "Named:id=41:RecordSelector"
                  _StyleDefs(81)  =   ":id=41,.parent=34,.bgcolor=&HCEDFDE&,.bgpicMode=0,.borderColor=&H80000005&"
                  _StyleDefs(82)  =   "Named:id=42:FilterBar"
                  _StyleDefs(83)  =   ":id=42,.parent=33"
                  _StyleDefs(84)  =   "bmp(0):id=1,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(85)  =   "bmp(1):id=1,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(86)  =   "bmp(2):id=1,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
                  _StyleDefs(87)  =   "bmp(3):id=1,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(88)  =   "bmp(4):id=1,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(89)  =   "bmp(5):id=1,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
                  _StyleDefs(90)  =   "bmp(6):id=1,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
                  _StyleDefs(91)  =   "bmp(7):id=1,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(92)  =   "bmp(8):id=1,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(93)  =   "bmp(9):id=1,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(94)  =   "bmp(10):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
                  _StyleDefs(95)  =   "bmp(11):id=1,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(96)  =   "bmp(12):id=1,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(97)  =   "bmp(13):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
                  _StyleDefs(98)  =   "bmp(14):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(99)  =   "bmp(15):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(100) =   "bmp(16):id=1,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
                  _StyleDefs(101) =   "bmp(17):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(102) =   "bmp(18):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(103) =   "bmp(19):id=1,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
                  _StyleDefs(104) =   "bmp(20):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(105) =   "bmp(21):id=1,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(106) =   "bmp(22):id=1,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
                  _StyleDefs(107) =   "bmp(23):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(108) =   "bmp(24):id=1,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(109) =   "bmp(25):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
                  _StyleDefs(110) =   "bmp(26):id=1,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(111) =   "bmp(27):id=1,797v797v797v7wAAAA=="
                  _StyleDefs(112) =   "bmp(0):id=2,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(113) =   "bmp(1):id=2,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(114) =   "bmp(2):id=2,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
                  _StyleDefs(115) =   "bmp(3):id=2,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(116) =   "bmp(4):id=2,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(117) =   "bmp(5):id=2,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
                  _StyleDefs(118) =   "bmp(6):id=2,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
                  _StyleDefs(119) =   "bmp(7):id=2,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(120) =   "bmp(8):id=2,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(121) =   "bmp(9):id=2,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(122) =   "bmp(10):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
                  _StyleDefs(123) =   "bmp(11):id=2,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(124) =   "bmp(12):id=2,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(125) =   "bmp(13):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
                  _StyleDefs(126) =   "bmp(14):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(127) =   "bmp(15):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(128) =   "bmp(16):id=2,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
                  _StyleDefs(129) =   "bmp(17):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(130) =   "bmp(18):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(131) =   "bmp(19):id=2,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
                  _StyleDefs(132) =   "bmp(20):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(133) =   "bmp(21):id=2,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(134) =   "bmp(22):id=2,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
                  _StyleDefs(135) =   "bmp(23):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(136) =   "bmp(24):id=2,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(137) =   "bmp(25):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
                  _StyleDefs(138) =   "bmp(26):id=2,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(139) =   "bmp(27):id=2,797v797v797v7wAAAA=="
                  _StyleDefs(140) =   "bmp(0):id=3,KAAAABsAAAASAAAAAQAYAAAAAADoBQAAAAAAAAAAAAAAAAAAAAAAAIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(141) =   "bmp(1):id=3,nIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIycnIyc"
                  _StyleDefs(142) =   "bmp(2):id=3,nIycnIycnAAAAJSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSl"
                  _StyleDefs(143) =   "bmp(3):id=3,pZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpZSlpQAAAJytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(144) =   "bmp(4):id=3,rZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZytrZyt"
                  _StyleDefs(145) =   "bmp(5):id=3,rZytrQAAAKW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1"
                  _StyleDefs(146) =   "bmp(6):id=3,taW1taW1taW1taW1taW1taW1taW1taW1taW1taW1tQAAAK29va29va29va29va29va29va29va29"
                  _StyleDefs(147) =   "bmp(7):id=3,va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(148) =   "bmp(8):id=3,vQAAAK29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29va29"
                  _StyleDefs(149) =   "bmp(9):id=3,va29va29va29va29va29va29va29va29va29vQAAALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(150) =   "bmp(10):id=3,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAA"
                  _StyleDefs(151) =   "bmp(11):id=3,ALXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXG"
                  _StyleDefs(152) =   "bmp(12):id=3,xrXGxrXGxrXGxrXGxrXGxrXGxrXGxrXGxgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(153) =   "bmp(13):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3O"
                  _StyleDefs(154) =   "bmp(14):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(155) =   "bmp(15):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAL3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3O"
                  _StyleDefs(156) =   "bmp(16):id=3,zr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3Ozr3OzgAAAM7W1s7W"
                  _StyleDefs(157) =   "bmp(17):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(158) =   "bmp(18):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1gAAAM7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W"
                  _StyleDefs(159) =   "bmp(19):id=3,1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1s7W1gAAANbn59bn59bn"
                  _StyleDefs(160) =   "bmp(20):id=3,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(161) =   "bmp(21):id=3,59bn59bn59bn59bn59bn5wAAANbn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn"
                  _StyleDefs(162) =   "bmp(22):id=3,59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn59bn5wAAAN7v797v797v797v"
                  _StyleDefs(163) =   "bmp(23):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(164) =   "bmp(24):id=3,797v797v797v797v7wAAAN7v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(165) =   "bmp(25):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v7wAAAN7v797v797v797v797v"
                  _StyleDefs(166) =   "bmp(26):id=3,797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v797v"
                  _StyleDefs(167) =   "bmp(27):id=3,797v797v797v7wAAAA=="
               End
            End
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   6075
            Left            =   60
            TabIndex        =   2
            Top             =   60
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   10716
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   529
            Style           =   7
            ImageList       =   "ImageList1"
            Appearance      =   1
         End
      End
   End
End
Attribute VB_Name = "frmSetUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_NowUserName As String
Private rs As New RecordSet

Dim Col As TrueOleDBGrid80.Column
Dim cols As TrueOleDBGrid80.Columns

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            AddOneSubSystem
        Case 1
            DeleteOneSubSystem
        Case 2
            AddAllSubSystem
        Case 3
            DeleteAllSubSystem
    End Select
    FillGrid
End Sub

'---增加一项
Private Sub AddOneSubSystem()
    
    With Adodc3.RecordSet
        .AddNew
        .Fields("B_userName") = m_NowUserName
        .Fields("B_subSystem") = Adodc2.RecordSet("B_subSystem")
        .Update
    End With
End Sub

Private Sub DeleteOneSubSystem()
    On Error Resume Next
    TDBGrid3.delete
    TDBGrid3.Update
End Sub

'---增加所有子项
Private Sub AddAllSubSystem()
    Dim strSQL As String
    
    Gm.cnnTool.IniConnection8DM Gm.SysID.DBInfo
    'cnn.InitializeConnection
    strSQL = "Insert Into G_SUserSub (B_userName,B_subSystem) "
    strSQL = strSQL & "Select '" & Trim(m_NowUserName) & "',B_subSystem From G_SubSystem"
    strSQL = strSQL & " Where B_subSystem Not In (Select B_subSystem From G_SUserSub Where B_userName ='" & Trim(m_NowUserName) & "')"
    Gm.cnnTool.cnn.Execute strSQL
End Sub

'---删除所有子项
Private Sub DeleteAllSubSystem()
    Dim strSQL As String
    
    'cnn.InitializeConnection
    Gm.cnnTool.IniConnection8DM Gm.SysID.DBInfo
    strSQL = "Delete G_SUserSub Where B_userName= '" & Trim(m_NowUserName) & "'"
    Gm.cnnTool.cnn.Execute strSQL
End Sub



Private Sub Form_Load()
    ActiveBar21.ClientAreaControl = C1Elastic1
    ActiveBar21.RecalcLayout
    
    AnimateForm Me
    
    FillTreeView
    
    GetMenuClass
End Sub

Private Sub FillTreeView()
    Dim rs As New RecordSet
    Dim strSQL As String
    Dim Nodx As Node
    
    Set rs = New RecordSet
    
    strSQL = "Select B_UserDes,B_Username From G_SystemUser Where B_UserDes<>'管理员'"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenForwardOnly, adLockReadOnly
    
    TreeView1.Nodes.Clear
    Set Nodx = TreeView1.Nodes.add(, tvwFirst, "F", "系统用户", 1, 1)
    Nodx.Expanded = True
    
    Do While Not rs.EOF
        Set Nodx = TreeView1.Nodes.add("F", tvwChild, Trim(rs("B_UserDes")), rs("B_UserName") & "-" & rs("B_UserDes"), 3, 2)
        rs.movenext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub NavigatorNode(ByVal m_Key As String)
    On Error Resume Next
    
    If m_Key = "F" Then
        Exit Sub
    End If
    
    Dim o As Node
    
    For Each o In TreeView1.Nodes
        If Mid(o.Key, 2, Len(o.Key) - 1) = m_Key Then
            o.Selected = True
            Exit Sub
        End If
    Next
End Sub


Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "新建用户"
            AddNewUser
        Case "编辑用户"
            EditUser
        Case "删除用户"
            DeleteUser
       Case "复制权限"
            CopyPermissions
        Case "清除口令"
            ClearPassWord
        Case "关闭"
            Unload Me
    End Select
End Sub

Private Sub EditUser()
    On Error GoTo IFERR
    Dim sUserName As String
    Dim sDepartment As String
    
    Dim frm1 As New frmSetUserOperation
    With frm1
        .m_UserName = m_NowUserName
        .Show vbModal
    End With
    

    Unload frm1
    Set frm1 = Nothing
    
    FillTreeView
    Exit Sub
    
IFERR:
    Exit Sub
End Sub

'新增用户
Private Sub AddNewUser()
    On Error GoTo IFERR
    Dim sUserName As String
    Dim sDepartment As String
    
    Dim frm1 As New frmSetUserOperation
    With frm1
        .Show vbModal
    End With
    

    Unload frm1
    Set frm1 = Nothing
    
    FillTreeView
    Exit Sub
IFERR:

    Exit Sub
End Sub

'删除用户
Private Sub DeleteUser()
    If Len(m_NowUserName) < 1 Then
        Exit Sub
    End If
    Dim strSQL As String
    If MsgBox("是否要删除用户,此用户删除后将不能恢复?", vbExclamation + vbOKCancel + vbDefaultButton2, "清除") = vbOK Then
        'cnn.InitializeConnection
        Gm.cnnTool.IniConnection8DM Gm.SysID.DBInfo
        
        strSQL = "Delete From G_UserPro Where B_UserName='" & Trim(m_NowUserName) & "'"
        Gm.cnnTool.cnn.Execute strSQL
        
        strSQL = "Delete From G_SUserSub Where B_UserName='" & Trim(m_NowUserName) & "'"
        Gm.cnnTool.cnn.Execute strSQL
        
        strSQL = "Delete From G_SystemUser Where B_UserName='" & Trim(m_NowUserName) & "'"
        Gm.cnnTool.cnn.Execute strSQL
        
        FillTreeView
    End If
End Sub

'清除口令
Private Sub ClearPassWord()
    If Len(m_NowUserName) < 1 Then
        Exit Sub
    End If
    If MsgBox("是否要清除口令?", vbExclamation + vbOKCancel + vbDefaultButton2, "清除") = vbOK Then
        Dim strSQL As String
        
        strSQL = "Update G_SystemUser Set B_Password='' Where B_UserName ='" & m_NowUserName & "'"
        Gm.cnnTool.cnn.Execute strSQL


        MsgBox "口令已清除!", vbInformation, "清除"
    End If
End Sub


'初始化单据
Private Sub FillGrid()
    On Error GoTo IFERR
    If Len(m_NowUserName) < 1 Then
        Exit Sub
    End If
    Dim strSQL As String
   
    strSQL = "Select B_SubSystem From G_SubSystem Where B_SubSystem Not In ("
    strSQL = strSQL & "Select B_SubSystem From G_SUserSub Where B_UserName='" & Trim(m_NowUserName) & "'"
    strSQL = strSQL & ")"
    Debug.Print strSQL
    Debug.Print Gm.cnnTool.cnnStr
    With Adodc2
        .ConnectionString = Gm.cnnTool.cnnStr
        .RecordSource = strSQL
        .Refresh
    End With

    
    strSQL = "Select B_ID,B_UserName,B_SubSystem  From G_SUserSub Where B_UserName ='" & m_NowUserName & "'"

    With Adodc3
        .ConnectionString = Gm.cnnTool.cnnStr
        .RecordSource = strSQL
        .Refresh
    End With

    
    strSQL = "Select G_UserPro.B_ObjectID,G_MenuItems.B_ObjectName From G_UserPro,G_MenuItems Where G_UserPro.B_ObjectID=G_MenuItems.B_ObjectID And B_userName='" & Trim(m_NowUserName) & "' Order By G_UserPro.B_ObjectID"
    With Adodc5
        .ConnectionString = Gm.cnnTool.cnnStr
        .RecordSource = strSQL
        .Refresh
    End With
    
    Dim rs1 As New RecordSet
    Dim rs2 As New RecordSet
    
    Set rs1 = New RecordSet
    Set rs2 = New RecordSet
    
    strSQL = "Select B_ObjectID,B_MenuItem,B_MenuClass From G_MenuItems Where len(isnull(B_ObjectID,''))>0"
    rs1.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    strSQL = "Select * From G_UserPro Where B_UserName='" & m_NowUserName & "'"
    rs2.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    Set rs = Nothing
    
    rs.Fields.Append "B_Check", adInteger, 4, adFldIsNullable
    rs.Fields.Append "B_ObjectID", adVarChar, 20, adFldIsNullable
    rs.Fields.Append "B_MenuItem", adVarChar, 300, adFldIsNullable
    rs.Fields.Append "B_MenuClass", adVarChar, 40, adFldIsNullable
    
    '在2016年6月20日 19:18:45更新为细颗粒度权限系统时添加
    rs.Fields.Append "B_New", adVarChar, 40, adFldIsNullable
    rs.Fields.Append "B_Update", adVarChar, 40, adFldIsNullable
    rs.Fields.Append "B_Delete", adVarChar, 40, adFldIsNullable
    rs.Fields.Append "B_Print", adVarChar, 40, adFldIsNullable
    rs.Open
    
    Do While Not rs1.EOF
        rs.AddNew
        
        rs("B_Check") = 0
        rs("B_ObjectID") = rs1("B_ObjectID")
        rs("B_MenuItem") = rs1("B_MenuItem")
        rs("B_MenuClass") = rs1("B_MenuClass")
        
        rs!B_new = 0
        rs!B_Update = 0
        rs!B_Delete = 0
            
        rs.Update
    
        rs1.movenext
    Loop
    
    Do While Not rs2.EOF
        rs.Filter = ""
        rs.Filter = "B_ObjectID='" & rs2("B_ObjectID") & "'"
        
        If rs2("B_ObjectID") = "190200" Then
            Debug.Print "当前对象编号：190200"
        End If
        
        If rs.RecordCount > 0 Then
            rs("B_Check") = 1
            rs!B_new = IIf(IsNull(rs2!B_new), 0, rs2!B_new)
            rs!B_Update = IIf(IsNull(rs2!B_Update), 0, rs2!B_Update)
            rs!B_Delete = IIf(IsNull(rs2!B_Delete), 0, rs2!B_Delete)
            'rs!B_Print = IIf(IsNull(rs2!B_Print), 0, rs2!B_Print)
            
            rs.Update
        End If
        
        rs2.movenext
    Loop
    
    rs.Filter = ""
    If Not rs.EOF Then
        rs.MoveFirst
        
    End If
    
    rs.Sort = " B_ObjectID"
    Set TDBGrid4.DataSource = rs
    
    
    '设置网格列的可编辑性
    Dim cls1 As New clsGridShow
    Dim szFields As String
    szFields = "B_New,B_Update,B_Delete,B_Check"
    cls1.SetColLockedExcept TDBGrid4, szFields, ","
    
    
    Exit Sub
IFERR:
    Dim szErr As String
    szErr = "错误发生于获取指定用户的权限时" & vbNewLine & Err.Description
    MsgBox szErr, vbOKOnly + vbInformation, "警告"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Dim cls1 As New clspI
    cls1.RefreshFrmsInCache
    
    Set cls1 = Nothing
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    
    Dim bk1 As String
    
    Dim bk2 As String
    
    bk1 = IIf(IsNull(LastRow), "", LastRow)
    bk2 = TDBGrid1.bookmark
    
    If bk1 <> bk2 Then
        GetProbDetail
    End If
End Sub

Private Function GetFilter() As String
    On Error Resume Next
    Dim tmp As String
    Dim n As Integer
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
                    tmp = tmp & Col.DataField & " LIKE '" & Col.FilterText & "%'"
            End Select

        End If
    Next Col
                
    GetFilter = tmp
End Function

Private Sub TDBGrid4_FilterChange()
'    On Error GoTo errHandler
'
'    Set cols = TDBGrid4.Columns
'    Dim c As Integer
'    c = TDBGrid4.col
'    TDBGrid4.HoldFields
'    rs.Filter = GetFilter()
'    TDBGrid4.col = c
'    TDBGrid4.EditActive = True
'    Exit Sub
'
'errHandler:
'    MsgBox Err.Source & ":" & vbCrLf & Err.Description


    ExeTDBGridFilterChange TDBGrid4, rs
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key <> "F" Then
        m_NowUserName = Node.Key
        Debug.Print m_NowUserName
        Dim rs As New RecordSet
        Dim sql As String
        sql = "select * from G_SystemUser where B_UserDes='" & m_NowUserName & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
        m_NowUserName = rs!B_UserName
        FillGrid
    Else
        m_NowUserName = ""
    End If
End Sub

Private Sub GetMenuClass()
    
    With Adodc1
        .ConnectionString = Gm.cnnTool.cnnStr
        .RecordSource = "Select * From G_MenuClass"
        .Refresh
    End With
End Sub

'取得明细
Private Sub GetProbDetail()
    'rs
    rs.Filter = "B_MenuClass='" & Adodc1.RecordSet("B_MenuClass") & "'"
End Sub

Private Sub ActiveBar22_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "确认权限"
            ConfirmedAuthority
        Case "选择"
            SelectOne
        
        Case "不选"
            UnSelectOne
        
        Case "全选"
            SelectAll
        
        Case "全否"
            UnSelectAll
            
            
        Case "行全选"
            Row_SelectAll
        Case "行全不选"
            Row_SelectAllNo
        Case "行反选"
            Row_SelectRev
        
        Case "列全选"
            Col_SelectAll
            
        Case "列全不选"
            Col_SelectAllNo
        Case "列反选"
            Col_SelectRev
    End Select
End Sub

'选择
Private Sub SelectOne()
    On Error GoTo IFERR

    Dim rs1 As New RecordSet
    Dim Row As Variant
    
    If TDBGrid4.SelRange Then
        Set rs1 = rs
        For Each Row In TDBGrid4.SelBookmarks
            rs1.bookmark = Row
            
            rs1("B_Check") = 1
            rs1.Update
        Next Row
    End If
    TDBGrid4.Update
    Exit Sub
IFERR:
    Exit Sub
End Sub

'不选
Private Sub UnSelectOne()
    On Error GoTo IFERR

    Dim rs1 As New RecordSet
    Dim Row As Variant
    
    If TDBGrid4.SelRange Then
        Set rs1 = rs
        For Each Row In TDBGrid4.SelBookmarks
            rs1.bookmark = Row
            
            rs1("B_Check") = 0
            rs1.Update
        Next Row
    End If
    TDBGrid4.Update
    Exit Sub
IFERR:
    Exit Sub
End Sub

'全选
Private Sub SelectAll()
    On Error Resume Next
    rs.MoveFirst
    Do While Not rs.EOF
        rs("B_Check") = 1
        rs.Update
        
        rs.movenext
    Loop
    rs.MoveFirst
End Sub

'全否
Private Sub UnSelectAll()
    On Error Resume Next
    rs.MoveFirst
    Do While Not rs.EOF
        rs("B_Check") = 0
        rs.Update
        
        rs.movenext
    Loop
    rs.MoveFirst
End Sub

'确认权限
Private Sub ConfirmedAuthority()
    '根据权限进行设置
    Dim rs1 As New RecordSet
    Dim strSQL As String
    Dim lTemp As Long
    Dim rsUI As New RecordSet
'    Dim rs2 As New RecordSet
'    Dim sql As String
'
'    sql = "select B_UserName from G_SystemUser where B_UserDes='" & m_NowUserName & "'"
'    rs2.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Set rs1 = New RecordSet
'
    strSQL = "Select * From G_UserPro Where B_UserName='" & m_NowUserName & "'"
'    strSQL = "Select * From G_UserPro Where B_UserName='" & rs2!B_UserName & "'"
    Debug.Print strSQL
    rs1.Open strSQL, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    
    Set rsUI = rs.Clone
    rsUI.MoveFirst
    rsUI.Filter = ""
    Do While Not rsUI.EOF
        rs1.Filter = "B_ObjectID='" & Trim(rsUI("B_ObjectID")) & "'"
        If rsUI("B_ObjectID") = "11G009" Then
            Debug.Print "AAA"
        End If
        If rsUI("B_Check") = 1 Or rsUI("B_Check") = -1 Then
            '已选择
            If rs1.RecordCount < 1 Then
                rs1.Filter = ""
                rs1.AddNew
            End If
            
            rs1("B_UserName") = m_NowUserName
'            rs1("B_UserName") = rs2!B_UserName
            rs1("B_ObjectID") = rsUI("B_ObjectID")
        
        
            'ltemp = IIf(IsNull(TDBGrid4.Columns("B_New").Value), 0, TDBGrid4.Columns("B_New").Value)
            lTemp = IIf(IsNull(rsUI!B_new), 0, rsUI!B_new)
            rs1!B_new = Abs(lTemp)
            
            'ltemp = IIf(IsNull(TDBGrid4.Columns("B_Update").Value), 0, TDBGrid4.Columns("B_Update").Value)
            lTemp = IIf(IsNull(rsUI!B_Update), 0, rsUI!B_Update)
            rs1!B_Update = Abs(lTemp)
            
            
            
            'ltemp = IIf(IsNull(TDBGrid4.Columns("B_Delete").Value), 0, TDBGrid4.Columns("B_Delete").Value)
            lTemp = IIf(IsNull(rsUI!B_Delete), 0, rsUI!B_Delete)
            rs1!B_Delete = Abs(lTemp)
            
            rs1.Update
                
        Else
            '未选择
            If rs1.RecordCount > 0 Then
                '删除
'                strSQL = "Delete From G_UserPro Where B_ID=" & rs1("B_ID")
'                Gm.cnnTool.cnn.Execute strSQL
                rs1.delete
                rs1.Update
            End If
        End If

        rsUI.movenext
    Loop
    
    rs1.Close
    Set rs1 = Nothing
    
    
    '重新读取权限
    Gm.Authority.InitClass
End Sub

Private Sub Row_SelectAll()
'    TDBGrid4.Columns("B_New").Value = 1
'    TDBGrid4.Columns("B_Update").Value = 1
'    TDBGrid4.Columns("B_Delete").Value = 1
    rs!B_new = 1
    rs!B_Update = 1
    rs!B_Delete = 1
End Sub

Private Sub Row_SelectAllNo()
    TDBGrid4.Columns("B_New").Value = 0
    TDBGrid4.Columns("B_Update").Value = 0
    TDBGrid4.Columns("B_Delete").Value = 0
End Sub

Private Sub Row_SelectRev()
    If TDBGrid4.Columns("B_New").Value = 0 Then
        TDBGrid4.Columns("B_New").Value = 1
    Else
        TDBGrid4.Columns("B_New").Value = 0
    End If
    
    If TDBGrid4.Columns("B_Update").Value = 0 Then
        TDBGrid4.Columns("B_Update").Value = 1
    Else
        TDBGrid4.Columns("B_Update").Value = 0
    End If
    
    If TDBGrid4.Columns("B_Delete").Value = 0 Then
        TDBGrid4.Columns("B_Delete").Value = 1
    Else
        TDBGrid4.Columns("B_Delete").Value = 0
    End If
End Sub

Private Sub Col_SelectAll()
    Dim lBookmark As Long
    
    lBookmark = TDBGrid4.bookmark
    
    TDBGrid4.MoveFirst
    Do While Not TDBGrid4.EOF
        TDBGrid4.Columns(TDBGrid4.Col).Value = 1
        TDBGrid4.movenext
    Loop
    
    TDBGrid4.bookmark = lBookmark
End Sub

Private Sub Col_SelectAllNo()
    Dim lBookmark As Long
    
    lBookmark = TDBGrid4.bookmark
    
    TDBGrid4.MoveFirst
    Do While Not TDBGrid4.EOF
        TDBGrid4.Columns(TDBGrid4.Col).Value = 0
        TDBGrid4.movenext
    Loop
    
    TDBGrid4.bookmark = lBookmark
End Sub

Private Sub Col_SelectRev()
    Dim lBookmark As Long
    
    lBookmark = TDBGrid4.bookmark
    
    TDBGrid4.MoveFirst
    Do While Not TDBGrid4.EOF
        If TDBGrid4.Columns(TDBGrid4.Col).Value = 0 Then
            TDBGrid4.Columns(TDBGrid4.Col).Value = 1
        Else
            TDBGrid4.Columns(TDBGrid4.Col).Value = 0
        End If
        TDBGrid4.movenext
    Loop
    
    TDBGrid4.bookmark = lBookmark
End Sub


Private Sub ExeTDBGridFilterChange(ByRef vTDBGrid As TDBGrid, ByRef vRs As RecordSet)
    On Error GoTo IFERR
    Dim Col As Integer
    Col = vTDBGrid.Col
       
    vTDBGrid.HoldFields
    vRs.Filter = GetTDBGridFilterString(vTDBGrid)
    vTDBGrid.Col = Col
    vTDBGrid.EditActive = True
       
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

Private Sub CopyPermissions()
Dim m_MeUserName As String
'm_NowUserName
m_MeUserName = Gm.SysID.SystemUser

 If MsgBox("是否把当前用户的权限复制给我选中的用户？", vbExclamation + vbOKCancel + vbDefaultButton2, "删除") = vbOK Then
        Dim sql As String
        sql = "exec usp_CopyPermissions'" & m_MeUserName & "','" & m_NowUserName & "'"
        
        Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
        
         MsgBox "复制完毕!", vbInformation, "完成"
End If

End Sub
