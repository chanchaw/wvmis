VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "Actbar2.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.2#0"; "Codejock.Controls.v16.2.4.ocx"
Begin VB.Form frmColor_Edit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ɫ���ƻ�"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColor_Edit.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   16440
   StartUpPosition =   2  '��Ļ����
   Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar22 
      Height          =   9600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16440
      _LayoutVersion  =   1
      _ExtentX        =   28998
      _ExtentY        =   16933
      _DataPath       =   ""
      Bands           =   "frmColor_Edit.frx":038A
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   9135
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   16455
         _cx             =   29025
         _cy             =   16113
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   800
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   3263743
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "ɫ���ɹ�|ɫ������"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   6
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   0   'False
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   3
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
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar23 
            Height          =   9135
            Left            =   17070
            TabIndex        =   113
            Top             =   330
            Width           =   16455
            _LayoutVersion  =   1
            _ExtentX        =   29025
            _ExtentY        =   16113
            _DataPath       =   ""
            Bands           =   "frmColor_Edit.frx":0552
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   6435
               Left            =   480
               TabIndex        =   114
               TabStop         =   0   'False
               Top             =   960
               Width           =   10515
               _cx             =   18547
               _cy             =   11351
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
               GridRows        =   4
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmColor_Edit.frx":1CF2
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.PictureBox Picture3 
                  BorderStyle     =   0  'None
                  Height          =   4665
                  Left            =   90
                  ScaleHeight     =   4665
                  ScaleWidth      =   10335
                  TabIndex        =   115
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   10335
                  Begin XtremeSuiteControls.FlatEdit FlatEdit38 
                     Height          =   375
                     Left            =   4740
                     TabIndex        =   116
                     Top             =   2796
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
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
                  End
                  Begin XtremeSuiteControls.PushButton PushButton10 
                     Height          =   375
                     Left            =   6300
                     TabIndex        =   117
                     Top             =   180
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit39 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1440
                     TabIndex        =   118
                     Top             =   180
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit40 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   4740
                     TabIndex        =   119
                     Top             =   1012
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit41 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   8040
                     TabIndex        =   120
                     Top             =   180
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit42 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1440
                     TabIndex        =   121
                     Top             =   1012
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit43 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1440
                     TabIndex        =   122
                     Top             =   1904
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit44 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   8040
                     TabIndex        =   123
                     Top             =   1904
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit45 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   4740
                     TabIndex        =   124
                     Top             =   180
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit46 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1440
                     TabIndex        =   125
                     Top             =   2796
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit47 
                     Height          =   975
                     Left            =   4740
                     TabIndex        =   126
                     Top             =   3600
                     Width           =   5115
                     _Version        =   1048578
                     _ExtentX        =   9022
                     _ExtentY        =   1720
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     MultiLine       =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit48 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   8040
                     TabIndex        =   127
                     Top             =   1012
                     Width           =   1695
                     _Version        =   1048578
                     _ExtentX        =   2990
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit49 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   4740
                     TabIndex        =   128
                     Top             =   1904
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.DateTimePicker DateTimePicker1 
                     Height          =   375
                     Left            =   8040
                     TabIndex        =   129
                     Top             =   2796
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   68
                     Format          =   1
                  End
                  Begin XtremeSuiteControls.PushButton PushButton11 
                     Height          =   375
                     Left            =   9720
                     TabIndex        =   130
                     Top             =   1012
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit50 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1440
                     TabIndex        =   145
                     Top             =   3660
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label63 
                     Height          =   255
                     Left            =   360
                     TabIndex        =   146
                     Top             =   3720
                     Width           =   615
                     _Version        =   1048578
                     _ExtentX        =   1085
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "���ۣ�"
                     ForeColor       =   255
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label62 
                     Height          =   315
                     Left            =   3600
                     TabIndex        =   143
                     Top             =   3690
                     Width           =   975
                     _Version        =   1048578
                     _ExtentX        =   1720
                     _ExtentY        =   556
                     _StockProps     =   79
                     Caption         =   "��ע:"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label61 
                     Height          =   255
                     Left            =   3600
                     TabIndex        =   142
                     Top             =   2856
                     Width           =   555
                     _Version        =   1048578
                     _ExtentX        =   979
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "����:"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label60 
                     Height          =   255
                     Left            =   7080
                     TabIndex        =   141
                     Top             =   1964
                     Width           =   855
                     _Version        =   1048578
                     _ExtentX        =   1508
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "��������"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label59 
                     Height          =   255
                     Left            =   360
                     TabIndex        =   140
                     Top             =   2856
                     Width           =   615
                     _Version        =   1048578
                     _ExtentX        =   1085
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "������"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label58 
                     Height          =   375
                     Left            =   360
                     TabIndex        =   139
                     Top             =   1904
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "ɫ�ţ�"
                     ForeColor       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label57 
                     Height          =   255
                     Left            =   3480
                     TabIndex        =   138
                     Top             =   240
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "ɫ����Ӧ�̣�"
                     ForeColor       =   255
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label56 
                     Height          =   255
                     Left            =   360
                     TabIndex        =   137
                     Top             =   1072
                     Width           =   615
                     _Version        =   1048578
                     _ExtentX        =   1085
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "���أ�"
                     ForeColor       =   255
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label55 
                     Height          =   255
                     Left            =   6960
                     TabIndex        =   136
                     Top             =   240
                     Width           =   615
                     _Version        =   1048578
                     _ExtentX        =   1085
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "�ŷ���"
                     ForeColor       =   255
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label54 
                     Height          =   255
                     Left            =   3600
                     TabIndex        =   135
                     Top             =   1072
                     Width           =   615
                     _Version        =   1048578
                     _ExtentX        =   1085
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "Ʒ����"
                     ForeColor       =   255
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label53 
                     Height          =   255
                     Left            =   240
                     TabIndex        =   134
                     Top             =   240
                     Width           =   855
                     _Version        =   1048578
                     _ExtentX        =   1508
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "�����ţ�"
                     ForeColor       =   255
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label52 
                     Height          =   255
                     Left            =   7080
                     TabIndex        =   133
                     Top             =   1072
                     Width           =   615
                     _Version        =   1048578
                     _ExtentX        =   1085
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "��ɫ��"
                     ForeColor       =   255
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label51 
                     Height          =   255
                     Left            =   3600
                     TabIndex        =   132
                     Top             =   1964
                     Width           =   615
                     _Version        =   1048578
                     _ExtentX        =   1085
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "ƥ����"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label50 
                     Height          =   255
                     Left            =   7080
                     TabIndex        =   131
                     Top             =   2856
                     Width           =   855
                     _Version        =   1048578
                     _ExtentX        =   1508
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "���ڣ�"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
               End
               Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
                  Height          =   1530
                  Left            =   90
                  TabIndex        =   144
                  Top             =   4815
                  Width           =   10335
                  _ExtentX        =   18230
                  _ExtentY        =   2699
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
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=4339"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
                  Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
                  Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(6)=   "Column(1).Width=4339"
                  Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
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
                  _StyleDefs(8)   =   ":id=1,.fontname=����"
                  _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgpicMode=2,.bgbmp=1,.bold=0"
                  _StyleDefs(11)  =   ":id=2,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=134"
                  _StyleDefs(12)  =   ":id=2,.fontname=����"
                  _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgpicMode=2,.bgbmp=2,.bold=0"
                  _StyleDefs(14)  =   ":id=3,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=134"
                  _StyleDefs(15)  =   ":id=3,.fontname=����"
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
         Begin ActiveBar2LibraryCtl.ActiveBar2 ActiveBar21 
            Height          =   9135
            Left            =   15
            TabIndex        =   2
            Top             =   330
            Width           =   16455
            _LayoutVersion  =   1
            _ExtentX        =   29025
            _ExtentY        =   16113
            _DataPath       =   ""
            Bands           =   "frmColor_Edit.frx":1D78
            Begin C1SizerLibCtl.C1Elastic C1Elastic1 
               Height          =   9075
               Left            =   0
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   720
               Width           =   15975
               _cx             =   28178
               _cy             =   16007
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
               GridRows        =   4
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmColor_Edit.frx":4AFE
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.PictureBox Picture1 
                  BorderStyle     =   0  'None
                  Height          =   8895
                  Left            =   90
                  ScaleHeight     =   8895
                  ScaleWidth      =   15795
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   15795
                  Begin VB.PictureBox Picture2 
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     ForeColor       =   &H80000008&
                     Height          =   375
                     Left            =   5400
                     ScaleHeight     =   345
                     ScaleWidth      =   1905
                     TabIndex        =   7
                     TabStop         =   0   'False
                     Top             =   660
                     Width           =   1935
                  End
                  Begin VB.Frame Frame2 
                     Height          =   75
                     Left            =   60
                     TabIndex        =   6
                     Top             =   7020
                     Width           =   14895
                  End
                  Begin VB.Frame Frame1 
                     Height          =   75
                     Left            =   180
                     TabIndex        =   5
                     Top             =   2160
                     Width           =   14955
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit33 
                     Height          =   375
                     Left            =   5400
                     TabIndex        =   8
                     Top             =   1740
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit22 
                     Height          =   375
                     Left            =   7620
                     TabIndex        =   9
                     Top             =   3495
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
                     _ExtentY        =   661
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
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit18 
                     Height          =   375
                     Left            =   7620
                     TabIndex        =   10
                     Top             =   2385
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
                     _ExtentY        =   661
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
                  End
                  Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
                     Height          =   1275
                     Left            =   12840
                     TabIndex        =   11
                     Top             =   2760
                     Width           =   2055
                     _ExtentX        =   3625
                     _ExtentY        =   2249
                     _LayoutType     =   0
                     _RowHeight      =   -2147483647
                     _WasPersistedAsPixels=   0
                     Columns(0)._VlistStyle=   0
                     Columns(0)._MaxComboItems=   5
                     Columns(0).DataField=   ""
                     Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                     Columns.Count   =   1
                     Splits(0)._UserFlags=   0
                     Splits(0).RecordSelectors=   0   'False
                     Splits(0).RecordSelectorWidth=   953
                     Splits(0)._SavedRecordSelectors=   0   'False
                     Splits(0).ScrollBars=   2
                     Splits(0).AllowColSelect=   0   'False
                     Splits(0).DividerColor=   15790320
                     Splits(0).SpringMode=   0   'False
                     Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                     Splits(0)._ColumnProps(0)=   "Columns.Count=1"
                     Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
                     Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                     Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
                     Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
                     Splits.Count    =   1
                     PrintInfos(0)._StateFlags=   3
                     PrintInfos(0).Name=   "piInternal 0"
                     PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
                     PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
                     PrintInfos(0).PageHeaderHeight=   0
                     PrintInfos(0).PageFooterHeight=   0
                     PrintInfos.Count=   1
                     AllowUpdate     =   0   'False
                     DefColWidth     =   0
                     HeadLines       =   1
                     FootLines       =   1
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
                     _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=114,.bold=0,.fontsize=825,.italic=0"
                     _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
                     _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
                     _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
                     _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
                     _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
                     _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                     _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
                     _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
                     _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
                     _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
                     _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
                     _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
                     _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                     _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
                     _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
                     _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
                     _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
                     _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
                     _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
                     _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
                     _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
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
                     _StyleDefs(40)  =   "Named:id=33:Normal"
                     _StyleDefs(41)  =   ":id=33,.parent=0"
                     _StyleDefs(42)  =   "Named:id=34:Heading"
                     _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                     _StyleDefs(44)  =   ":id=34,.wraptext=-1"
                     _StyleDefs(45)  =   "Named:id=35:Footing"
                     _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                     _StyleDefs(47)  =   "Named:id=36:Selected"
                     _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                     _StyleDefs(49)  =   "Named:id=37:Caption"
                     _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
                     _StyleDefs(51)  =   "Named:id=38:HighlightRow"
                     _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                     _StyleDefs(53)  =   "Named:id=39:EvenRow"
                     _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                     _StyleDefs(55)  =   "Named:id=40:OddRow"
                     _StyleDefs(56)  =   ":id=40,.parent=33"
                     _StyleDefs(57)  =   "Named:id=41:RecordSelector"
                     _StyleDefs(58)  =   ":id=41,.parent=34"
                     _StyleDefs(59)  =   "Named:id=42:FilterBar"
                     _StyleDefs(60)  =   ":id=42,.parent=33"
                  End
                  Begin MSComCtl2.DTPicker DTPicker1 
                     Height          =   375
                     Left            =   10800
                     TabIndex        =   12
                     Top             =   2385
                     Width           =   1575
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   219283457
                     CurrentDate     =   43041
                  End
                  Begin XtremeSuiteControls.PushButton PushButton1 
                     Height          =   375
                     Left            =   3480
                     TabIndex        =   13
                     Top             =   660
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     Enabled         =   0   'False
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit2 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   5400
                     TabIndex        =   14
                     Top             =   120
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit5 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   12720
                     TabIndex        =   15
                     Top             =   660
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit6 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   16
                     Top             =   1200
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit7 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   5400
                     TabIndex        =   17
                     Top             =   1200
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit8 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   18
                     Top             =   2385
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit9 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   5400
                     TabIndex        =   19
                     Top             =   2385
                     Width           =   975
                     _Version        =   1048578
                     _ExtentX        =   1720
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit11 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   20
                     Top             =   660
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit15 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   5400
                     TabIndex        =   21
                     Top             =   3495
                     Width           =   975
                     _Version        =   1048578
                     _ExtentX        =   1720
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit19 
                     Bindings        =   "frmColor_Edit.frx":4B85
                     DataField       =   "B_CodeID"
                     BeginProperty DataFormat 
                        Type            =   1
                        Format          =   "0"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   2052
                        SubFormatType   =   1
                     EndProperty
                     DataMember      =   "0.00"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   22
                     Top             =   7260
                     Width           =   735
                     _Version        =   1048578
                     _ExtentX        =   1296
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit20 
                     DataField       =   "B_CodeID"
                     BeginProperty DataFormat 
                        Type            =   0
                        Format          =   "0"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   2052
                        SubFormatType   =   0
                     EndProperty
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   5400
                     TabIndex        =   23
                     Top             =   7260
                     Width           =   735
                     _Version        =   1048578
                     _ExtentX        =   1296
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit21 
                     DataField       =   "B_CodeID"
                     BeginProperty DataFormat 
                        Type            =   0
                        Format          =   "0"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   2052
                        SubFormatType   =   0
                     EndProperty
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   9000
                     TabIndex        =   24
                     Top             =   7260
                     Width           =   735
                     _Version        =   1048578
                     _ExtentX        =   1296
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.PushButton PushButton2 
                     Height          =   375
                     Left            =   3480
                     TabIndex        =   25
                     Top             =   2385
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.PushButton PushButton3 
                     Height          =   375
                     Left            =   6360
                     TabIndex        =   26
                     Top             =   2385
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.PushButton PushButton5 
                     Height          =   375
                     Left            =   6360
                     TabIndex        =   27
                     Top             =   3495
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin MSComCtl2.DTPicker DTPicker2 
                     Height          =   375
                     Left            =   10800
                     TabIndex        =   28
                     Top             =   3495
                     Width           =   1575
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   219283457
                     CurrentDate     =   43041
                  End
                  Begin XtremeSuiteControls.ComboBox ComboBox4 
                     Height          =   345
                     Left            =   12780
                     TabIndex        =   29
                     Top             =   7275
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   609
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Style           =   2
                     Text            =   "ComboBox1"
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit1 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   9000
                     TabIndex        =   30
                     Top             =   660
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit3 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   31
                     Top             =   120
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit4 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   9000
                     TabIndex        =   32
                     Top             =   180
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit10 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   12720
                     TabIndex        =   33
                     Top             =   120
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit13 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   435
                     Left            =   1920
                     TabIndex        =   34
                     Top             =   2880
                     Width           =   13275
                     _Version        =   1048578
                     _ExtentX        =   23416
                     _ExtentY        =   767
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MultiLine       =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit17 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   435
                     Left            =   1920
                     TabIndex        =   35
                     Top             =   4050
                     Width           =   13275
                     _Version        =   1048578
                     _ExtentX        =   23416
                     _ExtentY        =   767
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MultiLine       =   -1  'True
                  End
                  Begin XtremeSuiteControls.ComboBox ComboBox5 
                     Height          =   345
                     Left            =   9000
                     TabIndex        =   36
                     Top             =   1215
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   609
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Style           =   2
                     Text            =   "ComboBox1"
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit14 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   37
                     Top             =   3495
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.PushButton PushButton4 
                     Height          =   375
                     Left            =   3480
                     TabIndex        =   38
                     Top             =   3495
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit12 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   14100
                     TabIndex        =   39
                     Top             =   2385
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit16 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   14100
                     TabIndex        =   40
                     Top             =   3495
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit23 
                     Height          =   375
                     Left            =   7620
                     TabIndex        =   41
                     Top             =   4665
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
                     _ExtentY        =   661
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
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit24 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   5400
                     TabIndex        =   42
                     Top             =   4665
                     Width           =   975
                     _Version        =   1048578
                     _ExtentX        =   1720
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.PushButton PushButton6 
                     Height          =   375
                     Left            =   6360
                     TabIndex        =   43
                     Top             =   4665
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit25 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   44
                     Top             =   4665
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.PushButton PushButton7 
                     Height          =   375
                     Left            =   3480
                     TabIndex        =   45
                     Top             =   4665
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit26 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   435
                     Left            =   1920
                     TabIndex        =   46
                     Top             =   5220
                     Width           =   13275
                     _Version        =   1048578
                     _ExtentX        =   23416
                     _ExtentY        =   767
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MultiLine       =   -1  'True
                  End
                  Begin MSComCtl2.DTPicker DTPicker3 
                     Height          =   375
                     Left            =   10800
                     TabIndex        =   47
                     Top             =   4665
                     Width           =   1575
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   219283457
                     CurrentDate     =   43041
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit27 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   14100
                     TabIndex        =   48
                     Top             =   4665
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit28 
                     Height          =   375
                     Left            =   7620
                     TabIndex        =   49
                     Top             =   5835
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
                     _ExtentY        =   661
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
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit29 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   5460
                     TabIndex        =   50
                     Top             =   5835
                     Width           =   975
                     _Version        =   1048578
                     _ExtentX        =   1720
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.PushButton PushButton8 
                     Height          =   375
                     Left            =   6420
                     TabIndex        =   51
                     Top             =   5820
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit30 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   52
                     Top             =   5835
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.PushButton PushButton9 
                     Height          =   375
                     Left            =   3480
                     TabIndex        =   53
                     Top             =   5835
                     Width           =   375
                     _Version        =   1048578
                     _ExtentX        =   661
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   ".."
                     UseVisualStyle  =   -1  'True
                  End
                  Begin MSComCtl2.DTPicker DTPicker4 
                     Height          =   375
                     Left            =   10860
                     TabIndex        =   54
                     Top             =   5835
                     Width           =   1575
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   219283457
                     CurrentDate     =   43041
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit31 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   14100
                     TabIndex        =   55
                     Top             =   5835
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit32 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   435
                     Left            =   1920
                     TabIndex        =   56
                     Top             =   6390
                     Width           =   13275
                     _Version        =   1048578
                     _ExtentX        =   23416
                     _ExtentY        =   767
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   -2147483643
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MultiLine       =   -1  'True
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit34 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   9000
                     TabIndex        =   57
                     Top             =   1740
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit35 
                     Height          =   375
                     Left            =   12720
                     TabIndex        =   58
                     Top             =   1740
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
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
                  End
                  Begin XtremeSuiteControls.ComboBox ComboBox1 
                     Height          =   390
                     Left            =   1920
                     TabIndex        =   59
                     Top             =   7830
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   688
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Style           =   2
                     Text            =   "ComboBox1"
                  End
                  Begin XtremeSuiteControls.ComboBox ComboBox2 
                     Height          =   390
                     Left            =   5400
                     TabIndex        =   60
                     Top             =   7830
                     Width           =   1575
                     _Version        =   1048578
                     _ExtentX        =   2778
                     _ExtentY        =   688
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Style           =   2
                     Text            =   "ComboBox2"
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit36 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   12720
                     TabIndex        =   61
                     Top             =   1200
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   16777215
                  End
                  Begin XtremeSuiteControls.FlatEdit FlatEdit37 
                     DataField       =   "B_CodeID"
                     DataSource      =   "Adodc1"
                     Height          =   375
                     Left            =   1920
                     TabIndex        =   62
                     Top             =   1740
                     Width           =   1935
                     _Version        =   1048578
                     _ExtentX        =   3413
                     _ExtentY        =   661
                     _StockProps     =   77
                     ForeColor       =   0
                     BackColor       =   14737632
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BackColor       =   14737632
                  End
                  Begin XtremeSuiteControls.Label Label49 
                     Height          =   255
                     Left            =   240
                     TabIndex        =   112
                     Top             =   1800
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "��     ����"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label48 
                     Height          =   255
                     Left            =   11280
                     TabIndex        =   111
                     Top             =   1260
                     Width           =   975
                     _Version        =   1048578
                     _ExtentX        =   1720
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "Ⱦ����ɫ:"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label47 
                     Height          =   255
                     Left            =   4140
                     TabIndex        =   110
                     Top             =   7898
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "ϸ�뵥:"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin VB.Label Label46 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "��ǩģ��:"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   240
                     TabIndex        =   109
                     Top             =   7860
                     Width           =   975
                  End
                  Begin XtremeSuiteControls.Label Label45 
                     Height          =   255
                     Left            =   11280
                     TabIndex        =   108
                     Top             =   1800
                     Width           =   1275
                     _Version        =   1048578
                     _ExtentX        =   2249
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "ʵ��Ͷ����:"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label44 
                     Height          =   255
                     Left            =   7680
                     TabIndex        =   107
                     Top             =   1800
                     Width           =   1275
                     _Version        =   1048578
                     _ExtentX        =   2249
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "Ԥ��Ͷ����:"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label43 
                     Height          =   255
                     Left            =   4080
                     TabIndex        =   106
                     Top             =   1800
                     Width           =   1275
                     _Version        =   1048578
                     _ExtentX        =   2249
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "Ԥ������:"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label42 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   105
                     Top             =   6420
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "��ӹ���ע��"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label41 
                     Height          =   195
                     Left            =   12660
                     TabIndex        =   104
                     Top             =   5925
                     Width           =   1455
                     _Version        =   1048578
                     _ExtentX        =   2566
                     _ExtentY        =   344
                     _StockProps     =   79
                     Caption         =   "��ӹ��ӹ���:"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label40 
                     Height          =   375
                     Left            =   9540
                     TabIndex        =   103
                     Top             =   5835
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "��ӹ����ڣ�"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label39 
                     Height          =   375
                     Left            =   4140
                     TabIndex        =   102
                     Top             =   5835
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "��ӹ�������"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label38 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   101
                     Top             =   5835
                     Width           =   1335
                     _Version        =   1048578
                     _ExtentX        =   2355
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "��ӹ���λ3��"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label37 
                     Height          =   375
                     Left            =   7080
                     TabIndex        =   100
                     Top             =   5835
                     Width           =   495
                     _Version        =   1048578
                     _ExtentX        =   873
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "�绰:"
                  End
                  Begin XtremeSuiteControls.Label Label36 
                     Height          =   375
                     Left            =   9480
                     TabIndex        =   99
                     Top             =   4665
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "��ӹ����ڣ�"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label35 
                     Height          =   195
                     Left            =   12660
                     TabIndex        =   98
                     Top             =   4755
                     Width           =   1455
                     _Version        =   1048578
                     _ExtentX        =   2566
                     _ExtentY        =   344
                     _StockProps     =   79
                     Caption         =   "��ӹ��ӹ���:"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label34 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   97
                     Top             =   5250
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "��ӹ���ע��"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label33 
                     Height          =   375
                     Left            =   4080
                     TabIndex        =   96
                     Top             =   4665
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "��ӹ�������"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label32 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   95
                     Top             =   4665
                     Width           =   1335
                     _Version        =   1048578
                     _ExtentX        =   2355
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "��ӹ���λ2��"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label31 
                     Height          =   375
                     Left            =   7080
                     TabIndex        =   94
                     Top             =   4665
                     Width           =   495
                     _Version        =   1048578
                     _ExtentX        =   873
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "�绰:"
                  End
                  Begin XtremeSuiteControls.Label Label30 
                     Height          =   375
                     Left            =   7080
                     TabIndex        =   93
                     Top             =   3495
                     Width           =   495
                     _Version        =   1048578
                     _ExtentX        =   873
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "�绰:"
                  End
                  Begin XtremeSuiteControls.Label Label29 
                     Height          =   255
                     Left            =   7080
                     TabIndex        =   92
                     Top             =   2445
                     Width           =   435
                     _Version        =   1048578
                     _ExtentX        =   767
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "�绰:"
                  End
                  Begin XtremeSuiteControls.Label Label28 
                     Height          =   195
                     Left            =   12660
                     TabIndex        =   91
                     Top             =   3585
                     Width           =   1455
                     _Version        =   1048578
                     _ExtentX        =   2566
                     _ExtentY        =   344
                     _StockProps     =   79
                     Caption         =   "��ӹ��ӹ���:"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label27 
                     Height          =   255
                     Left            =   12660
                     TabIndex        =   90
                     Top             =   2445
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "Ⱦ���ӹ���:"
                     ForeColor       =   255
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label14 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   89
                     Top             =   3495
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "��ӹ���λ��"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label18 
                     Height          =   375
                     Left            =   7680
                     TabIndex        =   88
                     Top             =   1200
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "���ȹ��գ�"
                     ForeColor       =   255
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label26 
                     Height          =   1575
                     Left            =   14880
                     TabIndex        =   87
                     Top             =   0
                     Width           =   855
                     _Version        =   1048578
                     _ExtentX        =   1508
                     _ExtentY        =   2778
                     _StockProps     =   79
                     ForeColor       =   255
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   15
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                  End
                  Begin XtremeSuiteControls.Label Label25 
                     Height          =   255
                     Left            =   9840
                     TabIndex        =   86
                     Top             =   7320
                     Width           =   495
                     _Version        =   1048578
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "����"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label24 
                     Height          =   255
                     Left            =   6240
                     TabIndex        =   85
                     Top             =   7320
                     Width           =   495
                     _Version        =   1048578
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "����"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label12 
                     Height          =   255
                     Left            =   2760
                     TabIndex        =   84
                     Top             =   7320
                     Width           =   495
                     _Version        =   1048578
                     _ExtentX        =   873
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "����"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label23 
                     Height          =   255
                     Left            =   7680
                     TabIndex        =   83
                     Top             =   720
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "��   �ͣ�"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label22 
                     Height          =   255
                     Left            =   4080
                     TabIndex        =   82
                     Top             =   720
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "��ɫ ��ʶ:"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label4 
                     Height          =   255
                     Left            =   11280
                     TabIndex        =   81
                     Top             =   180
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "��     �أ�"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label21 
                     Height          =   255
                     Left            =   11280
                     TabIndex        =   80
                     Top             =   7320
                     Width           =   1335
                     _Version        =   1048578
                     _ExtentX        =   2355
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "��װ��ʽ��"
                     ForeColor       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label20 
                     Height          =   255
                     Left            =   7680
                     TabIndex        =   79
                     Top             =   7320
                     Width           =   975
                     _Version        =   1048578
                     _ExtentX        =   1720
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "��   �ӣ�"
                     ForeColor       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label lbl 
                     Height          =   375
                     Left            =   4080
                     TabIndex        =   78
                     Top             =   7260
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "��    �أ�"
                     ForeColor       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label19 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   77
                     Top             =   7260
                     Width           =   1455
                     _Version        =   1048578
                     _ExtentX        =   2566
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "ֽ      �ܣ�"
                     ForeColor       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label17 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   76
                     Top             =   4080
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "��ӹ���ע��"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label16 
                     Height          =   375
                     Left            =   9480
                     TabIndex        =   75
                     Top             =   3495
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "��ӹ����ڣ�"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label15 
                     Height          =   375
                     Left            =   4080
                     TabIndex        =   74
                     Top             =   3495
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "��ӹ�������"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label13 
                     Height          =   375
                     Left            =   240
                     TabIndex        =   73
                     Top             =   2910
                     Width           =   1695
                     _Version        =   1048578
                     _ExtentX        =   2990
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "Ⱦ��  ��ע��"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label6 
                     Height          =   375
                     Left            =   9480
                     TabIndex        =   72
                     Top             =   2385
                     Width           =   1095
                     _Version        =   1048578
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "Ⱦ�����ڣ�"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label1 
                     Height          =   255
                     Left            =   240
                     TabIndex        =   71
                     Top             =   180
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "��  �� �ţ�"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label2 
                     Height          =   255
                     Left            =   4080
                     TabIndex        =   70
                     Top             =   180
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "Ʒ     ����"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label3 
                     Height          =   255
                     Left            =   7680
                     TabIndex        =   69
                     Top             =   180
                     Width           =   975
                     _Version        =   1048578
                     _ExtentX        =   1720
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "��   ����"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label5 
                     Height          =   255
                     Left            =   240
                     TabIndex        =   68
                     Top             =   720
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "��     ɫ��"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label7 
                     Height          =   375
                     Left            =   11280
                     TabIndex        =   67
                     Top             =   660
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   661
                     _StockProps     =   79
                     Caption         =   "ɫ     �ţ�"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label8 
                     Height          =   255
                     Left            =   240
                     TabIndex        =   66
                     Top             =   1260
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "��     ����"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label9 
                     Height          =   255
                     Left            =   4080
                     TabIndex        =   65
                     Top             =   1260
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "��  �� ����"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label10 
                     Height          =   255
                     Left            =   240
                     TabIndex        =   64
                     Top             =   2445
                     Width           =   1335
                     _Version        =   1048578
                     _ExtentX        =   2355
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "Ⱦ     ����"
                     ForeColor       =   255
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin XtremeSuiteControls.Label Label11 
                     Height          =   255
                     Left            =   4080
                     TabIndex        =   63
                     Top             =   2445
                     Width           =   1215
                     _Version        =   1048578
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "Ⱦ�� ������"
                     ForeColor       =   255
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "����"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmColor_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public colorid As String
Public departid As String
Public departmentid As String
Public processid As String
Public processmentid As String
Public processid2 As String
Public processmentid2 As String
Public processid3 As String
Public processmentid3 As String
Public ordertocolorid As String '�����н������ӵĶ�����ϸ������
Public lGroupID As Long
Private num As Long '��ʶ������������
Public id As String
Public Valuation As String
Public itemid As String
Private theidColor As String
Private a As Long

Private client As String
Public colororderid As String
Public colorplanid As String
Private B_orderitemid As String
Private rss As RecordSet


Private Type RGB
        Red   As Byte
        Green   As Byte
        Blue   As Byte
End Type

'Private theOrderID As String
'
'Public Property Let OrderID(ByVal vData As String)
'    theOrderID = vData
'End Property
'
'Private Sub FillData()
'    If Len(theOrderID) <= 0 Then
'        Exit Sub
'    End If
'
'    SetComboListIndex theOrderID
'End Sub
'Private Sub SetComboListIndex(ByVal vData As String)
'    Dim i As Long
'    For i = 0 To ComboBox1.ListCount - 1
'        If ComboBox1.List(i) = vData Then
'            ComboBox1.ListIndex = i
'        End If
'    Next
'End Sub
Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "����"
            save (0)
        Case "���沢���Ʊ���ͬ"
            save (1)
'            Saveandcopy_1
        Case "���沢���Ʊ�����"
            save (2)
'            Saveandcopy
        Case "��һ��"
            movefrist
        Case "��һ��"
            moveshang
        Case "��һ��"
            movenext
        Case "���"
            movelast
        Case "�˳�"
            Unload Me
    End Select
End Sub
    


Private Sub ComboBox5_Click()
'        If Len(ComboBox5.Text) > 0 Then
            Dim rs As New RecordSet
            Dim sql As String
            sql = "Select B_ProgressItem From G_ProgressCraft where B_Parent='" & ComboBox5.Text & "'"
            rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
'        End If
        TDBGrid1.DataSource = rs
        TDBGrid1.Columns("B_ProgressItem").Caption = "���ȹ�����ϸ"
        TDBGrid1.MarqueeStyle = dbgHighlightRow
End Sub


Private Sub FlatEdit11_Change()
        On Error Resume Next
        Dim rs As New RecordSet
        Dim sql As String
        sql = "select * from G_Color where B_SID='" & colorid & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

        Dim ys As RGB
        With ys
            .Red = rs!B_Red
            .Green = rs!B_Green
            .Blue = rs!B_Blue
        End With
        Picture2.BackColor = rs!B_hex
End Sub

Private Sub FlatEdit33_Change()
    Dim f As Long
    If Val(FlatEdit33.Text) <> 0 Then
        f = Val(FlatEdit7.Text) * (Val(FlatEdit33.Text) + 1)
        FlatEdit34.Text = f
    End If
End Sub


Private Sub Form_Load()
    InitFrm
'    dingdanhao
    baozhuang
    process
'    FillData
    TDBGrid1.Visible = False
    DateTimePicker1.Value = Now
    Grid
    C1Tab1.CurrTab = 1
End Sub
Private Sub InitFrm()
    With ActiveBar21
        .ClientAreaControl = C1Elastic1
        .RecalcLayout
    End With
    With ActiveBar22
        .ClientAreaControl = C1Tab1
        .RecalcLayout
    End With
        With ActiveBar23
        .ClientAreaControl = C1Elastic2
        .RecalcLayout
    End With
    DTPicker1.Value = Now
    DTPicker2.Value = Now
    DTPicker3.Value = Now
    DTPicker4.Value = Now
    setValuation
    Sum
    biaoq
    xima
End Sub

'��ǩģ��
Private Sub biaoq()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "SELECT *From G_JBCPSet Where B_Effective = 1 ORDER BY B_Order"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Do While Not rs.EOF
        ComboBox1.AddItem "" & rs!B_GroupName & ""
        rs.movenext
    Loop
End Sub
'ϸ��
Private Sub xima()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "SELECT * From G_JRKXMDYS Where B_Effective = 1 ORDER BY B_Order"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    Do While Not rs.EOF
        ComboBox2.AddItem "" & rs!B_GroupName & ""
        rs.movenext
    Loop
End Sub
Private Sub Sum()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_BillDetailColor a inner join G_BillColor b on a.B_ID=b.B_ID where b.B_BelongOrderID='" & id & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    a = rs.RecordCount
End Sub


Private Sub baozhuang()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "Select B_SID From G_PackWay Where 1=1"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

    Do While Not rs.EOF
        ComboBox4.AddItem rs!B_sid
        rs.movenext
    Loop
End Sub
Private Sub process()
    Dim rs As New RecordSet
    Dim sql As String
    sql = " select B_SID from G_ProgressCraftCT where 1=1"
    Debug.Print sql
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

    Do While Not rs.EOF
        ComboBox5.AddItem rs!B_sid
        rs.movenext
    Loop
End Sub






'---------------------------------------���ð�ť�Ĵ����¼�
Private Sub PushButton1_Click()
      On Error Resume Next
      Dim sql As String
      Dim rs As New RecordSet
      Dim frm1 As New frmpopupColor
      frm1.Show vbModal
      FlatEdit11.Text = Trim(frm1.colorname)
      colorid = frm1.colorid

        sql = "select * from G_Color where B_SID='" & colorid & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly

        Dim ys As RGB
        With ys
            .Red = rs!B_Red
            .Green = rs!B_Green
            .Blue = rs!B_Blue
        End With
        Picture2.BackColor = rs!B_hex
      Unload frm1
End Sub

'Private Sub PushButton1_FetchCellStyle(ByVal Condition As Integer, _
'    ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, _
'    ByVal CellStyle As TrueOleDBGrid80.StyleDi)
'    On Error Resume Next
'    Dim ys As RGB
''    CellStyle.BackColor = TDBGrid3.Columns("B_Hex").CellValue(Bookmark)
'    CellStyle.ForeColor = PushButton1.CellValue(Bookmark)
'End Sub

Private Sub PushButton2_Click()
    Dim frm1 As New frmpopupDepart
      frm1.Show vbModal
      FlatEdit8.Text = Trim(frm1.departName)
     departid = frm1.departid
      Unload frm1
End Sub

Private Sub PushButton3_Click()
    Dim frm1 As New frmpopupDepartment
      frm1.Show vbModal
      FlatEdit9.Text = Trim(frm1.departmentName)
      FlatEdit18.Text = frm1.phone
     departmentid = frm1.departmentid
      Unload frm1
End Sub

Private Sub PushButton4_Click()
    Dim frm1 As New frmpopupProcess
    frm1.Show vbModal
    FlatEdit14.Text = Trim(frm1.processName)
    
    processid = frm1.processid
    Unload frm1
End Sub

Private Sub PushButton5_Click()
    Dim frm1 As New frmpopupProcessment
    frm1.Show vbModal
    FlatEdit15.Text = Trim(frm1.processmentName)
    FlatEdit22.Text = frm1.phone
    processmentid = frm1.processmentid
    Unload frm1
End Sub
Private Sub PushButton7_Click()
    Dim frm1 As New frmpopupProcess
    frm1.Show vbModal
    FlatEdit25.Text = Trim(frm1.processName)
    processid2 = frm1.processid
    Unload frm1
End Sub
Private Sub PushButton6_Click()
    Dim frm1 As New frmpopupProcessment
    frm1.Show vbModal
    FlatEdit24.Text = Trim(frm1.processmentName)
    FlatEdit23.Text = frm1.phone
    processmentid2 = frm1.processmentid
    Unload frm1
End Sub
Private Sub PushButton9_Click()
    Dim frm1 As New frmpopupProcess
    frm1.Show vbModal
    FlatEdit30.Text = Trim(frm1.processName)
    processid3 = frm1.processid
    Unload frm1
End Sub
Private Sub PushButton8_Click()
    Dim frm1 As New frmpopupProcessment
    frm1.Show vbModal
    FlatEdit29.Text = Trim(frm1.processmentName)
    FlatEdit28.Text = frm1.phone
    processmentid3 = frm1.processmentid
    Unload frm1
End Sub

'---------------------------------------���ÿؼ�����ֻ���������ֺ�С����


Private Sub setValuation()
    If Len(Valuation) > 0 Then
        If Len(Trim(Valuation)) = 2 Then
            FlatEdit7.Enabled = False
            FlatEdit7.BackColor = &HC0C0C0
            num = 1
        Else
            FlatEdit6.Enabled = False
            FlatEdit6.BackColor = &HC0C0C0
              num = 2
        End If
    End If
End Sub
Private Sub save(ByVal num As Long)

     If yanzhenColor(id) = False Then
        Exit Sub
    End If
     If Trim(FlatEdit8.Text) = "" Then
        MsgBox "Ⱦ������Ϊ��", vbInformation, "��ʾ"
        Exit Sub
    End If
    If Trim(FlatEdit9.Text) = "" Then
        MsgBox "Ⱦ����������Ϊ��", vbInformation, "��ʾ"
        Exit Sub
    End If
    If Trim(FlatEdit12.Text) <= 0 Then
        MsgBox "Ⱦ���ӹ��Ѳ���Ϊ��", vbInformation, "��ʾ"
        Exit Sub
    End If

     If Trim(ComboBox5.Text) = "" Then
        MsgBox "���ȹ��ղ���Ϊ��", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If Trim(FlatEdit35.Text) = "" Then
        MsgBox "ʵ��Ͷ��������Ϊ��", vbInformation, "��ʾ"
        Exit Sub
    End If

    If num = 0 Then
             SavedetailColor
    End If
    If num = 1 Then
             Saveandcopy_1
    End If
         If num = 2 Then
             Saveandcopy
    End If
End Sub

Private Sub SavedetailColor()
     If Len(itemid) > 0 Then
        savedetail_update
    Else
        Dim rs As New RecordSet
        Dim sql As String
        '2018��4��13��19:53:30
        sql = "select * from G_BillColor where B_belongorderid='" & id & "'"
        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
'        sql = "select * from G_BilldetailColor where B_itemid='" & itemid & "'"
'        rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
      
    If rs.RecordCount <= 0 Then
           savemain
    Else
            theidColor = rs!B_id
    End If
            savedetail
    End If
    validation (id)
    Me.Hide
End Sub

Private Sub savemain()
    Set clsBL = New clsBL
    Dim sql As String
            Dim rs As New RecordSet
            sql = "select * from G_DraftBillWhite where 1=1 "
            rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            rs.AddNew
            Dim a As String
            a = Format(Now, "YYYY-MM-DD")
            rs!B_datecreate = a
            rs.Update
            theidColor = rs!B_id
            
            Dim rs1 As New RecordSet
            Dim sql1 As String
            sql1 = "select * from G_BillColor where 1=1 "
            rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            rs1.AddNew
               rs1!B_id = theidColor
               Dim b As String
               b = Format(Now, "YYYY-MM-DD")
               rs1!B_datecreate = b
               rs1!B_BID = B_BID_CC
               rs1!B_ObjectID = B_ObjectID_CC
               rs1!B_BillType = B_BillType_CC
               rs1!B_Codeid = clsBL.GetFrameCodeDetail_01(B_ObjectID_CC)
               rs1!B_BelongOrderID = id
               rs1.Update
               Dim rs2 As New RecordSet
               Dim sql2 As String
               sql2 = "delete from G_DraftBillColor where B_ID='" & theidwhite & "'"
               Gm.cnnTool.cnn.Execute sql2
     
End Sub
Private Sub savedetail()
    Dim lIncr As Long
    Dim szBC13 As String
    
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_DraftBillDetailColor where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
         '��ȡ���µ�һ���������������
    lIncr = GetNewBCIncr
    szBC13 = GetBC13(FillGetBC12(lIncr))
    rs.AddNew
    rs!B_datecreate = Now
    rs.Update
     itemid = rs!B_ItemID
     Dim rs1 As New RecordSet
     Dim sql1 As String
     sql1 = "select * from G_BillDetailColor where 1=1"
      rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
      rs1.AddNew
      rs1!B_GroupID = lGroupID
      
      rs1!B_orderitemid = IIf(IsNull(ordertocolorid), 0, ordertocolorid)
    
      rs1!B_BC13 = szBC13
      rs1!B_BCIncr = lIncr
      
      rs1!B_ItemID = itemid
      rs1!B_id = theidColor
      rs1!B_ItemIDB = FlatEdit3.Text
      rs1!B_GoodsNameAlias = FlatEdit2.Text
      rs1!B_Width = FlatEdit4.Text
      rs1!B_weight = FlatEdit10.Text
      rs1!B_color = colorid
      rs1!B_orderColor = FlatEdit11.Text
      rs1!B_Producer = FlatEdit1.Text
      rs1!B_SeHao = FlatEdit5.Text
     
      rs1!B_meter = FlatEdit6.Text
      rs1!B_kg = FlatEdit7.Text
      rs1!B_BoxQty = IIf(IsNull(FlatEdit37.Text), 0, FlatEdit37.Text)
      rs1!B_depart = IIf(IsNull(departid), "", departid)
      rs1!B_department = IIf(IsNull(departmentid), "", departmentid)

      rs1!B_departdate = Format(DTPicker1.Value, "YYYY-MM-DD")
      
      rs1!B_flowCard = FlatEdit12.Text
      rs1!B_departdannote = FlatEdit13.Text
      rs1!B_processunit = IIf(IsNull(processid), "", processid)
      rs1!B_processdocumentary = IIf(IsNull(processmentid), "", processmentid)
      
      rs1!B_phone1 = FlatEdit22.Text
      rs1!B_processdate = Format(DTPicker2.Value, "YYYY-MM-DD")
      rs1!B_processcost = FlatEdit16.Text
      rs1!B_processnote = FlatEdit17.Text
      rs1!B_processunit2 = IIf(IsNull(processid2), "", processid2)
      rs1!B_processdocumentary2 = IIf(IsNull(processmentid2), "", processmentid2)
      rs1!B_phone2 = FlatEdit23.Text
      rs1!B_processdate2 = DTPicker3.Value
      rs1!B_processCost2 = FlatEdit27.Text
      rs1!B_processnote2 = FlatEdit26.Text
      rs1!B_processunit3 = IIf(IsNull(processid3), "", processid3)
      rs1!B_processdocumentary3 = IIf(IsNull(processmentid3), "", processmentid3)
      rs1!B_processdate3 = DTPicker4.Value
      rs1!B_processCost3 = FlatEdit31.Text
      rs1!B_processnote3 = FlatEdit32.Text
      rs1!B_Progressprocess = ComboBox5.Text
      rs1!B_Paper = FlatEdit19.Text
      rs1!B_pocket = FlatEdit20.Text
      rs1!B_Empty = FlatEdit21.Text
      rs1!B_packagstyle = ComboBox4.Text
      rs1!B_departCost = FlatEdit12.Text
      rs1!B_fold = FlatEdit33.Text
      rs1!B_Cast = FlatEdit34.Text
      rs1!B_PracticeCast = FlatEdit35.Text
      rs1!B_LabelTemplate = ComboBox1.Text
      rs1!B_DetailTemplate = ComboBox2.Text
      rs1!B_DepartColor = FlatEdit36.Text
'      rs1!B_Progressprocess = ComboBox5.Text
'      rs1!B_Paper = FlatEdit19.Text
'      rs1!B_pocket = FlatEdit20.Text
'      rs1!B_Empty = FlatEdit21.Text
'      rs1!B_packagstyle = ComboBox4.Text
'      rs1!B_departCost = FlatEdit12.Text

     
      rs1.Update
      Dim sql2 As String
      sql2 = "delete from G_DraftBillDetailColor where B_itemid='" & itemid & "'"
      Gm.cnnTool.cnn.Execute sql2
End Sub

Private Sub savedetail_update()
        Dim sql As String
        Dim a As String
        Dim b As String
        Dim c As String
        Dim d As String
        a = Format(DTPicker1.Value, "YYYY-MM-dd")
        b = Format(DTPicker2.Value, "YYYY-MM-dd")
        c = Format(DTPicker3.Value, "YYYY-MM-dd")
        d = Format(DTPicker4.Value, "YYYY-MM-dd")
         sql = "exec usp_updateColordetail '" & itemid & "','" & FlatEdit3.Text & "','" & FlatEdit2.Text & "','" & FlatEdit4.Text & "','" & FlatEdit10.Text & "'"
         sql = sql & ",'" & colorid & "','" & FlatEdit1.Text & "','" & FlatEdit5.Text & "','" & FlatEdit6.Text & "','" & FlatEdit7.Text & "','" & FlatEdit37.Text & "','" & departid & "','" & departmentid & "','" & FlatEdit18.Text & "'"
         sql = sql & ",'" & a & "','" & FlatEdit13.Text & "','" & processid & "','" & processmentid & "','" & FlatEdit22.Text & "','" & b & "'"
         sql = sql & ",'" & FlatEdit16.Text & "','" & FlatEdit17.Text & "'"
         sql = sql & ",'" & processid2 & "','" & processmentid2 & "','" & FlatEdit23.Text & "','" & c & "','" & FlatEdit27.Text & "','" & FlatEdit26.Text & "'"
         sql = sql & ",'" & processid3 & "','" & processmentid3 & "','" & FlatEdit28.Text & "','" & d & "','" & FlatEdit31.Text & "','" & FlatEdit32.Text & "'"
         sql = sql & ",'" & ComboBox5.Text & "','" & FlatEdit19.Text & "','" & FlatEdit20.Text & "','" & FlatEdit21.Text & "','" & ComboBox4.Text & "','" & FlatEdit12.Text & "','" & FlatEdit33.Text & "','" & FlatEdit34.Text & "','" & FlatEdit35.Text & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & FlatEdit36.Text & "'"
         sql = sql & ",'','','','','','','',''"
         Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
        
End Sub

Private Sub Saveandcopy()
        Dim sql As String
        Dim a As String
        Dim b As String
        Dim c As String
        Dim d As String
        a = Format(DTPicker1.Value, "YYYY-MM-dd")
        b = Format(DTPicker2.Value, "YYYY-MM-dd")
        c = Format(DTPicker3.Value, "YYYY-MM-dd")
        d = Format(DTPicker3.Value, "YYYY-MM-dd")
         sql = "exec usp_saveandCopy '" & id & "','" & FlatEdit2.Text & "','" & FlatEdit4.Text & "','" & FlatEdit10.Text & "','" & FlatEdit1.Text & "','" & FlatEdit5.Text & "'"
         sql = sql & ",'" & departid & "','" & departmentid & "','" & FlatEdit18.Text & "'"
         sql = sql & ",'" & a & "','" & FlatEdit13.Text & "','" & processid & "','" & processmentid & "','" & FlatEdit22.Text & "','" & b & "'"
         sql = sql & ",'" & FlatEdit16.Text & "','" & FlatEdit17.Text & "'"
         sql = sql & ",'" & processid2 & "','" & processmentid2 & "','" & FlatEdit23.Text & "','" & c & "','" & FlatEdit27.Text & "','" & FlatEdit26.Text & "'"
         sql = sql & ",'" & processid3 & "','" & processmentid3 & "','" & FlatEdit28.Text & "','" & d & "','" & FlatEdit31.Text & "','" & FlatEdit32.Text & "'"
         sql = sql & ",'" & ComboBox5.Text & "','" & FlatEdit19.Text & "','" & FlatEdit20.Text & "','" & FlatEdit21.Text & "','" & ComboBox4.Text & "','" & FlatEdit12.Text & "','" & FlatEdit33.Text & "','" & FlatEdit34.Text & "','" & FlatEdit35.Text & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & FlatEdit3.Text & "','" & FlatEdit36.Text & "'"
         sql = sql & ",'','','','',0,0,0,0"
         Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
        MsgBox "����ɹ������Ƴɹ�", vbInformation, "��ʾ"
        validation (id)
        Me.Hide
End Sub
Private Sub Saveandcopy_1()
          Dim sql As String
        Dim a As String
        Dim b As String
         Dim c As String
        Dim d As String
        a = Format(DTPicker1.Value, "YYYY-MM-dd")
        b = Format(DTPicker2.Value, "YYYY-MM-dd")
         c = Format(DTPicker3.Value, "YYYY-MM-dd")
        d = Format(DTPicker3.Value, "YYYY-MM-dd")
         sql = "exec usp_saveandCopy '" & id & "','" & FlatEdit2.Text & "','" & FlatEdit4.Text & "','" & FlatEdit10.Text & "','" & FlatEdit1.Text & "','" & FlatEdit5.Text & "'"
         sql = sql & ",'" & departid & "','" & departmentid & "','" & FlatEdit18.Text & "'"
         sql = sql & ",'" & a & "','" & FlatEdit13.Text & "','" & processid & "','" & processmentid & "','" & FlatEdit22.Text & "','" & b & "'"
         sql = sql & ",'" & FlatEdit16.Text & "','" & FlatEdit17.Text & "'"
         sql = sql & ",'" & processid2 & "','" & processmentid2 & "','" & FlatEdit23.Text & "','" & c & "','" & FlatEdit27.Text & "','" & FlatEdit26.Text & "'"
         sql = sql & ",'" & processid3 & "','" & processmentid3 & "','" & FlatEdit28.Text & "','" & d & "','" & FlatEdit31.Text & "','" & FlatEdit32.Text & "'"
         sql = sql & ",'" & ComboBox5.Text & "','" & FlatEdit19.Text & "','" & FlatEdit20.Text & "','" & FlatEdit21.Text & "','" & ComboBox4.Text & "','" & FlatEdit12.Text & "','" & FlatEdit33.Text & "','" & FlatEdit34.Text & "','" & FlatEdit35.Text & "','" & ComboBox1.Text & "','" & ComboBox2.Text & "','" & "" & "','" & FlatEdit36.Text & "'"
         sql = sql & ",'','','','',0,0,0,0"
         Debug.Print sql
        Gm.cnnTool.cnn.Execute sql
        MsgBox "����ɹ������Ƴɹ�", vbInformation, "��ʾ"
        validation (id)
        Me.Hide
End Sub

Private Sub movefrist()
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select top 1 * from G_BilldetailColor where B_ID=(select B_ID from G_BillColor where B_BelongOrderID='" & id & "') "
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
     If rs.RecordCount <= 0 Then
            MsgBox "��ǰû������", vbInformation, "��ʾ"
     Else
            Label26.Caption = 1
            itemid = rs!B_ItemID
            openbill
     End If
End Sub

Private Sub moveshang()
     Dim rs As New RecordSet
     Dim sql As String
     sql = "select top 1 * from G_BilldetailColor where B_ID=(select B_ID from G_BillColor where B_BelongOrderID='" & id & "') and B_itemid<'" & itemid & "'  Order by B_itemid desc"
      rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
      
      If rs.RecordCount <= 0 Then
        MsgBox "�Ѿ��ǵ�һ����", vbOKOnly + vbInformation, "��ʾ"
     Else
        Label26.Caption = Val(Label26.Caption) - 1
        itemid = rs!B_ItemID
        openbill
     End If
End Sub

Private Sub movenext()
     Dim rs As New RecordSet
     Dim sql As String
     sql = "select top 1 * from G_BilldetailColor where B_ID=(select B_ID from G_BillColor where B_BelongOrderID='" & id & "') and B_itemid>'" & itemid & "'  Order by B_itemid asc"
      rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
      
      If rs.RecordCount <= 0 Then
        MsgBox "���һ����", vbOKOnly + vbInformation, "��ʾ"
     Else
        Label26.Caption = Val(Label26.Caption) + 1
        itemid = rs!B_ItemID
        openbill
     End If
End Sub

Private Sub movelast()
     Dim rs As New RecordSet
     Dim sql As String
     sql = "select top 1 * from G_BilldetailColor where B_ID=(select B_ID from G_BillColor where B_BelongOrderID='" & id & "')   Order by B_itemid desc"
      rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
      
      If rs.RecordCount <= 0 Then
        MsgBox "��ǰû���κ�����", vbOKOnly + vbInformation, "��ʾ"
     Else
         Label26.Caption = a
       itemid = rs!B_ItemID
        openbill
     End If
End Sub

Private Sub openbill()
     Dim sql As String
    Dim rs As New RecordSet
    sql = "select *from G_Billdetailcolor where B_itemid='" & itemid & "' "
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    FlatEdit3.Text = IIf(IsNull(rs!B_ItemIDB), "", rs!B_ItemIDB)
    FlatEdit2.Text = IIf(IsNull(rs!B_GoodsNameAlias), "", rs!B_GoodsNameAlias)
    FlatEdit4.Text = IIf(IsNull(rs!B_Width), "", rs!B_Width)
    FlatEdit10.Text = IIf(IsNull(rs!B_weight), "", rs!B_weight)
    colorid = IIf(IsNull(rs!B_color), "", rs!B_color)
'    FlatEdit11.Text = GetColorName(rs!B_Color)
    FlatEdit11.Text = IIf(IsNull(rs!B_orderColor), "", rs!B_orderColor)
    FlatEdit1.Text = IIf(IsNull(rs!B_Producer), "", rs!B_Producer)
    FlatEdit5.Text = IIf(IsNull(rs!B_SeHao), "", rs!B_SeHao)
    FlatEdit37.Text = IIf(IsNull(rs!B_BoxQty), "", rs!B_BoxQty)
'    If Trim(rs!B_Meter) > 0 Then
        FlatEdit6.Text = rs!B_meter
'        FlatEdit6.Enabled = True
'         FlatEdit7.Enabled = False
'         FlatEdit7.BackColor = &HC0C0C0
'    End If
'    If Trim(rs!B_KG) > 0 Then
        FlatEdit7.Text = rs!B_kg
'        FlatEdit6.Enabled = False
'         FlatEdit7.Enabled = True
'         FlatEdit6.BackColor = &HC0C0C0
'    End If
'    frm1.FlatEdit6.Text = IIf(IsNull(rs!B_meter), "", rs!B_meter)
'    frm1.FlatEdit7.Text = IIf(IsNull(rs!B_KG), "", rs!B_KG)
    FlatEdit8.Text = getDepartName(IIf(IsNull(rs!B_depart), "", rs!B_depart))
    departid = IIf(IsNull(rs!B_depart), "", rs!B_depart)
    FlatEdit9.Text = getprocessName(IIf(IsNull(rs!B_department), "", rs!B_department))
    departmentid = IIf(IsNull(rs!B_department), "", rs!B_department)
    If IIf(IsNull(rs!B_departdate), "", rs!B_departdate) = "" Then
        DTPicker1.Value = Now
    Else
        DTPicker1.Value = rs!B_departdate
    End If
'    frm1.DTPicker1.Value = IIf(IsNull(rsdetailColor!B_departdate), "", rsdetailColor!B_departdate)
'    frm1.FlatEdit12.Text = IIf(IsNull(rs!B_flowCard), "", rs!B_flowCard)
    FlatEdit13.Text = IIf(IsNull(rs!B_departdannote), "", rs!B_departdannote)
    FlatEdit14.Text = getDepartName(IIf(IsNull(rs!B_processunit), "", rs!B_processunit))
    processid = IIf(IsNull(rs!B_processunit), "", rs!B_processunit)
    FlatEdit15.Text = getprocessName(IIf(IsNull(rs!B_processdocumentary), "", rs!B_processdocumentary))
    processmentid = IIf(IsNull(rs!B_processdocumentary), "", rs!B_processdocumentary)
   If IIf(IsNull(rs!B_processdate), "", rs!B_processdate) = "" Then
        DTPicker2.Value = Now
    Else
        DTPicker2.Value = rs!B_processdate
    End If
'    frm1.DTPicker2.Value = IIf(IsNull(rs!B_processdate), "", rs!B_processdate)
    FlatEdit17.Text = IIf(IsNull(rs!B_processnote), "", rs!B_processnote)
     If IIf(IsNull(rs!B_Progressprocess), "", rs!B_Progressprocess) = "" Then
            ComboBox5.Text = ""
     Else
        ComboBox5.Text = GetProgressCraftCT(IIf(IsNull(rs!B_Progressprocess), "", rs!B_Progressprocess))
     End If
'    frm1.ComboBox5.Text = GetProgressCraftCT(rs!B_Progressprocess)
    FlatEdit19.Text = IIf(IsNull(rs!B_Paper), "", rs!B_Paper)
    FlatEdit20.Text = IIf(IsNull(rs!B_pocket), "", rs!B_pocket)
    FlatEdit21.Text = IIf(IsNull(rs!B_Empty), "", rs!B_processnote)
    If IIf(IsNull(rs!B_packagstyle), "", rs!B_packagstyle) = "" Then
            ComboBox4.Text = ""
     Else
        ComboBox4.Text = GetB_packagstyle(rs!B_packagstyle)
     End If
End Sub

Private Function GetColorName(ByVal da As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select B_Name from G_Color where B_SID='" & da & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        GetColorName = rs!B_name
    Else
        GetColorName = ""
    End If
End Function

Private Function getDepartName(ByVal da As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select B_ClientName from G_ContactCompany where B_Clientid='" & da & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        getDepartName = rs!B_ClientName
    Else
        getDepartName = ""
    End If
End Function
Private Function getprocessName(ByVal da As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select B_Name from G_Employee where B_SID='" & da & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        getprocessName = rs!B_name
    Else
        getprocessName = ""
    End If
End Function

Private Function GetProgressCraftCT(ByVal da As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select B_SID from G_ProgressCraftCT where B_SID='" & da & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        GetProgressCraftCT = rs!B_sid
    Else
        GetProgressCraftCT = ""
    End If
End Function
Private Function GetB_packagstyle(ByVal da As String) As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "Select B_SID From G_PackWay where B_SID='" & da & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        GetB_packagstyle = rs!B_sid
    Else
        GetB_packagstyle = ""
    End If
End Function

Private Sub FlatEdit19_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit20_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit21_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit12_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit16_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit33_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit35_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit19_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            FlatEdit20.SetFocus
    End Select
End Sub
Private Sub FlatEdit20_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            FlatEdit21.SetFocus
    End Select
End Sub
Private Sub FlatEdit21_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            ComboBox4.SetFocus
    End Select
End Sub
Private Sub FlatEdit33_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error Resume Next
    Select Case KeyCode
        Case 13
            FlatEdit35.SetFocus
    End Select
End Sub
Private Sub validation(ByVal id As String)
    Dim sql As String
    Dim rs As New RecordSet
    sql = "select * from G_BillDetailColor where B_ID in (select B_ID from G_BillColor where B_BelongOrderid='" & id & "')"
    rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    Do While Not rs.EOF
        If IIf(IsNull(rs!B_Date), "", rs!B_Date) <> "" Then
                Exit Sub
        End If
        rs.movenext
    Loop
    rs.MoveFirst
    Do While Not rs.EOF
        If IIf(IsNull(rs!B_departdate), "", rs!B_departdate) = "" Then
            Exit Sub
        End If
        rs.movenext
    Loop
    rs.MoveFirst
    Do While Not rs.EOF
        rs!B_Date = Now
        rs.movenext
    Loop
End Sub
Public Function yanzhenColor(ByVal theid As String) As Boolean
    Dim sql As String
    Dim rs As New RecordSet
    Dim sql1 As String
    Dim rs1 As New RecordSet
    Dim sql2 As String
    Dim rs2 As New RecordSet
       Dim sql3 As String
    Dim rs3 As New RecordSet
    sql = "Select * From G_SystemUser where B_UserName='" & Gm.SysID.SystemUser & "'"
     rs.Open sql, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
    If rs!B_SuperAdmin = 1 Then
        yanzhenColor = True
        Exit Function
    End If
    
    sql1 = "select distinct B_date from G_BilldetailColor where B_ID=(select B_ID from G_BillColor where B_belongorderid='" & theid & "')"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
'    sql2 = "select * from G_BillColor where B_Belongorderid='" & theid & "'"
 sql2 = "SELECT * FROM G_UserPro WHERE B_username='" & Gm.SysID.SystemUser & "' AND B_objectid='11S005'"
    rs2.Open sql2, Gm.cnnTool.cnn, adOpenStatic, adLockPessimistic
   If rs2.RecordCount > 0 Then
        If IIf(IsNull(rs2!B_new), 0, rs2!B_new) = 1 Then
            yanzhenColor = True
        Else
            yanzhenColor = False
            MsgBox "������Ȩ��", vbInformation, "��ʾ"
            Exit Function
        End If
        If IIf(IsNull(rs1!B_Date), "", rs1!B_Date) <> "" Then
        
             sql3 = "SELECT B_value FROM G_Config_OneInt WHERE B_groupname='֯��ϵͳ_��ͬ�����޸�ʱ��'"
            rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
            
            If DateDiff("s", rs1!B_Date, Now) > rs3!B_Value Then
                yanzhenColor = False
                MsgBox "�Ѿ��������������ݵ�ʱ�䲻�ܽ����޸�", vbInformation, "��ʾ"
            Else
                yanzhenColor = True
            End If
        End If
    Else
        yanzhenColor = False
        MsgBox "��û�д�Ȩ��", vbInformation, "��ʾ"
    End If
End Function


Private Sub ActiveBar23_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "����"
            savecg
        Case "�˳�"
            Unload Me
        Case "����"
            AddNew
        Case "ɾ��"
            de
    End Select
End Sub


Private Sub PushButton10_Click()
    Dim frm1 As New frmPopupClient_Edit
     frm1.a = "ɫ����Ӧ��"
     frm1.Show vbModal
    client = frm1.clientid
    FlatEdit45.Text = frm1.ClientName
    Unload frm1
End Sub

Private Sub Grid()
    On Error Resume Next
    
    Set rss = New RecordSet
    Dim sql1 As String
    Dim rs As New RecordSet
    Dim sql As String
    sql = "select * from G_BilldetailColor where B_itemid='" & colororderid & "' "
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
   
'    sql1 = "select * from G_BilldetailColor where B_orderitemid='" & rs!B_orderitemid & "' and isnull(B_intype,0)=1"
'    rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    sql1 = "exec usp_SelectColororder '" & rs!B_orderitemid & "' "
    rss.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    TDBGrid2.DataSource = rss
    If rss.RecordCount > 0 Then
        rss.MoveFirst
    End If
    setgrid
    
    
    
    
    
End Sub
Private Sub setgrid()
    TDBGrid2.Columns("B_ItemIDB").Caption = "������"
    TDBGrid2.Columns("B_ClientName").Caption = "�ͻ�"
    TDBGrid2.Columns("B_Width").Caption = "�ŷ�"
    TDBGrid2.Columns("B_weight").Caption = "����"
    TDBGrid2.Columns("B_GoodsNameAlias").Caption = "Ʒ��"
    TDBGrid2.Columns("B_Name").Caption = "��ɫ"
    TDBGrid2.Columns("B_SeHao").Caption = "ɫ��"
    TDBGrid2.Columns("B_ps").Caption = "ƥ��"
    TDBGrid2.Columns("B_kg").Caption = "����"
    TDBGrid2.Columns("B_meter").Caption = "����"
    TDBGrid2.Columns("B_qty").Caption = "����"
    TDBGrid2.Columns("B_departdate").Caption = "����"
    TDBGrid2.Columns("B_MemoDetail").Caption = "��ע"
    TDBGrid2.Columns("B_hex").Caption = "ɫ��"
    TDBGrid2.Columns("B_price").Caption = "����"
    TDBGrid2.Columns("B_ItemIDB").width = 900
    TDBGrid2.Columns("B_Width").width = 900
    TDBGrid2.Columns("B_weight").width = 900
    TDBGrid2.Columns("B_SeHao").width = 900
    TDBGrid2.Columns("B_ps").width = 900
    TDBGrid2.Columns("B_kg").width = 900
    TDBGrid2.Columns("B_qty").width = 900
    TDBGrid2.Columns("B_meter").width = 900
    TDBGrid2.Columns("B_hex").width = 900
    TDBGrid2.Columns("B_price").width = 900
    
    TDBGrid2.Columns("B_Clientid").Visible = False
    TDBGrid2.Columns("B_Clientid").AllowSizing = False
    TDBGrid2.Columns("B_Clientid").Locked = True
        TDBGrid2.Columns("B_color").Visible = False
    TDBGrid2.Columns("B_color").AllowSizing = False
    TDBGrid2.Columns("B_color").Locked = True
      TDBGrid2.Columns("B_itemid").Visible = False
    TDBGrid2.Columns("B_itemid").AllowSizing = False
    TDBGrid2.Columns("B_itemid").Locked = True
    
    TDBGrid2.MarqueeStyle = dbgHighlightRow
    TDBGrid2.HoldFields
End Sub


Private Sub savecg()
    If Trim(client) = "" Then
        MsgBox "�ͻ�����Ϊ��", vbInformation, "��ʾ"
        Exit Sub
    End If
    If Trim(FlatEdit41.Text) = "" Then
        MsgBox "�ŷ�����Ϊ��", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If Trim(FlatEdit42.Text) = "" Then
        MsgBox "���ز���Ϊ��", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If Trim(FlatEdit40.Text) = "" Then
        MsgBox "Ʒ������Ϊ��", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If Trim(colorid) = "" And Trim(FlatEdit43.Text) = "" Then
        MsgBox "��ɫ����ɫ�Ų���Ϊ��", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If Trim(FlatEdit49.Text) = "" And Trim(FlatEdit44.Text) = "" And Trim(FlatEdit46.Text) = "" And Trim(FlatEdit38.Text) = "" Then
        MsgBox "ƥ��,������,����,����,��д��һ", vbInformation, "��ʾ"
        Exit Sub
    End If
    If Trim(FlatEdit50.Text) = "" Then
        MsgBox "���۲���Ϊ��", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    If B_orderitemid = "" Then
        saveALL
    Else
        saveupdate
    End If
End Sub
Private Sub saveALL()
    Dim idd As String
    Dim rs As New RecordSet
    Dim sql As String
    Dim rs1 As New RecordSet
    Dim sql1 As String
    Dim sql2 As String
    Dim rs3 As New RecordSet
    Dim sql3 As String
    
    sql3 = "select * from G_BilldetailColor where B_itemid='" & colororderid & "' "
    rs3.Open sql3, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    
    sql = "select * from G_draftBilldetailColor where 1=1"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs.AddNew
    rs!B_datecreate = Now
    rs.Update
    idd = rs!B_ItemID
    
    sql1 = "select * from G_BilldetailColor where 1=1"
    rs1.Open sql1, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
    rs1.AddNew
    rs1!B_ItemID = idd
    rs1!B_ItemIDB = FlatEdit39.Text
    rs1!B_Clientid = client
    rs1!B_Width = FlatEdit41.Text
    rs1!B_weight = FlatEdit42.Text
    rs1!B_GoodsNameAlias = FlatEdit40.Text
    rs1!B_color = colorplanid
    rs1!B_orderColor = FlatEdit48.Text
    rs1!B_SeHao = FlatEdit43.Text
    rs1!B_price = FlatEdit50.Text
    rs1!B_ps = FlatEdit49.Text
    rs1!B_kg = FlatEdit44.Text
    If Len(Trim(FlatEdit46.Text)) > 0 Then
        rs1!B_meter = IIf(IsNull(FlatEdit46.Text), 0, FlatEdit46.Text)
    End If
    If Len(Trim(FlatEdit38.Text)) > 0 Then
        rs1!B_qty = IIf(IsNull(FlatEdit38.Text), 0, FlatEdit38.Text)
    End If
    rs1!B_departdate = DateTimePicker1.Value
    rs1!B_MemoDetail = FlatEdit47.Text
    rs1!B_intype = 1
    rs1!B_orderitemid = rs3!B_orderitemid
    rs1.Update
    
    sql2 = "delete from G_draftBilldetailColor where B_itemid='" & idd & "'"
    Gm.cnnTool.cnn.Execute sql2
    MsgBox "����ɹ�", vbInformation, "��ʾ"
    rss.requery
End Sub

Private Sub PushButton11_Click()
'    Dim frm1 As New frmpopupColor
'    frm1.Show vbModal
'    If frm1.bsaved = True Then
'        colorplanid = frm1.colorid
'        FlatEdit48.Text = frm1.colorname
'    End If
    On Error Resume Next
    Dim sql As String
    Dim rs As New RecordSet
    Dim frm1 As New frmpopupColor
    frm1.Show vbModal
    FlatEdit48.Text = Trim(frm1.colorname)
    colorplanid = frm1.colorid
    If frm1.bsaved = True Then
        FlatEdit48.Enabled = True
    Else
        If colorplanid = "" Then
             FlatEdit48.Enabled = False
        End If
    End If
    
        
    Unload frm1
End Sub

Private Sub saveupdate()
    Dim sql As String
    Dim rs As New RecordSet
    sql = "exec usp_OrderColor '" & B_orderitemid & "','" & FlatEdit39.Text & "','" & client & "','" & FlatEdit41.Text & "','" & FlatEdit42.Text & "','" & FlatEdit40.Text & "',"
    sql = sql & "'" & colorid & "','" & FlatEdit43.Text & "','" & FlatEdit49.Text & "','" & FlatEdit44.Text & "',"
    sql = sql & "'" & FlatEdit46.Text & "','" & FlatEdit38.Text & "','" & DateTimePicker1.Value & "','" & FlatEdit47.Text & "','" & FlatEdit48.Text & "','" & FlatEdit50.Text & "'"
    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset
    MsgBox "�������", vbInformation, "��ʾ"
    rss.requery
End Sub

'Private Function clientname(ByVal client As String)
'    Dim rs As New RecordSet
'    Dim sql As String
'    sql = "select * from G_ContactCompany where B_clientid='" & client & "'"
'    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
'    If rs.RecordCount > 0 Then
'        clientname = rs!B_ClientName
'    Else
'        clientname = ""
'    End If
'End Function
'
'Private Function colorname(ByVal client As String)
'    Dim rs As New RecordSet
'    Dim sql As String
'    sql = "select * from G_Color where B_sid='" & client & "'"
'    rs.Open sql, Gm.cnnTool.cnn, adOpenKeyset, adLockPessimistic
'    If rs.RecordCount > 0 Then
'        colorname = rs!B_name
'    Else
'        colorname = ""
'    End If
'End Function

Private Sub FlatEdit49_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub

Private Sub FlatEdit44_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit46_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub

Private Sub FlatEdit38_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub
Private Sub FlatEdit50_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then
        Exit Sub
     End If
     If KeyAscii = 8 Then
        Exit Sub
     End If
     If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
     End If
End Sub

Private Sub TDBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If
    
    B_orderitemid = rss!B_ItemID
    FlatEdit39.Text = rss!B_ItemIDB
    client = rss!B_Clientid
    FlatEdit45.Text = rss!B_ClientName
    FlatEdit41.Text = rss!B_Width
    FlatEdit42.Text = rss!B_weight
    FlatEdit40.Text = rss!B_GoodsNameAlias
    colorplanid = rss!B_color
    FlatEdit48.Text = rss!B_name
    FlatEdit43.Text = rss!B_SeHao
    FlatEdit49.Text = rss!B_ps
    FlatEdit44.Text = rss!B_kg
    FlatEdit46.Text = IIf(IsNull(rss!B_meter), "", rss!B_meter)
    FlatEdit38.Text = IIf(IsNull(rss!B_qty), "", rss!B_qty)
    DateTimePicker1.Value = rss!B_departdate
    FlatEdit47.Text = rss!B_MemoDetail
    FlatEdit50.Text = rss!B_price
End Sub

Private Sub AddNew()
    Dim a As String
    a = FlatEdit39.Text
    B_orderitemid = ""
    colorplanid = ""
    client = ""
    
     FlatEdit39.Text = ""
    FlatEdit45.Text = ""
    FlatEdit41.Text = ""
    FlatEdit42.Text = ""
    FlatEdit40.Text = ""
    FlatEdit48.Text = ""
    FlatEdit43.Text = ""
    FlatEdit49.Text = ""
    FlatEdit44.Text = ""
    FlatEdit46.Text = ""
    FlatEdit38.Text = ""
    FlatEdit47.Text = ""
    FlatEdit50.Text = ""
    FlatEdit39.Text = a
End Sub

Private Sub de()
    If TDBGrid2.ApproxCount <= 0 Then
        Exit Sub
    End If
    Dim sql As String
    Dim rs As New RecordSet
    sql = "delete from G_BilldetailColor where B_itemid='" & rss!B_ItemID & "'" '
    Gm.cnnTool.cnn.Execute sql
    rss.requery
End Sub

'�ӱ�G_BillDetailColor��ȡ��ǰ����һ���������������
Private Function GetNewBCIncr() As Long
    Dim rs As New RecordSet
    strSQL = "select top 1 * from G_BillDetailColor order by B_BCIncr desc"
    Debug.Print strSQL
    rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    Dim lRturn As Long
    If rs.RecordCount <= 0 Then
        lRturn = 1
    Else
        lRturn = IIf(IsNull(rs!B_BCIncr), 0, rs!B_BCIncr) + 1
    End If
    
    rs.Close
    Set rs = Nothing
    
    GetNewBCIncr = lRturn
End Function
Private Function GetBC13(ByVal vBC12 As String) As String
    Dim szRturn As String
    szRturn = GetEAN13CheckOut(vBC12)
    
    GetBC13 = vBC12 & szRturn
End Function

'��ȡ���µ�һ��13λ����
Private Function GetBC13Ex() As String
    Dim szIncr As String
    szIncr = GetNewBCIncr
    
    Dim szBC12 As String
    szBC12 = FillGetBC12(GetNewBCIncr)
    
    GetBC13Ex = GetBC13(szBC12)
End Function
'������������ⳤ�ȵ��������ֵ��ַ�������
'����ֵ������BC13�����ǰ��12λ�ַ�
Private Function FillGetBC12(ByVal vIncr As String) As String
    Dim cls1 As New clsString
    Dim szReturn As String
    
    szReturn = cls1.FillRepeat(vIncr, 11, "0", True)
    szReturn = COLORBC13FIRST & szReturn
    
    FillGetBC12 = szReturn
End Function
