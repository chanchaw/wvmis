VERSION 5.00
Object = "{5404359C-E9EA-4988-8878-9A3A03D932FC}#3.0#0"; "ccCtlButton.ocx"
Begin VB.Form frmSetGridRowsCols 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ȡ��������"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetGridRowsCols.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8190
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1920
      TabIndex        =   12
      Text            =   "0"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1920
      TabIndex        =   11
      Text            =   "1"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   7935
   End
   Begin ccCtlButton.ccButton ccButton1 
      Height          =   495
      Index           =   0
      Left            =   4800
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "��ʼ��ȡ"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   6360
      TabIndex        =   5
      Text            =   "4"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   6360
      TabIndex        =   4
      Text            =   "10"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmSetGridRowsCols.frx":038A
      Top             =   240
      Width           =   7935
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   7935
   End
   Begin ccCtlButton.ccButton ccButton1 
      Height          =   495
      Index           =   1
      Left            =   6480
      TabIndex        =   8
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "ȡ���˳�"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ų�ǰj�У�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ų�ǰi�У�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ȡ��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4320
      TabIndex        =   3
      Top             =   2880
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ȡ��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   4320
      TabIndex        =   2
      Top             =   2160
      Width           =   1890
   End
End
Attribute VB_Name = "frmSetGridRowsCols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public w_Rows As Long
Public w_Cols As Long

Public w_ExcludeRows As Long  '�ų�������
Public w_ExcludeCols As Long  '�ų�������

Public W_ExcelImportDefaultGroupName As String   '��G_ExcelImportDefault�е�B_GroupName

Private strSQL As String
Private A_rs As RecordSet

Private Sub Save()
    w_Rows = Val(Trim(Text2(0).Text))
    w_Cols = Val(Trim(Text2(1).Text))
    
    w_ExcludeRows = Val(Trim$(Text2(2).Text))
    w_ExcludeCols = Val(Trim$(Text2(3).Text))
    
    Me.Hide
End Sub

Private Sub ccButton1_Click(Index As Integer)
    Select Case Index
        Case 0
            Save
        Case 1
            Me.Hide
    End Select
End Sub

Private Sub OpenBill()
    Dim clsDataBase1 As clsDataBase
    
    'Ϊ�˼����ϰ汾�����
    '�ϰ汾��û��UI��4���������趨Ĭ��ֵ�����ñ�G_ExcelImportDefault��
    '���Ե��ϰ汾ִ�е������ʱ���ж����û���趨W_ExcelImportDefaultGroupName���˳�
    '�����ʢ����ϵͳ��ʼ�и����ñ���,��Ҫͨ���ñ��ȡ���õ�Ĭ��ֵ
    If Len(W_ExcelImportDefaultGroupName) <= 0 Then
        Exit Sub
    End If
    
    
    Set clsDataBase1 = New clsDataBase
    clsDataBase1.initCls Gm.SysID.DBInfo.DBName
    If clsDataBase1.JudgeTableExist("G_ExcelImportDefault") = False Then
        Exit Sub
    End If
    
    If Len(Trim(W_ExcelImportDefaultGroupName)) <= 0 Then
        Exit Sub
    End If
    
    Set A_rs = New RecordSet
    strSQL = "Select * From G_ExcelImportDefault Where B_GroupName='" & W_ExcelImportDefaultGroupName & "'"
    A_rs.Open strSQL, Gm.cnnTool.cnn, adOpenStatic, adLockReadOnly
    
    
    If A_rs.RecordCount <= 0 Then
        A_rs.Close
        Set A_rs = Nothing
        Exit Sub
    End If
    
    Text2(2).Text = IIf(IsNull(A_rs!B_RowCountExclude), 0, A_rs!B_RowCountExclude)
    Text2(3).Text = IIf(IsNull(A_rs!B_ColCountExclude), 0, A_rs!B_ColCountExclude)
    Text2(0).Text = IIf(IsNull(A_rs!B_RowCountInClude), 0, A_rs!B_RowCountInClude)
    Text2(1).Text = IIf(IsNull(A_rs!B_ColCountInClude), 0, A_rs!B_ColCountInClude)
    
    A_rs.Close
    Set A_rs = Nothing
End Sub

Private Sub Form_Load()
    OpenBill
End Sub
