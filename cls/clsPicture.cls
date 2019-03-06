VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'�����������ڻ�ͼ���Ŵ���С�Լ���ת
'Ĭ���ڿؼ��������뻭ͼ
Private A_FileName As String
Private A_stdPicture As StdPicture
Private A_ZoomedPicture As StdPicture
Private A_PictureBox As PictureBox
Private A_Rate As Double

Private A_StartX As Double  '��PictureBox��ʼ������ʼλ�õ�X����
Private A_StartY As Double  '��PictureBox��ʼ������ʼλ�õ�Y����

Private A_ZoomRate As Double 'ÿ�����ŵı���.��ʼ��Ϊ0.01��ÿ�����ŵı�����1%

Private Type PictureSize
    width As Single
    height As Single
End Type

Private Type BitMap
   bmType As Long         'ͼ�����ͣ�0 ��ʾ��λͼ
   bmWidth As Long        'ͼ����ȣ����أ�
   bmHeight As Long       'ͼ��߶ȣ����أ�
   bmWidthBytes As Long   'ÿһ��ͼ����ֽ���
   bmPlanes As Integer    'ͼ���ͼ����
   bmBitsPixel As Integer 'ͼ���λ��
   bmBits As Long         'λͼ���ڴ�ָ��
End Type

Private A_PictureSize As PictureSize
Private ctP180 As Double

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long


Public Property Let ZoomRate(ByVal vData As Double)
    A_ZoomRate = vData
End Property

Public Property Get ZoomRate() As Double
    ZoomRate = A_ZoomRate
End Property

Public Property Let StartX(ByVal vData As Double)
    A_StartX = vData
End Property

Public Property Get StartX() As Double
    StartX = A_StartX
End Property

Public Property Let StartY(ByVal vData As Double)
    A_StartY = vData
End Property

Public Property Get StartY() As Double
    StartY = A_StartY
End Property

'��ʼ���Ż�����ʼ���꣬Ĭ���ڿؼ����м俪ʼ��
Private Sub InitStartPosition(ByVal vRate As Double)
    A_StartX = A_PictureBox.width / 2 - A_stdPicture.width * vRate / 2
    
    If A_StartX < 0 Then
        A_StartX = 0
    End If
    A_StartY = 0
End Sub

Public Sub InitCls(ByVal vFileName As String, ByRef vPictureBox As PictureBox, _
    Optional ByVal vRate As Double = 1)
    On Error Resume Next
    ctP180 = 4 * Atn(1) 'Բ����
    
    A_FileName = vFileName
    Set A_stdPicture = LoadPicture(vFileName)
    Set A_PictureBox = vPictureBox
    A_Rate = vRate
    
    
    '����ʱ��û��������ʾ����
    '��ô�Զ�������
    '�Զ�����=�ؼ��Ŀ��� / ͼƬ�Ŀ���
    If A_Rate = 1 Then
        A_Rate = GetZoomRate
    End If


    A_ZoomRate = 0.01

    '��ʼ���Ż��������
    InitStartPosition A_Rate


    '��ͼ
    PaintPicture A_Rate
End Sub
Public Sub InitCls1(ByVal vFileName As String, ByRef vPictureBox As Image, _
    Optional ByVal vRate As Double = 1)
    On Error Resume Next
    ctP180 = 4 * Atn(1) 'Բ����
    
    A_FileName = vFileName
    Set A_stdPicture = LoadPicture(vFileName)
    Set A_PictureBox = vPictureBox
    A_Rate = vRate
    
    
    '����ʱ��û��������ʾ����
    '��ô�Զ�������
    '�Զ�����=�ؼ��Ŀ��� / ͼƬ�Ŀ���
    If A_Rate = 1 Then
        A_Rate = GetZoomRate
    End If


    A_ZoomRate = 0.01

    '��ʼ���Ż��������
    InitStartPosition A_Rate


    '��ͼ
    PaintPicture A_Rate
End Sub
'����ͼƬ�ĳߴ����ʾͼƬ�Ŀؼ��ĳߴ�
'��ȡ����  �߶ȵı��ʣ�����С�ı�������
'������ʾͼƬ�ı���
Private Function GetZoomRate() As Double
    Dim rateWidth As Double
    Dim rateHeight As Double
    
    rateWidth = A_PictureBox.width / A_stdPicture.width
    rateHeight = A_PictureBox.height / A_stdPicture.height
    
    If rateWidth < rateHeight Then
        GetZoomRate = rateWidth
    Else
        GetZoomRate = rateHeight
    End If
End Function

Public Sub Dispose()
    Set A_stdPicture = Nothing
End Sub

Private Sub PaintPicture(ByVal vRate As Double)
    A_PictureBox.Cls
    A_PictureBox.PaintPicture A_stdPicture, A_StartX, A_StartY, _
        A_stdPicture.width * vRate, A_stdPicture.height * vRate

'    A_PictureBox.PaintPicture LoadPicture(A_FileName), A_StartX, A_StartY, _
'        A_stdPicture.width * vRate, A_stdPicture.height * vRate

    '��¼��ǰ��ͼƬ�ߴ�
    A_PictureSize.width = A_stdPicture.width * vRate
    A_PictureSize.height = A_stdPicture.height * vRate
    
    Set A_ZoomedPicture = A_PictureBox.Picture
End Sub


Private Sub PaintPicture8Size(ByVal vWidth As Single, ByVal vHeight As Single)
    A_PictureBox.Cls
    A_PictureBox.PaintPicture A_stdPicture, A_StartX, A_StartY, _
        vWidth, vHeight
        
    '��¼��ǰ��ͼƬ�ߴ�
    A_PictureSize.width = vWidth
    A_PictureSize.height = vHeight
    
    '���ź��ͼ��
    Set A_ZoomedPicture = A_PictureBox.Picture
End Sub

'ˢ�£����»�ͼ
Public Sub Refresh()
    
End Sub

'����ͼƬ.�ڵ�ǰͼƬ�ĳߴ������
'����ͼƬ��ԭ�ߴ�Ļ���
'��ʵ�ֵ������ݼ�
Public Sub ZoomPicture(ByVal vRate As Double)
    Dim dZoom As Double
    dZoom = vRate
    
    A_PictureSize.width = A_PictureSize.width * dZoom
    A_PictureSize.height = A_PictureSize.height * dZoom
    
    
    'A_PictureBox.PaintPicture A_PictureBox.Image, 0, 0, A_PictureSize.width, A_PictureSize.height
    PaintPicture8Size A_PictureSize.width, A_PictureSize.height
End Sub


Public Sub MoveUp(ByVal vScale As Double)
    A_StartY = A_StartY - vScale
    PaintPicture8Size A_PictureSize.width, A_PictureSize.height
End Sub

Public Sub MoveDown(ByVal vScale As Double)
    A_StartY = A_StartY + vScale
    PaintPicture8Size A_PictureSize.width, A_PictureSize.height
End Sub

Public Sub MoveLeft(ByVal vScale As Double)
    A_StartX = A_StartX - vScale
    PaintPicture8Size A_PictureSize.width, A_PictureSize.height
End Sub

Public Sub MoveRight(ByVal vScale As Double)
    A_StartX = A_StartX + vScale
    PaintPicture8Size A_PictureSize.width, A_PictureSize.height
End Sub

'˳ʱ����ת90��
Public Sub Clockwise()
    Dim W1 As Long, H1 As Long, B1() As Byte, Bs1 As Long, BytesW1 As Long, Ps1 As Long
    Dim W2 As Long, H2 As Long, B2() As Byte, Bs2 As Long, BytesW2 As Long, Ps2 As Long
    Dim S1 As Long, S2 As Long, X As Long, Y As Long, x1 As Long, y1 As Long
    Dim CenX1 As Long, CenY1 As Long, CenX2 As Long, CenY2 As Long
    Dim ToJ As Single
    
    ToJ = 90 / 180 * ctP180 '��ת�Ƕ�ת����
    
    '����һ����俴�ƿ��п��ޣ�ʵ�����������ã��ָ���תǰ�ؼ���ԭͼ��ʹ�С
    'A_PictureBox.Picture = A_PictureBox.Picture
    'A_PictureBox.Picture = PaintPicture8Size(A_PictureSize.width, A_PictureSize.height)
    'Set A_PictureBox.Picture = A_ZoomedPicture
    
    
    '��תǰͼ�����ݣ����ȣ��߶ȣ���ɫ���飬���ֽ�����ÿ���ֽ�����ÿ�����ֽ���
    GetBmpDat A_PictureBox, W1, H1, B1, Bs1, BytesW1, Ps1
    CenX1 = Int(W1 * 0.5): CenY1 = Int(H1 * 0.5)       '��תǰ��ͼ�����ĵ�
    
    '������ת��ؼ��ĸ߶ȺͿ��ȣ�ҪԤ�����ô����ͼƬ�� ScaleMode Ϊ 3�����أ�
    W2 = Abs(W1 * Cos(ToJ)) + Abs(H1 * Sin(ToJ))  '��ת��ͼ�����
    H2 = Abs(H1 * Cos(ToJ)) + Abs(W1 * Sin(ToJ))  '��ת��ͼ��߶�
    X = A_PictureBox.width - A_PictureBox.ScaleWidth      'ͼƬ��߿򣺿���
    Y = A_PictureBox.height - A_PictureBox.ScaleHeight    'ͼƬ��߿򣺸߶�
   
    
    'A_PictureBox.Move A_PictureBox.Left, A_PictureBox.Top, X + W2, Y + H2
    A_PictureBox.Cls
    A_PictureBox.Line (0, 0)-(W2, H2), &HFFFFFF, BF
    
    '��ת��ͼ�����ݣ����ȣ��߶ȣ���ɫ���飬���ֽ�����ÿ���ֽ�����ÿ�����ֽ���
    GetBmpDat A_PictureBox, W2, H2, B2, Bs2, BytesW2, Ps2
    CenX2 = Int(W2 * 0.5): CenY2 = Int(H2 * 0.5)       '��ת��ͼ�����ĵ�

    '��ʾ��Ϣ
    A_PictureBox.CurrentX = 5: A_PictureBox.CurrentY = 5
    A_PictureBox.Print "�����У����Ժ�..."
    
    W1 = W1 - 1: H1 = H1 - 1
    For X = 0 To W2 - 1
    For Y = 0 To H2 - 1
       Zhuan -ToJ, CenX2, CenY2, X, Y, x1, y1           '�� x1,y1 �����ת����
       x1 = x1 - CenX2 + CenX1: y1 = y1 - CenY2 + CenY1 'ת��Ϊ��תǰ������
       
       S2 = XYtoIndex(X, Y, BytesW2, Ps2)    '��ת�����ص������� B2 �е�����
       If x1 < 0 Or x1 > W1 Or y1 < 0 Or y1 > H1 Then
          B2(S2 + 2) = 255: B2(S2 + 1) = 255: B2(S2) = 255 '����ԭͼ��������Ϊ��ɫ
       Else
          S1 = XYtoIndex(x1, y1, BytesW1, Ps1)  '��תǰ�����ص������� B1 �е�����
          B2(S2 + 2) = B1(S1 + 2): B2(S2 + 1) = B1(S1 + 1): B2(S2) = B1(S1) '������
       End If
    Next
    Next
    
    
    SetBitmapBits A_PictureBox.Image, Bs2, B2(0) '�� Picture1 ��ͼ������Ϊ��ת��Ķ������� B2()
End Sub

'��ʱ����ת90��
Public Sub AntiClockwise()
    
End Sub


Private Sub GetBmpDat(Kj As Control, W As Long, h As Long, b() As Byte, Bs As Long, BytesW As Long, Ps As Long)
   '��ȡ�ؼ� Kj ��ͼ������
   Dim MapInf As BitMap

   GetObject Kj.Image, Len(MapInf), MapInf '�� MapInf �õ� Kj ��ͼ����Ϣ
   W = MapInf.bmWidth: h = MapInf.bmHeight 'ͼ����ȡ��߶ȣ����أ�
   BytesW = MapInf.bmWidthBytes            'ÿ��ռ���ֽ���
   Ps = BytesW \ W                         'ÿ�������ֽ���(һ��Ϊ4)
   Bs = W * h * Ps                         '���ֽ���=����*�߶�*ÿ�������ֽ�
   ReDim b(0 To Bs - 1)
   GetBitmapBits Kj.Image, Bs, b(0)        '�� Kj ͼ���������ص����ɫֵ����������� B()
End Sub

Private Sub Zhuan(ToJ As Single, x0 As Long, y0 As Long, ByVal X As Long, ByVal Y As Long, x1 As Long, y1 As Long)
   '���� x,y Χ�� x0,y0 ˳ʱ����ת ToJ ���ȣ��� x1,y1 ������ת���λ��
   'ע�⣺ҪԤ������Բ���� ctP180 = 4 * Atn(1)
    Dim s As Single, j As Single
   
    X = X - x0: Y = Y - y0
    s = Sqr(X ^ 2 + Y ^ 2)               'X,Y �� x0,y0 �ľ���
    If s = 0 Then j = 0 Else j = Y / s   '��ˮƽ�ߵļнǵ�����ֵ

    If Abs(j) >= 1 Then
       If j > 0 Then j = ctP180 * 0.5 Else j = -ctP180 * 0.5 '90 ��ʱ���������
    Else
       j = Atn(j / Sqr(-j * j + 1)) '��ˮƽ�ߵļн�
    End If
   
    If X < 0 Then j = -ctP180 - j
    x1 = x0 + s * Cos(j + ToJ): y1 = y0 + s * Sin(j + ToJ) '������ת���λ��
End Sub

Private Function XYtoIndex(X As Long, Y As Long, BytesW As Long, Ps As Long) As Long
   '����ͼ������ x,y ����ɫ�����е����λ�á�
   'BytesW��ÿ��ͼ��ռ���ֽ�����Ps��ÿ�����ص�ռ���ֽ���(һ��Ϊ4)
   XYtoIndex = Y * BytesW + X * Ps
End Function