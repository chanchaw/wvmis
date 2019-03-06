VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dmXY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private theX As Double
Private theY As Double

Public Property Get X() As Double
    X = theX
End Property

Public Property Let X(ByVal vData As Double)
    theX = vData
End Property


Public Property Get Y() As Double
    Y = theY
End Property

Public Property Let Y(ByVal vData As Double)
    theY = vData
End Property

