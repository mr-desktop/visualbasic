Attribute VB_Name = "modMain"
Option Explicit

Type POINTAPI
        x As Long
        y As Long
End Type

Public Const ALTERNATE = 1

Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
