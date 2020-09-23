Attribute VB_Name = "modResizeWithParent"
Option Explicit

'/**
'* @project       TUM VB Praktikum WS 2001/2002
'* @spec          Aufgabe 4 - Analog Clock
'* @date          22.11.2001
'* @supervisor    D. Golias
'* @author        I. Carter, ian@computersciences.de (mail if you need support or want to comment)
'* @class         Module modResizeWithParent
'/**


'/**
'* @method        void resizeWithParent
'* @description   resizes e.g. a control Objects with its parent
'/**
Public Sub resizeWithParent(obj As Object, Optional lngBorderSize As Long = 0)
   
   obj.Top = obj.Parent.ScaleTop + lngBorderSize / 2
   obj.Left = obj.Parent.ScaleLeft + lngBorderSize / 2
   obj.Width = obj.Parent.ScaleWidth - lngBorderSize
   obj.Height = obj.Parent.ScaleHeight - lngBorderSize
   
End Sub
