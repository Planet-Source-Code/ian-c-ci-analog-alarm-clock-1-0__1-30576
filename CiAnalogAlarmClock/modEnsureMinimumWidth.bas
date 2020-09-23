Attribute VB_Name = "modEnsureMinimumWidth"
'*import modIsMinimizedOrMaximized

Option Explicit

'/**
'* @project       TUM VB Praktikum WS 2001/2002
'* @spec          Aufgabe 4 - Analog Clock
'* @date          22.11.2001
'* @supervisor    D. Golias
'* @author        I. Carter, ian@computersciences.de (mail if you need support or want to comment)
'* @class         Module modEnsureMinimumWidth
'/**


'/**
'* @method        void ensureMinimumWidth
'* @description   ensures that a form always keeps a minimum width
'*                in percent of the screen.
'/**
Public Sub ensureMinimumWidth(frm As Form, dblPercentWidth As Double)
   '*recalculate dblPercentWidth
   dblPercentWidth = dblPercentWidth / 100
   '*IsMinimizedOrMaximizedException
   If isMinimizedOrMaximized(frm) Then Exit Sub
   '*CheckAndEventuallyReset
   If frm.Width < Screen.Width * dblPercentWidth Then frm.Width = Screen.Width * dblPercentWidth
End Sub
