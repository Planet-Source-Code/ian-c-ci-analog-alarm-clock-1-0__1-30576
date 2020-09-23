Attribute VB_Name = "modEnsureMinimumHeigth"
'*import modIsMinimizedOrMaximized

Option Explicit

'/**
'* @project       TUM VB Praktikum WS 2001/2002
'* @spec          Aufgabe 4 - Analog Clock
'* @date          22.11.2001
'* @supervisor    D. Golias
'* @author        I. Carter, ian@computersciences.de (mail if you need support or want to comment)
'* @class         Module modEnsureMinimumHeigth
'/**


'/**
'* @method        void ensureMinimumHeight
'* @description   ensures that a form always keeps a minimum height
'*                in percent of the screen.
'/**
Public Sub ensureMinimumHeight(frm As Form, dblPercentHeight As Double)
   '*recalculate dblPercentHeight
   dblPercentHeight = dblPercentHeight / 100
   '*IsMinimizedOrMaximizedException
   If isMinimizedOrMaximized(frm) Then Exit Sub
   '*CheckAndEventuallyReset
   If frm.Height < Screen.Height * dblPercentHeight Then frm.Height = Screen.Height * dblPercentHeight
End Sub



