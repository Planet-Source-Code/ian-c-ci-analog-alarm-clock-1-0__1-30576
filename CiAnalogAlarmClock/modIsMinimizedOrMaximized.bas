Attribute VB_Name = "modIsMinimizedOrMaximized"
Option Explicit

'/**
'* @project       TUM VB Praktikum WS 2001/2002
'* @spec          Aufgabe 4 - Analog Clock
'* @date          22.11.2001
'* @supervisor    D. Golias
'* @author        I. Carter, ian@computersciences.de (mail if you need support or want to comment)
'* @class         Module modIsMinimizedOrMaximized
'/**


'/**
'* @method        boolean isMinimizedOrMaximized
'* @description   returns true if a form is min- or maximized
'/**
Public Function isMinimizedOrMaximized(frm As Form) As Boolean
       If frm.WindowState = vbMinimized Or frm.WindowState = vbMaximized Then isMinimizedOrMaximized = True: Exit Function  '*strict
       isMinimizedOrMaximized = False
End Function
