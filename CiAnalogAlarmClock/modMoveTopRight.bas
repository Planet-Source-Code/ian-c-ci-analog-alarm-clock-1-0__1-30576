Attribute VB_Name = "modMoveTopRight"
Option Explicit

'/**
'* @project       TUM VB Praktikum WS 2001/2002
'* @spec          Aufgabe 4 - Analog Clock
'* @date          22.11.2001
'* @supervisor    D. Golias
'* @author        I. Carter, ian@computersciences.de (mail if you need support or want to comment)
'* @class         Module modMoveTopRight
'/**


'/**
'* @method        void moveTopRight
'* @description   moves a form to the top right corner of the screen
'/**
Public Sub moveTopRight(frm As Form)
   With frm
      .Top = 0
      .Left = Screen.Width - .Width
   End With
End Sub
