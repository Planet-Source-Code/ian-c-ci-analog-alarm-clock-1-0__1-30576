Attribute VB_Name = "modArc"
Option Explicit

'/**
'* @project       TUM VB Praktikum WS 2001/2002
'* @spec          Aufgabe 4 - Analog Clock
'* @date          22.11.2001
'* @supervisor    D. Golias
'* @author        I. Carter, ian@computersciences.de (mail if you need support or want to comment)
'* @class         Module modArc
'/**


'/**
'* @method        void arc
'* @description   converts an angle to radians or vice versa
'* @assumes       arc(a) = a * 2 * pi/360°
'/**
Public Function arc(dblAngle As Double) As Double
  arc = dblAngle * 1.74532925199433E-02
End Function
