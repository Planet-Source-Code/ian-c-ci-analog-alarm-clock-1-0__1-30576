Attribute VB_Name = "modGetSmallerValue"
Option Explicit

'/**
'* @project       TUM VB Praktikum WS 2001/2002
'* @spec          Aufgabe 4 - Analog Clock
'* @date          22.11.2001
'* @supervisor    D. Golias
'* @author        I. Carter, ian@computersciences.de (mail if you need support or want to comment)
'* @class         Module modGetSmallerValue
'/**


'/**
'* @method        double getSmallerValue
'* @description   returns the smaller of two values
'/**
Public Function getSmallerValue(dblA As Double, dblB As Double) As Double
   getSmallerValue = dblA
   If dblB < dblA Then getSmallerValue = dblB
End Function
